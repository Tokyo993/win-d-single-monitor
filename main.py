# main.py
# Windows-only tray utility:
# - Intercepts Win + D via WinAPI low-level keyboard hook (WH_KEYBOARD_LL)
# - Blocks ONLY Win+D (does not break Win+R, Win+E, etc.)
# - Applies "show desktop" only for selected monitor (by cursor position)
# - Dark settings UI, tray icon, autostart shortcut in Startup

import os
import sys
import json
import threading
from pathlib import Path

import win32gui
import win32con
import win32api

import customtkinter as ctk

import pystray
from pystray import MenuItem as Item
from PIL import Image, ImageDraw

import ctypes
from ctypes import wintypes

# -----------------------------
# Config
# -----------------------------

APP_NAME = "Win+D Single Monitor"
CONFIG_PATH = Path(os.getenv("APPDATA", ".")) / "wind_fix_tray_config.json"


def load_config():
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return {
                "allowed_monitor": int(data.get("allowed_monitor", 0)),
                "autostart": bool(data.get("autostart", False)),
            }
        except Exception:
            pass
    return {"allowed_monitor": 0, "autostart": False}


def save_config(cfg: dict):
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# -----------------------------
# Autostart (Startup folder shortcut)
# -----------------------------

def startup_folder() -> Path:
    return Path(os.getenv("APPDATA", ".")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def shortcut_path() -> Path:
    return startup_folder() / f"{APP_NAME}.lnk"


def get_launch_target_and_args():
    exe = Path(sys.executable)
    script = Path(__file__).resolve()

    if getattr(sys, "frozen", False):
        return str(exe), ""

    target = str(exe)
    if exe.name.lower() == "python.exe":
        pythonw = exe.with_name("pythonw.exe")
        if pythonw.exists():
            target = str(pythonw)

    args = f"\"{script}\""
    return target, args


def set_autostart(enabled: bool):
    try:
        sf = startup_folder()
        sf.mkdir(parents=True, exist_ok=True)

        lnk = shortcut_path()
        if enabled:
            import win32com.client  # comes with pywin32
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(str(lnk))

            target, args = get_launch_target_and_args()
            shortcut.Targetpath = target
            shortcut.Arguments = args
            shortcut.WorkingDirectory = str(Path(__file__).resolve().parent)
            shortcut.IconLocation = target
            shortcut.save()
        else:
            if lnk.exists():
                lnk.unlink()
    except Exception as e:
        raise RuntimeError(f"Autostart error: {e}")


def is_autostart_enabled() -> bool:
    return shortcut_path().exists()


# -----------------------------
# Monitor helpers
# -----------------------------

def get_monitors():
    monitors = []
    for hmon, hdc, rect in win32api.EnumDisplayMonitors(None, None):
        info = win32api.GetMonitorInfo(hmon)
        monitors.append({
            "handle": hmon,
            "rect": rect,  # (l,t,r,b)
            "primary": info.get("Flags", 0) == 1
        })
    return monitors


def rect_info(rect):
    l, t, r, b = rect
    return {"x": l, "y": t, "w": r - l, "h": b - t}


def point_in_rect(x, y, rect):
    l, t, r, b = rect
    return l <= x < r and t <= y < b


def get_cursor_monitor_idx(monitors):
    x, y = win32gui.GetCursorPos()
    for i, m in enumerate(monitors):
        if point_in_rect(x, y, m["rect"]):
            return i
    return None


# -----------------------------
# Window helpers
# -----------------------------

def is_real_window(hwnd):
    if not win32gui.IsWindowVisible(hwnd):
        return False
    cls = win32gui.GetClassName(hwnd)
    if cls in ("Progman", "Shell_TrayWnd"):
        return False
    ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    if ex & win32con.WS_EX_TOOLWINDOW:
        return False
    return True


def get_window_monitor_idx(monitors, hwnd):
    try:
        l, t, r, b = win32gui.GetWindowRect(hwnd)
    except Exception:
        return None
    cx = (l + r) // 2
    cy = (t + b) // 2
    for i, m in enumerate(monitors):
        if point_in_rect(cx, cy, m["rect"]):
            return i
    return None


def enum_windows_on_monitor(monitors, idx):
    result = []

    def cb(hwnd, _):
        try:
            if not is_real_window(hwnd):
                return
            mi = get_window_monitor_idx(monitors, hwnd)
            if mi == idx:
                placement = win32gui.GetWindowPlacement(hwnd)
                if placement[1] != win32con.SW_SHOWMINIMIZED:
                    result.append(hwnd)
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    return result


# -----------------------------
# Core controller
# -----------------------------

class Controller:
    def __init__(self):
        self.cfg = load_config()
        self.monitors = get_monitors()
        self.allowed = int(self.cfg.get("allowed_monitor", 0))
        self.toggled = False
        self.minimized = []
        self._lock = threading.Lock()

    def refresh_monitors(self):
        with self._lock:
            self.monitors = get_monitors()

    def set_allowed(self, idx: int):
        with self._lock:
            self.allowed = max(0, idx)
            self.cfg["allowed_monitor"] = self.allowed
            save_config(self.cfg)

    def toggle_desktop_single_monitor(self):
        with self._lock:
            monitors = self.monitors
            allowed = self.allowed

        cursor_m = get_cursor_monitor_idx(monitors)

        # Important behavior:
        # We BLOCK Win+D always, to avoid global "show desktop" on both monitors.
        # But we only execute our minimize/restore if cursor is on allowed monitor.
        if cursor_m != allowed:
            return

        if not self.toggled:
            wins = enum_windows_on_monitor(monitors, allowed)
            self.minimized = wins
            for h in wins:
                try:
                    win32gui.ShowWindow(h, win32con.SW_MINIMIZE)
                except Exception:
                    pass
            self.toggled = True
        else:
            for h in reversed(self.minimized):
                try:
                    win32gui.ShowWindow(h, win32con.SW_RESTORE)
                except Exception:
                    pass
            self.minimized = []
            self.toggled = False


# -----------------------------
# WinAPI low-level keyboard hook
# -----------------------------

WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

VK_LWIN = 0x5B
VK_RWIN = 0x5C
VK_D = 0x44

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

ULONG_PTR = ctypes.c_uint64 if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_uint32
LRESULT  = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long

class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]

LowLevelProc = ctypes.WINFUNCTYPE(ctypes.c_longlong, wintypes.INT, wintypes.WPARAM, wintypes.LPARAM)

user32.SetWindowsHookExW.argtypes = (wintypes.INT, LowLevelProc, wintypes.HINSTANCE, wintypes.DWORD)
user32.SetWindowsHookExW.restype = wintypes.HHOOK

user32.CallNextHookEx.argtypes = (wintypes.HHOOK, wintypes.INT, wintypes.WPARAM, wintypes.LPARAM)
user32.CallNextHookEx.restype = LRESULT

user32.UnhookWindowsHookEx.argtypes = (wintypes.HHOOK,)
user32.UnhookWindowsHookEx.restype = wintypes.BOOL

user32.GetMessageW.argtypes = (ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT)
user32.GetMessageW.restype = wintypes.BOOL

user32.TranslateMessage.argtypes = (ctypes.POINTER(wintypes.MSG),)
user32.DispatchMessageW.argtypes = (ctypes.POINTER(wintypes.MSG),)

user32.PostThreadMessageW.argtypes = (wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
user32.PostThreadMessageW.restype = wintypes.BOOL

kernel32.GetModuleHandleW.argtypes = (wintypes.LPCWSTR,)
kernel32.GetModuleHandleW.restype = wintypes.HINSTANCE

kernel32.GetCurrentThreadId.restype = wintypes.DWORD

WM_QUIT = 0x0012


class WinDHook:
    """
    Blocks only Win+D. Everything else (Win+R etc.) passes normally.
    """
    def __init__(self, on_win_d):
        self.on_win_d = on_win_d
        self.hook = None
        self.thread = None
        self.thread_id = None

        self._win_down = False
        self._suppress_d_up = False

        self._proc = LowLevelProc(self._callback)

    def _callback(self, nCode, wParam, lParam):
        if nCode < 0:
            return user32.CallNextHookEx(self.hook, nCode, wParam, lParam)

        kbd = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
        vk = int(kbd.vkCode)

        is_down = wParam in (WM_KEYDOWN, WM_SYSKEYDOWN)
        is_up = wParam in (WM_KEYUP, WM_SYSKEYUP)

        # Track Win key state (pass through Win itself)
        if vk in (VK_LWIN, VK_RWIN):
            if is_down:
                self._win_down = True
            elif is_up:
                self._win_down = False
            return user32.CallNextHookEx(self.hook, nCode, wParam, lParam)

        # If Win is held and D is pressed -> trigger + suppress ONLY D
        if self._win_down and vk == VK_D and is_down:
            try:
                self.on_win_d()
            except Exception:
                pass
            self._suppress_d_up = True
            return 1  # swallow D down

        # Also swallow corresponding D up to avoid weirdness
        if self._suppress_d_up and vk == VK_D and is_up:
            self._suppress_d_up = False
            return 1

        return user32.CallNextHookEx(self.hook, nCode, wParam, lParam)

    def start(self):
        if self.thread and self.thread.is_alive():
            return

        def run():
            # install hook in this thread
            self.thread_id = kernel32.GetCurrentThreadId()
            hinst = kernel32.GetModuleHandleW(None)
            self.hook = user32.SetWindowsHookExW(WH_KEYBOARD_LL, self._proc, hinst, 0)
            if not self.hook:
                raise OSError("Failed to install keyboard hook")

            msg = wintypes.MSG()
            # message loop (required)
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))

            # cleanup
            if self.hook:
                user32.UnhookWindowsHookEx(self.hook)
                self.hook = None

        self.thread = threading.Thread(target=run, daemon=True)
        self.thread.start()

    def stop(self):
        # request the hook thread to quit its message loop
        try:
            if self.thread_id:
                user32.PostThreadMessageW(self.thread_id, WM_QUIT, 0, 0)
        except Exception:
            pass


# -----------------------------
# UI (Dark Settings)
# -----------------------------

class SettingsWindow:
    def __init__(self, ctrl: Controller, on_close_callback):
        self.ctrl = ctrl
        self.on_close_callback = on_close_callback

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.root = ctk.CTk()
        self.root.title(APP_NAME)
        self.root.geometry("640x520")
        self.root.resizable(False, False)

        self.root.protocol("WM_DELETE_WINDOW", self.close)

        title = ctk.CTkLabel(self.root, text=APP_NAME, font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(14, 6))

        subtitle = ctk.CTkLabel(
            self.root,
            text="Win + D перехватывается через WinAPI (Win+R и другие комбинации работают нормально).",
            font=ctk.CTkFont(size=12)
        )
        subtitle.pack(pady=(0, 10))

        self.mon_text = ctk.CTkTextbox(self.root, width=600, height=260)
        self.mon_text.pack(pady=(0, 10))
        self.mon_text.configure(state="normal")
        self.render_monitors()
        self.mon_text.configure(state="disabled")

        row = ctk.CTkFrame(self.root)
        row.pack(pady=10)

        ctk.CTkLabel(row, text="Разрешённый монитор:").pack(side="left", padx=(12, 8))

        self.monitor_var = ctk.StringVar(value=str(self.ctrl.allowed + 1))
        self.monitor_menu = ctk.CTkOptionMenu(
            row,
            values=self.available_monitor_values(),
            variable=self.monitor_var,
            command=self.on_monitor_change
        )
        self.monitor_menu.pack(side="left", padx=(0, 12))

        self.autostart_var = ctk.BooleanVar(value=is_autostart_enabled())
        self.autostart_chk = ctk.CTkCheckBox(
            self.root,
            text="Запускать вместе с Windows (Startup)",
            variable=self.autostart_var,
            command=self.on_autostart_toggle
        )
        self.autostart_chk.pack(pady=(8, 6))

        btns = ctk.CTkFrame(self.root)
        btns.pack(pady=12)

        ctk.CTkButton(btns, text="Refresh monitors", command=self.on_refresh).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="Close", command=self.close).pack(side="left", padx=8)

    def available_monitor_values(self):
        n = max(1, len(self.ctrl.monitors))
        return [str(i) for i in range(1, n + 1)]

    def render_monitors(self):
        self.ctrl.refresh_monitors()
        self.mon_text.delete("1.0", "end")

        for i, m in enumerate(self.ctrl.monitors):
            r = rect_info(m["rect"])
            primary = " (PRIMARY)" if m.get("primary") else ""
            self.mon_text.insert("end", f"Monitor {i+1}{primary}\n")
            self.mon_text.insert("end", f"  Resolution: {r['w']}x{r['h']}\n")
            self.mon_text.insert("end", f"  Position: x={r['x']} y={r['y']} → x={r['x']+r['w']} y={r['y']+r['h']}\n\n")

        self.mon_text.insert(
            "end",
            "Подсказка:\n- меньший x = монитор левее\n- отрицательный y = монитор выше основного\n"
        )

    def on_monitor_change(self, value: str):
        try:
            idx = int(value) - 1
            self.ctrl.set_allowed(idx)
        except Exception:
            pass

    def on_autostart_toggle(self):
        enabled = bool(self.autostart_var.get())
        try:
            set_autostart(enabled)
            self.ctrl.cfg["autostart"] = enabled
            save_config(self.ctrl.cfg)
        except Exception as e:
            self.autostart_var.set(is_autostart_enabled())
            self.show_error(str(e))

    def on_refresh(self):
        self.mon_text.configure(state="normal")
        self.render_monitors()
        self.mon_text.configure(state="disabled")

        self.monitor_menu.configure(values=self.available_monitor_values())
        current = self.ctrl.allowed + 1
        self.monitor_var.set(str(min(current, len(self.ctrl.monitors) or 1)))

    def show_error(self, msg: str):
        top = ctk.CTkToplevel(self.root)
        top.title("Error")
        top.geometry("520x160")
        top.resizable(False, False)

        ctk.CTkLabel(top, text="Ошибка", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(14, 6))
        box = ctk.CTkTextbox(top, width=480, height=70)
        box.pack()
        box.insert("1.0", msg)
        box.configure(state="disabled")
        ctk.CTkButton(top, text="OK", command=top.destroy).pack(pady=10)

    def run(self):
        self.root.mainloop()

    def close(self):
        try:
            self.root.destroy()
        except Exception:
            pass
        self.on_close_callback()


# -----------------------------
# Tray
# -----------------------------

def make_tray_icon_image():
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((10, 12, 54, 52), radius=10, outline=(255, 255, 255, 255), width=3)
    d.text((18, 24), "D", fill=(255, 255, 255, 255))
    return img


def run_app():
    ctrl = Controller()

    # WinAPI hook (blocks only Win+D)
    hook = WinDHook(on_win_d=ctrl.toggle_desktop_single_monitor)
    hook.start()

    settings_window_holder = {"open": False}

    def open_settings():
        if settings_window_holder["open"]:
            return
        settings_window_holder["open"] = True

        def on_close():
            settings_window_holder["open"] = False

        def ui_thread():
            win = SettingsWindow(ctrl, on_close)
            win.run()

        threading.Thread(target=ui_thread, daemon=True).start()

    def set_monitor_1():
        ctrl.set_allowed(0)

    def set_monitor_2():
        ctrl.set_allowed(1)

    def exit_app(icon, item):
        try:
            hook.stop()
        except Exception:
            pass
        icon.stop()

    icon = pystray.Icon(
        APP_NAME,
        make_tray_icon_image(),
        APP_NAME,
        menu=pystray.Menu(
            Item("Open settings", lambda: open_settings()),
            pystray.Menu.SEPARATOR,
            Item("Use Monitor 1", lambda: set_monitor_1()),
            Item("Use Monitor 2", lambda: set_monitor_2()),
            pystray.Menu.SEPARATOR,
            Item("Exit", exit_app),
        )
    )

    # apply autostart preference
    try:
        desired = bool(ctrl.cfg.get("autostart", False))
        if desired != is_autostart_enabled():
            set_autostart(desired)
    except Exception:
        pass

    icon.run()


if __name__ == "__main__":
    run_app()
