\# Win+D Single Monitor



Windows utility that intercepts Win+D and applies "Show Desktop"

only to the monitor under the cursor.



\## Features

\- Low-level WinAPI hook (WH\_KEYBOARD\_LL)

\- Blocks only Win+D

\- Tray icon

\- Autostart support

\- Dark settings UI



\## How it works

Uses a low-level keyboard hook and manually minimizes windows

only on the selected monitor.



\## Requirements

\- Windows 10 / 11

\- Python 3.10+

\- pywin32

\- customtkinter

\- pystray

\- Pillow



\## Run

python main.py

Потом:

