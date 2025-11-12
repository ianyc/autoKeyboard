1. open venv (PowerShell):
    > Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    > .\.venv\Scripts\Activate.ps1
2. run python:
    > python .\quickpaste_win.py

3. build: 
    > pyinstaller --noconsole --onefile --add-data "icon.ico;." quickpaste_win.py