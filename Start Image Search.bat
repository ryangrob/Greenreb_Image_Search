@echo off
REM Launches Image Search from source with no console window.
REM Use this while the .exe is older than the code - a rebuild (build.bat)
REM takes 10+ minutes, this starts the current version in seconds.
cd /d "%~dp0"
where pyw >nul 2>&1 && (start "" pyw "tkinter_app.py" & exit /b)
where pythonw >nul 2>&1 && (start "" pythonw "tkinter_app.py" & exit /b)
echo Could not find pyw or pythonw on this machine.
echo Falling back to a console launch - you can close this window afterwards.
py "tkinter_app.py"
