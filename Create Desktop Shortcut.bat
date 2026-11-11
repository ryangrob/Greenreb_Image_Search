@echo off
REM Run this ONCE. It puts an "Image Search" icon on your Desktop that
REM launches the app directly - no console window, no typing.
setlocal
cd /d "%~dp0"

set "PYW="
for /f "delims=" %%P in ('where pyw 2^>nul') do set "PYW=%%P"
if not defined PYW for /f "delims=" %%P in ('where pythonw 2^>nul') do set "PYW=%%P"
if not defined PYW (
    echo.
    echo Could not find pyw.exe or pythonw.exe on this computer.
    echo You can still start the app with "Start Image Search.bat".
    echo.
    pause
    exit /b 1
)

echo Using %PYW%
powershell -NoProfile -Command ^
  "$d=[Environment]::GetFolderPath('Desktop');" ^
  "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\Image Search.lnk');" ^
  "$s.TargetPath='%PYW%';" ^
  "$s.Arguments='\"%~dp0tkinter_app.py\"';" ^
  "$s.WorkingDirectory='%~dp0.';" ^
  "$s.IconLocation='%~dp0assets\app_icon.ico';" ^
  "$s.Description='Greenreb Image Search';" ^
  "$s.Save();" ^
  "Write-Host ('Created: ' + $d + '\Image Search.lnk')"

echo.
echo Done. "Image Search" is on your Desktop.
echo Right-click it and choose "Pin to taskbar" if you want it there too.
echo.
pause
