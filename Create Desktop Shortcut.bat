@echo off
REM Run this ONCE. It puts "Topshot" on your Desktop and in the Start menu,
REM launching the app directly - no console window, no typing. It also
REM removes the old "Image Search" shortcuts, which still show the old icon.
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
  "$places=@([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'));" ^
  "foreach ($d in $places) {" ^
  "  Remove-Item -LiteralPath ($d+'\Image Search.lnk') -ErrorAction SilentlyContinue;" ^
  "  $s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\Topshot.lnk');" ^
  "  $s.TargetPath='%PYW%';" ^
  "  $s.Arguments='\"%~dp0launcher.py\"';" ^
  "  $s.WorkingDirectory='%~dp0.';" ^
  "  $s.IconLocation='%~dp0assets\app_icon.ico,0';" ^
  "  $s.Description='Topshot - search the marketing photo library';" ^
  "  $s.Save();" ^
  "  Write-Host ('Created: ' + $d + '\Topshot.lnk')" ^
  "}"

REM Windows keeps old icons in a cache; this asks it to refresh.
ie4uinit.exe -show >nul 2>&1

echo.
echo Done. "Topshot" is on your Desktop and in the Start menu.
echo Right-click it and choose "Pin to taskbar" if you want it there too.
echo.
pause
