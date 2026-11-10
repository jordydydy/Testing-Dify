@echo off
rem Starts the Procurement Test Runner automatically every time you log in to Windows.
set "LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ProcurementTestRunner.lnk"
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($env:LNK); $s.TargetPath='wscript.exe'; $s.Arguments='\"%~dp0start_hidden.vbs\"'; $s.WorkingDirectory='%~dp0'; $s.WindowStyle=7; $s.Save()"
echo Auto-start installed: %LNK%
echo Starting it now...
wscript.exe "%~dp0start_hidden.vbs"
echo Open http://127.0.0.1:8765
pause
