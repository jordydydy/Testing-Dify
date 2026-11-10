@echo off
rem Removes auto-start and stops the running Test Runner.
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ProcurementTestRunner.lnk" 2>nul
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -match 'procurement_tester.py' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
echo Auto-start removed and the runner stopped.
pause
