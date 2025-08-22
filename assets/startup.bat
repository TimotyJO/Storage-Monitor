@echo off
REM Menambahkan registry untuk startup StorageMonitor.exe di folder C:\Program Files\Storage

set "appPath=C:\Program Files\Storage\StorageMonitor.exe"

reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v StorageMonitor /t REG_SZ /d "\"%appPath%\"" /f

echo Registry startup StorageMonitor sudah ditambahkan.
pause
