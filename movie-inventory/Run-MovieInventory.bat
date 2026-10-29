@echo off
rem Double-click to make a spreadsheet of every movie on the NAS.
rem Read-only: nothing on the NAS is changed.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0MovieInventory.ps1" %*
echo.
pause
