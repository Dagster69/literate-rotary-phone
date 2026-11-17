@echo off
rem Double-click to fetch the current MakeMKV beta key and register it.
rem Close MakeMKV first.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Get-MakeMKVKey.ps1"
echo.
pause
