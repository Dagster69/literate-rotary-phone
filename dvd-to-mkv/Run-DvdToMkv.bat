@echo off
rem Double-click to open the DVD to MKV converter.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0DvdToMkv.ps1" %*
if errorlevel 1 pause
