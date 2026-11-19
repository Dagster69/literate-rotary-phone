@echo off
rem Shows which original DVD folders can be deleted, then waits for you to type DELETE.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Delete-ConvertedOriginals.ps1" %*
echo.
pause
