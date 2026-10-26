@echo off
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ClaudeMailAddin /f >nul 2>nul
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /R /C:":3000 .*LISTENING"') do taskkill /PID %%p /F >nul 2>nul
echo Autostart entfernt und Server beendet.
pause
