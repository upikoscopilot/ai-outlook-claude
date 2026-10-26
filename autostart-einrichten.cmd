@echo off
setlocal
set "VBS=%~dp0start-addin-unsichtbar.vbs"
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ClaudeMailAddin /t REG_SZ /d "wscript.exe \"%VBS%\"" /f >nul
if errorlevel 1 (
  echo Autostart konnte nicht eingerichtet werden.
) else (
  echo Autostart eingerichtet: Der Claude-Mail-Server startet ab jetzt bei jeder Windows-Anmeldung unsichtbar.
  wscript.exe "%VBS%"
  echo Server wurde jetzt auch sofort gestartet.
)
pause
