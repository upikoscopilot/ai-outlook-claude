@echo off
setlocal
cd /d "%~dp0"
set "NODEDIR=%ProgramFiles%\nodejs"
if exist "%NODEDIR%\node.exe" goto havenode
where node >nul 2>nul && goto havenode_path
echo Node.js wurde nicht gefunden - wird jetzt ueber winget installiert ...
winget install -e --id OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements
if not exist "%NODEDIR%\node.exe" (
  echo.
  echo Node.js konnte nicht automatisch installiert werden.
  echo Bitte von https://nodejs.org die LTS-Version installieren und diese Datei erneut starten.
  pause
  exit /b 1
)
:havenode
set "PATH=%NODEDIR%;%PATH%"
:havenode_path
title Claude Mail Add-in
node serve.mjs
pause
