@echo off
chcp 65001 >nul
title DB View Export
setlocal
cd /d "%~dp0"
set "EXE="
for %%F in ("%~dp0DB*.exe") do set "EXE=%%~fF"
if not defined EXE (
  echo [ERROR] DB*.exe not found in this folder.
  echo         ^(maybe blocked or deleted by antivirus^)
  pause
  exit /b 1
)
"%EXE%" %*
echo.
pause
