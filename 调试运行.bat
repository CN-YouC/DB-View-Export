@echo off
title Export Debug
setlocal
cd /d "%~dp0"
echo === DB View Export (debug) ===
echo.
set "EXE="
for %%F in ("%~dp0DB*.exe") do set "EXE=%%~fF"
if not defined EXE (
  echo [ERROR] DB*.exe not found in this folder.
  pause
  exit /b 1
)
echo Using: %EXE%
echo.
"%EXE%" %*
echo.
echo [exit code] %errorlevel%
echo.
pause
