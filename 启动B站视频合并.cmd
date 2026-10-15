@echo off
setlocal EnableExtensions
set "PS1="
for %%F in ("%~dp0*.ps1") do set "PS1=%%~fF"
if not defined PS1 (
  echo [ERROR] The .ps1 script was not found next to this .cmd file.
  echo         Please keep both files in the same folder.
  echo.
  pause
  exit /b 1
)
start "" /min powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%PS1%"
exit /b 0