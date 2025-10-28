@echo off
setlocal
set "SCRIPT_DIR=%~dp0"

rem Prefer PowerShell 7 (pwsh); fall back to Windows PowerShell if not found
where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
    pwsh -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT_DIR%script.ps1" -Interactive -UI
    exit /b %ERRORLEVEL%
)

where powershell >nul 2>nul
if %ERRORLEVEL%==0 (
    powershell -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT_DIR%script.ps1" -Interactive -UI
    exit /b %ERRORLEVEL%
)

echo Neither PowerShell 7 (pwsh) nor Windows PowerShell was found in PATH.
exit /b 1
