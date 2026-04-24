@echo off
setlocal
set "SCRIPT_DIR=%~dp0"

powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Uninstall-AJTools.ps1"
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
    echo AJ Tools uninstall failed with exit code %EXIT_CODE%.
) else (
    echo AJ Tools uninstall completed.
)
pause
exit /b %EXIT_CODE%
