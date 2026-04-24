@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PACKAGE_ZIP="

if not "%~1"=="" (
    set "PACKAGE_ZIP=%~f1"
)

if not "%PACKAGE_ZIP%"=="" (
    if not exist "%PACKAGE_ZIP%" (
        echo The specified AJ Tools package was not found:
        echo %PACKAGE_ZIP%
        pause
        exit /b 1
    )
    goto found_zip
)

for %%F in ("%SCRIPT_DIR%..\artifacts\runtime-package\Release\AJTools-*.zip") do (
    set "PACKAGE_ZIP=%%~fF"
    goto found_zip
)

for %%F in ("%SCRIPT_DIR%AJTools-*.zip") do (
    set "PACKAGE_ZIP=%%~fF"
    goto found_zip
)

echo Could not find an AJTools-*.zip package next to this installer.
pause
exit /b 1

:found_zip
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Install-AJTools.ps1" -PackageZip "%PACKAGE_ZIP%"
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
    echo AJ Tools install failed with exit code %EXIT_CODE%.
) else (
    echo AJ Tools install completed.
)
pause
exit /b %EXIT_CODE%
