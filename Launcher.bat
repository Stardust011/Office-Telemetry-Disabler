@echo off
setlocal enabledelayedexpansion

set "PS5_PATH=%systemdrive%\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS7_PATH=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "PS7_PREVIEW_PATH=%ProgramFiles%\PowerShell\7-preview\pwsh.exe"

set "ELEVATION_PS="
if exist "%PS7_PREVIEW_PATH%" (
    set "ELEVATION_PS=%PS7_PREVIEW_PATH%"
) else if exist "%PS7_PATH%" (
    set "ELEVATION_PS=%PS7_PATH%"
) else if exist "%PS5_PATH%" (
    set "ELEVATION_PS=%PS5_PATH%"
) else (
    echo [ERROR] No PowerShell found for elevation!
    pause
    exit /b 1
)

if "%1"=="admin" goto :AdminMode

echo [INFO] Launch with admin rights...
"%ELEVATION_PS%" -Command "Start-Process cmd -ArgumentList '/c \"%~f0\" admin' -Verb RunAs"
exit /B

:AdminMode
pushd "%CD%"
CD /D "%~dp0"

title Office Privacy and Telemetry Disabler Launcher
set "SCRIPT_DIR=%~dp0"
set "PS_EXE="
set "PS_VERSION="

if exist "%PS7_PREVIEW_PATH%" (
    set "PS_EXE=%PS7_PREVIEW_PATH%"
    set "PS_VERSION=PowerShell 7 Preview"
) else if exist "%PS7_PATH%" (
    set "PS_EXE=%PS7_PATH%"
    set "PS_VERSION=PowerShell 7"
) else if exist "%PS5_PATH%" (
    set "PS_EXE=%PS5_PATH%"
    set "PS_VERSION=PowerShell 5"
) else (
    echo [ERROR] No compatible PowerShell version found!
    pause
    exit /b 1
)

set "PS_SCRIPT=%SCRIPT_DIR%script\office_privacy_telemetry_disabler.ps1"
if not exist "%PS_SCRIPT%" (
    set "PS_SCRIPT=%SCRIPT_DIR%office_privacy_telemetry_disabler.ps1"
)
if not exist "%PS_SCRIPT%" (
    echo [ERROR] office_privacy_telemetry_disabler.ps1 not found!
    echo Place it in the same folder as this launcher or in the script subfolder.
    pause
    exit /b 1
)

echo.
echo ====================================================
echo    Office Privacy and Telemetry Disabler Launcher
echo                Office 16.0 baseline
echo ====================================================
echo.
echo System Information:
echo  - PowerShell: %PS_VERSION%
echo  - Script: !PS_SCRIPT!
echo.
echo Default behavior:
echo  - Apply Office 16.0 telemetry/privacy baseline
echo  - Update-disabling options remain OFF unless you opt in
echo.

:confirmation
set /p "CONFIRM=Do you want to continue? (Y/N): "
if /i "!CONFIRM!"=="y" goto :proceed
if /i "!CONFIRM!"=="yes" goto :proceed
if /i "!CONFIRM!"=="n" goto :cancel
if /i "!CONFIRM!"=="no" goto :cancel

echo Invalid input. Please enter Y or N.
goto :confirmation

:cancel
echo.
echo Operation cancelled by user.
pause
exit /b 0

:proceed
cls
echo.
echo [INFO] Launching Office Privacy Disabler...
echo [INFO] PowerShell: %PS_VERSION%
echo.

cd /d "%SCRIPT_DIR%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((Get-Content -Raw -LiteralPath '!PS_SCRIPT!')))"

if %errorLevel% neq 0 (
    echo.
    echo [ERROR] Script encountered errors. Exit code: %errorLevel%
)

echo.
echo Press any key to exit...
pause >nul
exit /b 0
