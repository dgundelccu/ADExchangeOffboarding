@echo off
setlocal

title Employee Offboarding

echo.
echo ==================================================
echo   Employee Offboarding
echo ==================================================
echo.

where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo ERROR: Windows PowerShell was not found.
    echo This script requires Windows PowerShell 5.1.
    echo.
    pause
    exit /b 1
)

if not exist "%~dp0Invoke-EmployeeOffboarding.ps1" (
    echo ERROR: Invoke-EmployeeOffboarding.ps1 was not found.
    echo Keep this BAT file in the same folder as the PS1 file.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Invoke-EmployeeOffboarding.ps1" %*
set "ScriptExitCode=%ERRORLEVEL%"

echo.
if "%ScriptExitCode%"=="2" (
    echo ATTENTION: The script completed with follow-up items.
    echo Review the CSV audit log before closing the termination ticket.
) else if not "%ScriptExitCode%"=="0" (
    echo The PowerShell script returned exit code %ScriptExitCode%.
)

pause
exit /b %ScriptExitCode%
