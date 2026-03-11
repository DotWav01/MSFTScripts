@echo off
REM ============================================================
REM Launch-TSScanInstall.bat
REM ServiceUI Launcher for TSScan Server Interactive Installation
REM
REM Uses START /WAIT /MIN to immediately minimize this CMD window,
REM then launches PowerShell with -WindowStyle Hidden so no console
REM is visible to the user. The TSScan wizard GUI child process
REM still renders normally on the user desktop.
REM ============================================================

SET "SCRIPT_DIR=%~dp0"
IF "%SCRIPT_DIR:~-1%"=="\" SET "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

REM START /WAIT  = wait for PowerShell to finish before exiting (preserves exit code)
REM START /MIN   = minimise this CMD window immediately  
REM -WindowStyle Hidden = no PowerShell console window shown
START /WAIT /MIN "" powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT_DIR%\Install-TSScanServer.ps1"

EXIT /B %ERRORLEVEL%
