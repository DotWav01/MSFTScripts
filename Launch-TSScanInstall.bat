@echo off
REM ============================================================
REM Launch-TSScanInstall.bat
REM ServiceUI Launcher for TSScan Server Interactive Installation
REM 
REM This batch file is called directly by ServiceUI.exe so that
REM the entire process tree (PowerShell + TSScan wizard) runs
REM in the user's session (Session 1+) rather than Session 0.
REM
REM ServiceUI projects this process and all its children into
REM the active user session, making the installer GUI visible.
REM ============================================================

REM Get the directory where this batch file lives (the package root)
SET "SCRIPT_DIR=%~dp0"

REM Remove trailing backslash
IF "%SCRIPT_DIR:~-1%"=="\" SET "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

REM Call the PowerShell install script
REM -NoProfile    : Faster startup, avoids profile interference
REM -NonInteractive: PowerShell itself stays non-interactive (the EXE GUI is interactive)
REM -ExecutionPolicy Bypass : Standard for Intune deployments
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Install-TSScanServer.ps1"

REM Capture and pass through the exit code from PowerShell
EXIT /B %ERRORLEVEL%
