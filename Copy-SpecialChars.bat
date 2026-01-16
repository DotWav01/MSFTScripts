@echo off
setlocal EnableDelayedExpansion

:: Robocopy Script with Special Character Support
:: Handles paths with &, %, !, ^, and other special characters

echo ========================================
echo File Copy - Special Character Support
echo ========================================
echo.

:: Method 1: Use ^ to escape the & character
:: For paths with & symbol, escape it with ^
set "SOURCE=Z:\Share Documents\Strategy - FP^&A"
set "DESTINATION=Y:\Share Documents\Strategy - FP^&A"

:: Method 2: Alternative - Set paths in variables with quotes
:: This is another way to handle special characters
rem set "SOURCE=Z:\Share Documents\Strategy - FP&A"
rem set "DESTINATION=Y:\Share Documents\Strategy - FP&A"

:: Display the paths (using delayed expansion to show actual values)
echo Source: "!SOURCE!"
echo Destination: "!DESTINATION!"
echo.

:: Create log filename with timestamp
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YY=%dt:~2,2%" ^& set "YYYY=%dt:~0,4%" ^& set "MM=%dt:~4,2%" ^& set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%" ^& set "Min=%dt:~10,2%" ^& set "Sec=%dt:~12,2%"
set "timestamp=!YYYY!!MM!!DD!_!HH!!Min!!Sec!"

set "LOGFILE=!USERPROFILE!\Documents\FileCopy_!timestamp!.log"

echo Log File: "!LOGFILE!"
echo.

:: Verify source exists
if not exist "!SOURCE!" (
    echo ERROR: Source path does not exist: "!SOURCE!"
    echo.
    echo Troubleshooting tips:
    echo 1. Check if the path is correct
    echo 2. Make sure you have access to the network drive
    echo 3. Verify the folder name spelling including special characters
    echo.
    pause
    exit /b 1
)

echo Starting copy operation...
echo ** Preserving file integrity without permissions **
echo ** Handling paths with special characters **
echo.

:: Execute Robocopy with proper escaping
:: Using delayed expansion variables in quotes
robocopy "!SOURCE!" "!DESTINATION!" /E /COPY:DAT /DCOPY:DAT /R:3 /W:5 /MT:4 /NP /TEE /UNILOG:"!LOGFILE!" /XJD /XJF

:: Check exit code
set EXITCODE=!ERRORLEVEL!
echo.
echo ================
echo Operation Complete
echo ================
echo Exit Code: !EXITCODE!

if !EXITCODE! EQU 0 (
    echo Status: SUCCESS - No files needed copying
) else if !EXITCODE! EQU 1 (
    echo Status: SUCCESS - Files copied successfully
) else if !EXITCODE! EQU 2 (
    echo Status: SUCCESS - Some additional files found
) else if !EXITCODE! EQU 3 (
    echo Status: SUCCESS - Files copied and additional files found
) else if !EXITCODE! LEQ 7 (
    echo Status: WARNING - Copy completed with minor issues
    echo Check log file for details: "!LOGFILE!"
) else (
    echo Status: ERROR - Copy operation had significant issues
    echo Check log file for details: "!LOGFILE!"
)

echo.
echo Log file location: "!LOGFILE!"
echo.
echo ========================================
echo Special Character Handling Notes:
echo ========================================
echo For paths with ^& symbol: Use ^^ to escape: Strategy - FP^^&A
echo For paths with %% symbol: Use %%%% in batch files
echo For paths with !! symbol: Use delayed expansion
echo ========================================
echo.
pause
