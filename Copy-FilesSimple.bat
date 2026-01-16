@echo off
setlocal

:: Simple Robocopy for File Integrity (No Permissions)
:: Focuses on preserving timestamps, attributes, and versioning without security issues

echo ====================================
echo File Copy - Integrity Preservation
echo ====================================
echo.

:: Configuration - Update these paths as needed
set SOURCE="\\sharesource\Securitydocs\StrategyFiles"
set DESTINATION="\\DestinationSource\SecurityDocs\StrategyFiles"

:: Create log filename with timestamp
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YY=%dt:~2,2%" & set "YYYY=%dt:~0,4%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%" & set "Min=%dt:~10,2%" & set "Sec=%dt:~12,2%"
set "timestamp=%YYYY%%MM%%DD%_%HH%%Min%%Sec%"

set LOGFILE="%TEMP%\FileCopy_%timestamp%.log"

echo Source: %SOURCE%
echo Destination: %DESTINATION%
echo Log File: %LOGFILE%
echo.

echo Starting copy operation...
echo ** This preserves file integrity without permissions **
echo.

:: Execute Robocopy - Focus on Data, Attributes, Timestamps only
robocopy %SOURCE% %DESTINATION% /E /COPY:DAT /DCOPY:DAT /R:3 /W:5 /MT:4 /NP /TEE /UNILOG:%LOGFILE% /XJD /XJF

:: Check exit code
set EXITCODE=%ERRORLEVEL%
echo.
echo ================
echo Operation Complete
echo ================
echo Exit Code: %EXITCODE%

if %EXITCODE% EQU 0 (
    echo Status: SUCCESS - No files needed copying
) else if %EXITCODE% EQU 1 (
    echo Status: SUCCESS - Files copied successfully
) else if %EXITCODE% EQU 2 (
    echo Status: SUCCESS - Some additional files found
) else if %EXITCODE% EQU 3 (
    echo Status: SUCCESS - Files copied and additional files found
) else if %EXITCODE% LEQ 7 (
    echo Status: WARNING - Copy completed with minor issues
    echo Check log file for details: %LOGFILE%
) else (
    echo Status: ERROR - Copy operation had significant issues
    echo Check log file for details: %LOGFILE%
)

echo.
echo Log file: %LOGFILE%
echo.
pause
