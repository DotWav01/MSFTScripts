@echo off
setlocal

:: Robocopy Script for SecurityDocs with Full Metadata Preservation
:: This script copies files while maintaining all metadata, versioning, and permissions

echo Starting SecurityDocs copy operation...
echo.

set SOURCE="\\sharesource\Securitydocs\StrategyFiles"
set DESTINATION="\\DestinationSource\SecurityDocs\StrategyFiles"
set LOGFILE="%~dp0SecurityDocs_Copy_%DATE:~-4,4%%DATE:~-10,2%%DATE:~-7,2%_%TIME:~0,2%%TIME:~3,2%%TIME:~6,2%.log"

:: Remove spaces from log filename
set LOGFILE=%LOGFILE: =%

echo Source: %SOURCE%
echo Destination: %DESTINATION%
echo Log File: %LOGFILE%
echo.

:: Execute Robocopy with full metadata preservation
robocopy %SOURCE% %DESTINATION% /E /COPY:DATSOU /DCOPY:DAT /R:3 /W:10 /MT:8 /NP /TEE /UNILOG:%LOGFILE% /XJD /XJF

:: Check exit code and display result
set EXITCODE=%ERRORLEVEL%
echo.
echo Robocopy completed with exit code: %EXITCODE%

if %EXITCODE% LEQ 3 (
    echo SUCCESS: Files copied successfully
    echo Log file: %LOGFILE%
) else if %EXITCODE% LEQ 7 (
    echo WARNING: Copy completed with some issues
    echo Check log file for details: %LOGFILE%
) else (
    echo ERROR: Copy operation failed
    echo Check log file for details: %LOGFILE%
)

echo.
pause
