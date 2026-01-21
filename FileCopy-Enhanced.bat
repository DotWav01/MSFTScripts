@echo off
setlocal EnableDelayedExpansion
:: Enhanced Robocopy for File Integrity - Handles Special Characters
:: Fixed version with proper escaping for special characters like &, %, etc.
echo ====================================
echo File Copy - Integrity Preservation  
echo ====================================
echo.

:: Configuration - Update these paths as needed
:: Note: Use quotes around paths with spaces or special characters
set "SOURCE=Z:\Shared Documents\CopyTest"
set "DESTINATION=Y:\Shared Documents\New folder"

:: Alternative examples for paths with special characters:
:: set "SOURCE=Z:\Strategy - FP&A\Reports"
:: set "DESTINATION=C:\Backup\Strategy - FP&A"

:: Create log filename with timestamp
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YY=%dt:~2,2%" & set "YYYY=%dt:~0,4%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%" & set "Min=%dt:~10,2%" & set "Sec=%dt:~12,2%"
set "timestamp=%YYYY%%MM%%DD%_%HH%%Min%%Sec%"
set "LOGFILE=%USERPROFILE%\Documents\FileCopy_!timestamp!.log"

echo Source: "!SOURCE!"
echo Destination: "!DESTINATION!"
echo Log File: "!LOGFILE!"
echo.

:: Verify source exists
if not exist "!SOURCE!" (
    echo ERROR: Source path does not exist: "!SOURCE!"
    echo Please check the source path and try again.
    pause
    exit /b 1
)

:: Create destination directory if it doesn't exist
if not exist "!DESTINATION!" (
    echo Creating destination directory...
    mkdir "!DESTINATION!" 2>nul
    if !errorlevel! neq 0 (
        echo ERROR: Could not create destination directory: "!DESTINATION!"
        pause
        exit /b 1
    )
)

echo Starting copy operation...
echo ** This preserves file integrity without permissions **
echo ** Handles special characters in folder names **
echo.

:: Execute Robocopy with proper variable expansion and quoting
:: Using delayed expansion to handle special characters properly
robocopy "!SOURCE!" "!DESTINATION!" /E /COPY:DAT /DCOPY:DAT /R:3 /W:5 /MT:4 /NP /TEE /UNILOG:"!LOGFILE!" /XJD /XJF

:: Check exit code
set EXITCODE=!ERRORLEVEL!
echo.
echo ================
echo Operation Complete  
echo ================
echo Exit Code: !EXITCODE!

:: More detailed exit code interpretation
if !EXITCODE! EQU 0 (
    echo Status: SUCCESS - No files needed copying
) else if !EXITCODE! EQU 1 (
    echo Status: SUCCESS - Files copied successfully  
) else if !EXITCODE! EQU 2 (
    echo Status: SUCCESS - Some additional files found
) else if !EXITCODE! EQU 3 (
    echo Status: SUCCESS - Files copied and additional files found
) else if !EXITCODE! EQU 4 (
    echo Status: WARNING - Some mismatched files or directories detected
) else if !EXITCODE! EQU 5 (
    echo Status: SUCCESS - Some files were copied, some were mismatched
) else if !EXITCODE! EQU 6 (
    echo Status: WARNING - Additional files and mismatched files exist
) else if !EXITCODE! EQU 7 (
    echo Status: SUCCESS - Files copied, additional and mismatched files exist
) else if !EXITCODE! EQU 8 (
    echo Status: ERROR - Some files or directories could not be copied
) else if !EXITCODE! GTR 8 (
    echo Status: ERROR - Serious error occurred during copy operation
    echo Check log file for details: "!LOGFILE!"
) else (
    echo Status: UNKNOWN - Unexpected exit code
)

echo.
echo Log file location: "!LOGFILE!"
echo.

:: Display summary from log file if it exists
if exist "!LOGFILE!" (
    echo === Copy Summary ===
    findstr /C:"Files :" /C:"Dirs :" /C:"Bytes :" /C:"Times :" "!LOGFILE!" 2>nul
    echo.
)

echo Press any key to exit...
pause >nul
