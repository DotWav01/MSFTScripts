@echo off
setlocal EnableDelayedExpansion

:: Production Robocopy Script for FP&A Directory
:: Specifically handles the Strategy - FP&A folder path

echo ==========================================
echo File Copy - Strategy FP^&A Directory
echo ==========================================
echo.

:: Configure paths - Update these as needed
:: Method: Escape the & with ^ character
set "SOURCE=Z:\Share Documents\Strategy - FP^&A"
set "DESTINATION=Y:\Share Documents\Strategy - FP^&A"

:: Alternative method if the above doesn't work:
:: Uncomment these lines and comment out the lines above
rem set "SOURCE=Z:\Share Documents\Strategy - FP&A"
rem set "DESTINATION=Y:\Share Documents\Strategy - FP&A"

echo Source: "!SOURCE!"
echo Destination: "!DESTINATION!"
echo.

:: Create timestamped log file
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YYYY=%dt:~0,4%" ^& set "MM=%dt:~4,2%" ^& set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%" ^& set "Min=%dt:~10,2%" ^& set "Sec=%dt:~12,2%"
set "timestamp=!YYYY!!MM!!DD!_!HH!!Min!!Sec!"

set "LOGFILE=!USERPROFILE!\Documents\FPandA_Copy_!timestamp!.log"
echo Log File: "!LOGFILE!"
echo.

:: Validate source path exists
echo Checking source path...
if exist "!SOURCE!" (
    echo [OK] Source path found: "!SOURCE!"
    
    :: Show some details about the source
    for /f %%i in ('dir "!SOURCE!" /A:D /B 2^>nul ^| find /c /v ""') do set "FOLDERS=%%i"
    for /f %%i in ('dir "!SOURCE!" /A:-D /B 2^>nul ^| find /c /v ""') do set "FILES=%%i"
    echo     Contains: !FOLDERS! folders, !FILES! files
) else (
    echo [ERROR] Source path not found: "!SOURCE!"
    echo.
    echo Please check:
    echo 1. Network drive Z: is connected
    echo 2. Path spelling is correct: "Strategy - FP&A" (with spaces and ampersand)
    echo 3. You have read access to the source folder
    echo.
    pause
    exit /b 1
)

echo.

:: Check destination parent directory
for %%i in ("!DESTINATION!") do set "DEST_PARENT=%%~dpi"
set "DEST_PARENT=!DEST_PARENT:~0,-1!"

echo Checking destination access...
if exist "!DEST_PARENT!" (
    echo [OK] Destination parent exists: "!DEST_PARENT!"
    
    :: Test write access
    set "TESTFILE=!DEST_PARENT!\writetest_!RANDOM!.tmp"
    echo test > "!TESTFILE!" 2>nul
    if exist "!TESTFILE!" (
        echo [OK] Write access confirmed
        del "!TESTFILE!" >nul 2>&1
    ) else (
        echo [WARNING] Cannot verify write access to: "!DEST_PARENT!"
    )
) else (
    echo [ERROR] Destination parent not accessible: "!DEST_PARENT!"
    pause
    exit /b 1
)

echo.
echo ==========================================
echo Starting Copy Operation
echo ==========================================
echo.
echo This will copy all files and folders from:
echo   "!SOURCE!"
echo To:
echo   "!DESTINATION!"
echo.
echo Press Ctrl+C to cancel, or any key to continue...
pause >nul

echo.
echo Executing Robocopy...
echo Command: robocopy "!SOURCE!" "!DESTINATION!" [options]
echo.

:: Execute the copy with proper path handling
robocopy "!SOURCE!" "!DESTINATION!" /E /COPY:DAT /DCOPY:DAT /R:3 /W:5 /MT:4 /NP /TEE /UNILOG:"!LOGFILE!" /XJD /XJF

:: Process results
set "EXITCODE=!ERRORLEVEL!"
echo.
echo ==========================================
echo Copy Operation Results
echo ==========================================
echo Exit Code: !EXITCODE!

if !EXITCODE! EQU 0 (
    echo Status: SUCCESS - No files needed copying ^(all files up to date^)
    set "STATUS=SUCCESS"
) else if !EXITCODE! EQU 1 (
    echo Status: SUCCESS - Files copied successfully
    set "STATUS=SUCCESS"
) else if !EXITCODE! EQU 2 (
    echo Status: SUCCESS - Some additional files were detected
    set "STATUS=SUCCESS"
) else if !EXITCODE! EQU 3 (
    echo Status: SUCCESS - Files copied and additional files detected
    set "STATUS=SUCCESS"
) else (
    echo Status: WARNING/ERROR - Check log file for details
    set "STATUS=CHECK_LOG"
)

echo.
echo Log file: "!LOGFILE!"

if "!STATUS!"=="CHECK_LOG" (
    echo.
    echo Opening log file for review...
    notepad "!LOGFILE!"
)

echo.
echo Operation completed at: !DATE! !TIME!
pause
