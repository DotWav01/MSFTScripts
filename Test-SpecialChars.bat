@echo off
setlocal EnableDelayedExpansion

:: Special Character Testing Script
:: Use this to test different ways of handling special characters in paths

echo =======================================
echo Special Character Path Testing
echo =======================================
echo.

:: Test different methods for handling & in paths
echo Testing different methods for paths with ^& symbol:
echo.

:: Method 1: Escape with ^
set "PATH1=Z:\Share Documents\Strategy - FP^&A"
echo Method 1 (Escaped): "!PATH1!"

:: Method 2: Using quotes in assignment
set "PATH2=Z:\Share Documents\Strategy - FP&A"
echo Method 2 (Quoted): "!PATH2!"

:: Method 3: Double quotes
set PATH3="Z:\Share Documents\Strategy - FP&A"
echo Method 3 (Double Quoted): !PATH3!

echo.
echo =======================================
echo Testing Path Existence:
echo =======================================

:: Test which method works for your actual path
echo Testing if paths exist...
echo.

echo Testing Method 1:
if exist "!PATH1!" (
    echo [OK] Path exists using Method 1: "!PATH1!"
) else (
    echo [FAIL] Path does not exist using Method 1: "!PATH1!"
)

echo Testing Method 2:
if exist "!PATH2!" (
    echo [OK] Path exists using Method 2: "!PATH2!"
) else (
    echo [FAIL] Path does not exist using Method 2: "!PATH2!"
)

echo Testing Method 3:
if exist !PATH3! (
    echo [OK] Path exists using Method 3: !PATH3!
) else (
    echo [FAIL] Path does not exist using Method 3: !PATH3!
)

echo.
echo =======================================
echo Special Characters Reference:
echo =======================================
echo Character    How to Escape in Batch
echo ---------    ----------------------
echo ^&           Use ^^&  (caret before ampersand)
echo %%           Use %%%%  (double the percent)
echo ^^           Use ^^^^  (double the caret)
echo ^|           Use ^^|  (caret before pipe)
echo ^<           Use ^^<  (caret before less-than)
echo ^>           Use ^^>  (caret before greater-than)
echo !!           Use delayed expansion: EnableDelayedExpansion
echo.

echo =======================================
echo Recommended Robocopy Command Format:
echo =======================================
echo robocopy "!PATH1!" "destination" /E /COPY:DAT /DCOPY:DAT /R:3 /W:5 /MT:4 /NP /TEE /UNILOG:"logfile.log" /XJD /XJF
echo.

pause
