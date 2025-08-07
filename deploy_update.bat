REM deploy_update.bat - Updated for user-level installations
@echo off
echo Deploying Primus Implant Report Generator Update...

REM Build the latest version first
call build_installer.bat

if %errorlevel% neq 0 (
    echo Build failed, cannot deploy update
    pause
    exit /b 1
)

REM Set update deployment paths
REM Option 1: Network share (original method)
set "NETWORK_DEPLOY_PATH=\\CDIMANQ30\Creoman-Active\CADCAM\Software\Primus Implant Report Generator\UserUpdates"

REM Option 2: Local shared folder for testing
set "LOCAL_DEPLOY_PATH=C:\Shared\PrimusUpdates"

REM Choose deployment method
set "DEPLOY_PATH=%NETWORK_DEPLOY_PATH%"

echo Deploying to: %DEPLOY_PATH%

REM Create deployment directory
mkdir "%DEPLOY_PATH%" 2>nul

REM Copy new executable to deployment location
if exist "dist\Primus Implant Report Generator 1.0.5.exe" (
    copy "dist\Primus Implant Report Generator 1.0.5.exe" "%DEPLOY_PATH%\Primus Implant Report Generator.exe"
) else (
    echo Error: Built executable not found!
    pause
    exit /b 1
)

if %errorlevel% neq 0 (
    echo Failed to copy update to deployment location
    echo Check if network path is accessible: %DEPLOY_PATH%
    pause
    exit /b 1
)

REM Copy data files (users might need updated data)
copy "icon.ico" "%DEPLOY_PATH%\" 2>nul
copy "inosys_logo.png" "%DEPLOY_PATH%\" 2>nul
copy "Primus Implant List - Primus Implant List.csv" "%DEPLOY_PATH%\" 2>nul

REM Create version info file with detailed information
(
echo Version: 1.0.5
echo Build Date: %date% %time%
echo Description: User-level installation update
echo Executable: Primus Implant Report Generator.exe
echo Release Notes:
echo - Added case notes functionality
echo - Added print preview ^(View Report^)
echo - Added window size/position memory
echo - Improved user-level installation support
echo - Enhanced PDF report formatting
) > "%DEPLOY_PATH%\version.txt"

REM Create update manifest for version checking
(
echo {
echo   "version": "1.0.5",
echo   "build_date": "%date% %time%",
echo   "executable": "Primus Implant Report Generator.exe",
echo   "size": "%~z0",
echo   "required_files": [
echo     "Primus Implant Report Generator.exe",
echo     "Primus Implant List - Primus Implant List.csv",
echo     "inosys_logo.png",
echo     "icon.ico"
echo   ]
echo }
) > "%DEPLOY_PATH%\update_manifest.json"

REM Create deployment verification script
(
echo @echo off
echo echo Verifying Primus Implant Report Generator Update Deployment
echo echo.
echo echo Checking files...
echo if exist "Primus Implant Report Generator.exe" ^(
echo   echo [OK] Main executable found
echo ^) else ^(
echo   echo [ERROR] Main executable missing
echo ^)
echo.
echo if exist "Primus Implant List - Primus Implant List.csv" ^(
echo   echo [OK] CSV data file found
echo ^) else ^(
echo   echo [ERROR] CSV data file missing
echo ^)
echo.
echo if exist "inosys_logo.png" ^(
echo   echo [OK] Logo file found
echo ^) else ^(
echo   echo [ERROR] Logo file missing
echo ^)
echo.
echo echo Deployment verification complete.
echo pause
) > "%DEPLOY_PATH%\verify_deployment.bat"

REM Set appropriate permissions for network deployment
if "%DEPLOY_PATH%"=="%NETWORK_DEPLOY_PATH%" (
    echo Setting network permissions...
    REM Grant read access to domain users
    icacls "%DEPLOY_PATH%" /grant "Domain Users:R" /t >nul 2>&1
    icacls "%DEPLOY_PATH%\*" /grant "Domain Users:R" >nul 2>&1
)

echo.
echo Update deployed successfully to: %DEPLOY_PATH%
echo.
echo Files deployed:
dir "%DEPLOY_PATH%" /b

echo.
echo Users can now update through the application:
echo 1. Help ^> Check for Updates
echo 2. Updates install to user directories without admin rights
echo 3. No UAC prompts required for users
echo.

REM Create user notification script
(
echo @echo off
echo echo Primus Implant Report Generator Update Available - Version 1.0.5
echo echo.
echo echo New Features:
echo echo - Case Notes: Add special instructions to reports
echo echo - View Report: Preview before saving
echo echo - Window Memory: Remembers size and position
echo echo - Enhanced PDF formatting
echo echo.
echo echo To update:
echo echo 1. Open Primus Implant Report Generator
echo echo 2. Go to Help ^> Check for Updates
echo echo 3. Click Yes to install the update
echo echo.
echo echo The update installs automatically without administrator privileges.
echo echo.
echo pause
) > "%DEPLOY_PATH%\update_notification.bat"

echo Update notification script created: %DEPLOY_PATH%\update_notification.bat
echo.
echo Deployment Summary:
echo - Network path: %DEPLOY_PATH%
echo - Version: 1.0.5
echo - User installations will auto-detect this update
echo - No admin rights required for end users
echo.
pause