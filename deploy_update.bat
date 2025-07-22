REM deploy_update.bat - Deploy updates to user installations
@echo off
echo Deploying Primus Implant Report Generator Update...

REM Build the latest version first
call build_user_installer.bat

if %errorlevel% neq 0 (
    echo Build failed, cannot deploy update
    pause
    exit /b 1
)

REM Set update deployment path (you can change this to a network location)
set "DEPLOY_PATH=\\CDIMANQ30\Creoman-Active\CADCAM\Software\Primus Implant Report Generator\UserUpdates"

REM Create deployment directory
mkdir "%DEPLOY_PATH%" 2>nul

REM Copy new executable to deployment location
copy "dist\user\Primus Implant Report Generator.exe" "%DEPLOY_PATH%\Primus Implant Report Generator.exe"

if %errorlevel% neq 0 (
    echo Failed to copy update to deployment location
    pause
    exit /b 1
)

REM Copy data files
copy "icon.ico" "%DEPLOY_PATH%\"
copy "inosys_logo.png" "%DEPLOY_PATH%\"
copy "Primus Implant List - Primus Implant List.csv" "%DEPLOY_PATH%\"

REM Create version info file
(
echo Version: 1.0.3
echo Build Date: %date% %time%
echo Description: User-level installation update
) > "%DEPLOY_PATH%\version.txt"

echo.
echo Update deployed successfully to: %DEPLOY_PATH%
echo.
echo Users will be able to update automatically through the application.
echo.
pause

REM Optional: Create notification script for users
(
echo @echo off
echo echo Primus Implant Report Generator Update Available
echo echo.
echo echo A new version is available. To update:
echo echo 1. Open the Primus application
echo echo 2. Go to Help ^> Check for Updates
echo echo 3. Follow the update prompts
echo echo.
echo echo The update will install automatically without requiring administrator privileges.
echo echo.
echo pause
) > "%DEPLOY_PATH%\update_notification.bat"

echo Update notification script created: %DEPLOY_PATH%\update_notification.bat