@echo off
echo Building Primus Dental Implant Report Generator...
echo.

REM Clean previous builds
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

echo Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building executable...
pyinstaller primus_installer.spec

echo.
echo Build complete! Check the dist folder for your executable.
echo.

REM Optional: Create a simple installer folder structure
if exist "dist\Primus Dental Implant Report Generator.exe" (
    echo Creating installer package...
    mkdir "installer_package"
    copy "dist\Primus Dental Implant Report Generator.exe" "installer_package\"
    copy "README.md" "installer_package\" 2>nul
    copy "LICENSE.txt" "installer_package\" 2>nul
    echo.
    echo Installer package created in 'installer_package' folder!
)

pause