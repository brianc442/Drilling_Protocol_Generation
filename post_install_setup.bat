@echo off
REM post_install_setup.bat - Robust user directory setup
echo Setting up Primus Implant Report Generator...

REM Method 1: Try to get the user who initiated the installation
set "TARGET_USER="

REM Check if we're running under a specific user context
if defined ORIGINAL_USER (
    set "TARGET_USER=%ORIGINAL_USER%"
    echo Found original user from environment: %TARGET_USER%
    goto :SetupUser
)

REM Method 2: Use the user profile environment variable
if defined USERPROFILE (
    for %%F in ("%USERPROFILE%") do set "TARGET_USER=%%~nxF"
    echo Found user from USERPROFILE: %TARGET_USER%
    goto :SetupUser
)

REM Method 3: Try whoami command (most reliable)
for /f "tokens=2 delims=\" %%i in ('whoami 2^>nul') do (
    set "TARGET_USER=%%i"
    echo Found user from whoami: %TARGET_USER%
    goto :SetupUser
)

REM Method 4: Try query user command (fallback)
for /f "skip=1 tokens=1" %%i in ('query user 2^>nul ^| findstr /i "Active"') do (
    set "TARGET_USER=%%i"
    echo Found user from query user: %TARGET_USER%
    goto :SetupUser
)

REM Method 5: Use current logged-in user from environment
if defined USERNAME (
    set "TARGET_USER=%USERNAME%"
    echo Using USERNAME environment variable: %TARGET_USER%
    goto :SetupUser
)

REM If all methods fail, use a generic approach
echo Warning: Could not determine specific user, using current context
goto :SetupGeneric

:SetupUser
echo Setting up for user: %TARGET_USER%

REM Use LOCALAPPDATA if available, otherwise construct path
if defined LOCALAPPDATA (
    set "USER_DATA_DIR=%LOCALAPPDATA%\CreoDent\PrimusImplant"
) else (
    set "USER_DATA_DIR=C:\Users\%TARGET_USER%\AppData\Local\CreoDent\PrimusImplant"
)

REM Verify the user directory exists
if not exist "C:\Users\%TARGET_USER%" (
    echo Warning: User directory C:\Users\%TARGET_USER% does not exist
    echo Falling back to generic setup
    goto :SetupGeneric
)

echo Target directory: %USER_DATA_DIR%
goto :CreateDirectories

:SetupGeneric
echo Using generic setup approach...

REM Use LOCALAPPDATA or APPDATA as fallback
if defined LOCALAPPDATA (
    set "USER_DATA_DIR=%LOCALAPPDATA%\CreoDent\PrimusImplant"
) else if defined APPDATA (
    set "USER_DATA_DIR=%APPDATA%\..\Local\CreoDent\PrimusImplant"
) else (
    set "USER_DATA_DIR=%USERPROFILE%\AppData\Local\CreoDent\PrimusImplant"
)

echo Generic target directory: %USER_DATA_DIR%

:CreateDirectories
echo Creating directory structure...

REM Create the directory structure
mkdir "%USER_DATA_DIR%" 2>nul
mkdir "%USER_DATA_DIR%\updates" 2>nul
mkdir "%USER_DATA_DIR%\logs" 2>nul

REM Check if directories were created successfully
if not exist "%USER_DATA_DIR%" (
    echo Error: Failed to create user data directory
    echo Directory: %USER_DATA_DIR%
    goto :ErrorEnd
)

echo Successfully created directories

REM Copy data files to user directory
echo Copying data files...
copy "%~dp0*.csv" "%USER_DATA_DIR%\" >nul 2>&1
copy "%~dp0*.png" "%USER_DATA_DIR%\" >nul 2>&1
copy "%~dp0*.ico" "%USER_DATA_DIR%\" >nul 2>&1

REM Create registry entry for user data path (if we have permissions)
echo Setting registry entry...
reg add "HKLM\SOFTWARE\CreoDent\PrimusImplant" /v "UserDataPath" /t REG_SZ /d "%%LOCALAPPDATA%%\CreoDent\PrimusImplant" /f >nul 2>&1

REM Create desktop shortcut
echo Creating desktop shortcut...
call :CreateDesktopShortcut

REM Create Start Menu shortcut
echo Creating Start Menu shortcut...
call :CreateStartMenuShortcut

REM Create uninstaller
echo Creating uninstaller...
call :CreateUninstaller

echo.
echo Setup completed successfully!
echo Application data directory: %USER_DATA_DIR%
echo.
goto :End

:CreateDesktopShortcut
REM Get the installation directory
set "INSTALL_DIR=%~dp0"
set "EXE_PATH=%INSTALL_DIR%Primus Implant Report Generator 1.0.4.exe"

REM Create desktop shortcut using PowerShell
if defined USERPROFILE (
    set "DESKTOP_PATH=%USERPROFILE%\Desktop"
) else (
    set "DESKTOP_PATH=%PUBLIC%\Desktop"
)

set "SHORTCUT_PATH=%DESKTOP_PATH%\Primus Implant Report Generator.lnk"

powershell -Command "try { $WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%SHORTCUT_PATH%'); $Shortcut.TargetPath = '%EXE_PATH%'; $Shortcut.WorkingDirectory = '%INSTALL_DIR%'; $Shortcut.IconLocation = '%EXE_PATH%'; $Shortcut.Description = 'Primus Implant Report Generator'; $Shortcut.Save(); Write-Host 'Desktop shortcut created' } catch { Write-Host 'Desktop shortcut failed:' $_.Exception.Message }" 2>nul

exit /b

:CreateStartMenuShortcut
REM Create Start Menu shortcut
if defined APPDATA (
    set "START_MENU_PATH=%APPDATA%\Microsoft\Windows\Start Menu\Programs\CreoDent"
) else (
    set "START_MENU_PATH=%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\CreoDent"
)

mkdir "%START_MENU_PATH%" 2>nul
set "START_SHORTCUT_PATH=%START_MENU_PATH%\Primus Implant Report Generator.lnk"

powershell -Command "try { $WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%START_SHORTCUT_PATH%'); $Shortcut.TargetPath = '%EXE_PATH%'; $Shortcut.WorkingDirectory = '%INSTALL_DIR%'; $Shortcut.IconLocation = '%EXE_PATH%'; $Shortcut.Description = 'Primus Implant Report Generator'; $Shortcut.Save(); Write-Host 'Start Menu shortcut created' } catch { Write-Host 'Start Menu shortcut failed:' $_.Exception.Message }" 2>nul

exit /b

:CreateUninstaller
REM Create uninstaller script
(
echo @echo off
echo echo Uninstalling Primus Implant Report Generator...
echo echo.
echo.
echo REM Stop any running instances
echo taskkill /f /im "Primus Implant Report Generator 1.0.4.exe" 2^>nul
echo timeout /t 2 /nobreak ^>nul
echo.
echo REM Remove desktop shortcut
echo del "%USERPROFILE%\Desktop\Primus Implant Report Generator.lnk" 2^>nul
echo del "%PUBLIC%\Desktop\Primus Implant Report Generator.lnk" 2^>nul
echo.
echo REM Remove Start Menu shortcuts
echo rmdir /s /q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\CreoDent" 2^>nul
echo rmdir /s /q "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\CreoDent" 2^>nul
echo.
echo REM Remove user data directory ^(optional^)
echo echo.
echo echo Do you want to remove user data and settings? ^(Y/N^)
echo set /p "REMOVE_DATA="
echo if /i "%%REMOVE_DATA%%"=="Y" ^(
echo     echo Removing user data...
echo     rmdir /s /q "%USER_DATA_DIR%" 2^>nul
echo ^)
echo.
echo REM Remove registry entries
echo reg delete "HKLM\SOFTWARE\CreoDent\PrimusImplant" /f 2^>nul
echo.
echo echo Uninstallation complete.
echo pause
) > "%USER_DATA_DIR%\uninstall.bat"

exit /b

:ErrorEnd
echo.
echo Installation encountered errors. Please check:
echo 1. Run as administrator if needed
echo 2. Ensure user directory is accessible
echo 3. Check Windows permissions
echo.
pause
exit /b 1

:End
echo Installation script completed.
timeout /t 5 /nobreak >nul