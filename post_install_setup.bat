@echo off
setlocal enabledelayedexpansion
REM post_install_setup.bat - Install for actual user when running as admin
echo Setting up Primus Implant Report Generator...

set "TARGET_USER="
set "USER_FOUND=0"

echo Detecting actual logged-in user...

REM Method 1: Check environment for original user (if set by installer)
if defined ORIGINAL_USER (
    set "TARGET_USER=%ORIGINAL_USER%"
    set "USER_FOUND=1"
    echo Found original user from environment: %ORIGINAL_USER%
    goto :CheckUser
)

REM Method 3: Check for users with explorer.exe running
echo Checking for active explorer processes...
for /f "skip=1 tokens=1,2" %%a in ('wmic process where "name='explorer.exe'" get ProcessId^,GetOwner /format:csv 2^>nul') do (
    if not "%%b"=="" (
        for /f "tokens=2 delims=\" %%c in ("%%b") do (
            set "TEMP_USER=%%c"
            if not "!TEMP_USER!"=="" if not "!TEMP_USER!"=="SYSTEM" (
                if exist "C:\Users\!TEMP_USER!" (
                    set "TARGET_USER=!TEMP_USER!"
                    set "USER_FOUND=1"
                    echo Found user from explorer process: !TEMP_USER!
                    goto :CheckUser
                )
            )
        )
    )
)

REM Method 4: Find the most recently modified user profile (excluding system accounts)
echo Checking for most recently used profile...
set "NEWEST_TIME="
set "NEWEST_USER="
for /d %%a in (C:\Users\*) do (
    set "USERNAME=%%~nxa"
    if not "!USERNAME!"=="All Users" if not "!USERNAME!"=="Default" if not "!USERNAME!"=="Default User" if not "!USERNAME!"=="Public" (
        if exist "%%a\NTUSER.DAT" (
            for %%b in ("%%a\AppData\Local") do (
                if exist "%%b" (
                    set "PROFILE_TIME=%%~tb"
                    if "!NEWEST_TIME!"=="" set "NEWEST_TIME=!PROFILE_TIME!"
                    if "!PROFILE_TIME!" gtr "!NEWEST_TIME!" (
                        set "NEWEST_TIME=!PROFILE_TIME!"
                        set "NEWEST_USER=!USERNAME!"
                    )
                )
            )
        )
    )
)
if not "!NEWEST_USER!"=="" (
    set "TARGET_USER=!NEWEST_USER!"
    set "USER_FOUND=1"
    echo Found most recently used profile: !NEWEST_USER!
    goto :CheckUser
)

REM Method 5: If all else fails, show available users and let user choose
:PromptUser
echo.
echo Could not automatically detect the target user.
echo Available users on this system:
set "USER_COUNT=0"
for /d %%a in (C:\Users\*) do (
    set "USERNAME=%%~nxa"
    if not "!USERNAME!"=="All Users" if not "!USERNAME!"=="Default" if not "!USERNAME!"=="Default User" if not "!USERNAME!"=="Public" (
        if exist "%%a\NTUSER.DAT" (
            set /a USER_COUNT+=1
            echo   !USER_COUNT!. !USERNAME!
            set "USER_!USER_COUNT!=!USERNAME!"
        )
    )
)

if %USER_COUNT%==0 (
    echo No valid user profiles found!
    goto :ErrorEnd
)

if %USER_COUNT%==1 (
    set "TARGET_USER=!USER_1!"
    set "USER_FOUND=1"
    echo Only one user found, selecting: !USER_1!
    goto :CheckUser
)

echo.
echo Please select the user for whom to install this application:
echo Enter the number (1-!USER_COUNT!) or type the username directly:
set /p "USER_CHOICE=Your choice: "

REM Check if it's a number
echo !USER_CHOICE! | findstr /r "^[1-9][0-9]*$" >nul
if %errorlevel%==0 (
    if !USER_CHOICE! gtr 0 if !USER_CHOICE! leq !USER_COUNT! (
        call set "TARGET_USER=%%USER_!USER_CHOICE!%%"
        set "USER_FOUND=1"
        echo Selected user by number: !TARGET_USER!
        goto :CheckUser
    )
)

REM Treat as direct username input
set "TARGET_USER=!USER_CHOICE!"
set "USER_FOUND=1"
echo Entered username: !TARGET_USER!

:CheckUser
if "%USER_FOUND%"=="0" (
    echo Error: Could not determine target user for installation.
    goto :ErrorEnd
)

echo.
echo Target user: %TARGET_USER%

REM Validate that the user directory exists
set "USER_HOME_DIR=C:\Users\%TARGET_USER%"
if not exist "%USER_HOME_DIR%" (
    echo Error: User directory %USER_HOME_DIR% does not exist.
    echo Please verify the username is correct.
    echo.
    goto :PromptUser
)

echo Validated user directory: %USER_HOME_DIR%

REM Confirm with user before proceeding
echo.
echo About to install Primus Implant Report Generator for user: %TARGET_USER%
echo Installation directory: %USER_HOME_DIR%\AppData\Local\CreoDent\PrimusImplant
echo.
set /p "CONFIRM=Is this correct? (Y/N): "
if /i not "%CONFIRM%"=="Y" (
    echo Installation cancelled by user.
    goto :End
)

REM Set up paths for the target user
set "USER_DATA_DIR=%USER_HOME_DIR%\AppData\Local\CreoDent\PrimusImplant"
set "USER_DESKTOP=%USER_HOME_DIR%\Desktop"
set "USER_APPDATA=%USER_HOME_DIR%\AppData\Roaming"
set "USER_START_MENU=%USER_APPDATA%\Microsoft\Windows\Start Menu\Programs\CreoDent"

echo.
echo Target directory: %USER_DATA_DIR%
echo Desktop path: %USER_DESKTOP%
echo Start menu path: %USER_START_MENU%

:CreateDirectories
echo.
echo Creating directory structure...

REM Create the directory structure
mkdir "%USER_DATA_DIR%" 2>nul
mkdir "%USER_DATA_DIR%\updates" 2>nul
mkdir "%USER_DATA_DIR%\logs" 2>nul

REM Check if directories were created successfully
if not exist "%USER_DATA_DIR%" (
    echo Error: Failed to create user data directory: %USER_DATA_DIR%
    echo Please check permissions.
    goto :ErrorEnd
)

echo Successfully created directories

REM Copy the main executable to user directory first
echo Copying application files...
set "FILES_COPIED=0"
if exist "%~dp0Primus Implant Report Generator 1.0.4.exe" (
    copy "%~dp0Primus Implant Report Generator 1.0.4.exe" "%USER_DATA_DIR%\" >nul 2>&1
    if %errorlevel%==0 (
        echo Copied main executable
        set /a FILES_COPIED+=1
    )
)

REM Copy data files to user directory
if exist "%~dp0*.csv" (
    copy "%~dp0*.csv" "%USER_DATA_DIR%\" >nul 2>&1
    if %errorlevel%==0 (
        echo Copied CSV files
        set /a FILES_COPIED+=1
    )
)
if exist "%~dp0*.png" (
    copy "%~dp0*.png" "%USER_DATA_DIR%\" >nul 2>&1
    if %errorlevel%==0 (
        echo Copied PNG files
        set /a FILES_COPIED+=1
    )
)
if exist "%~dp0*.ico" (
    copy "%~dp0*.ico" "%USER_DATA_DIR%\" >nul 2>&1
    if %errorlevel%==0 (
        echo Copied ICO files
        set /a FILES_COPIED+=1
    )
)

if %FILES_COPIED%==0 (
    echo No data files found to copy
)

REM Set proper ownership of the created directories to the target user
echo Setting directory ownership...
icacls "%USER_DATA_DIR%" /setowner "%TARGET_USER%" /t /c >nul 2>&1
icacls "%USER_DATA_DIR%" /grant "%TARGET_USER%":(OI)(CI)F /t /c >nul 2>&1

REM Create registry entry for user data path
echo Setting registry entry...
reg add "HKLM\SOFTWARE\CreoDent\PrimusImplant" /v "UserDataPath" /t REG_SZ /d "%%LOCALAPPDATA%%\CreoDent\PrimusImplant" /f >nul 2>&1
reg add "HKLM\SOFTWARE\CreoDent\PrimusImplant" /v "InstalledForUser" /t REG_SZ /d "%TARGET_USER%" /f >nul 2>&1

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
echo ================================================
echo Setup completed successfully!
echo ================================================
echo Application installed for user: %TARGET_USER%
echo Application data directory: %USER_DATA_DIR%
echo Desktop shortcut: %USER_DESKTOP%\Primus Implant Report Generator.lnk
echo Start menu shortcut: %USER_START_MENU%\Primus Implant Report Generator.lnk
echo ================================================
echo.
goto :End

:CreateDesktopShortcut
REM Set paths to point to the USER's copy of the executable
set "USER_EXE_PATH=%USER_DATA_DIR%\Primus Implant Report Generator 1.0.4.exe"
set "SHORTCUT_PATH=%USER_DESKTOP%\Primus Implant Report Generator.lnk"

REM Ensure desktop directory exists
if not exist "%USER_DESKTOP%" mkdir "%USER_DESKTOP%" 2>nul

REM Verify the executable was copied to user directory
if not exist "%USER_EXE_PATH%" (
    echo Warning: Executable not found in user directory, using original location
    set "USER_EXE_PATH=%~dp0Primus Implant Report Generator 1.0.4.exe"
)

echo Creating desktop shortcut pointing to: %USER_EXE_PATH%

REM Create desktop shortcut using PowerShell (simpler version)
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%SHORTCUT_PATH%'); $Shortcut.TargetPath = '%USER_EXE_PATH%'; $Shortcut.WorkingDirectory = '%USER_DATA_DIR%'; $Shortcut.IconLocation = '%USER_EXE_PATH%'; $Shortcut.Description = 'Primus Implant Report Generator'; $Shortcut.Save()" 2>nul

if exist "%SHORTCUT_PATH%" (
    echo Desktop shortcut created successfully
    REM Set ownership of the shortcut
    icacls "%SHORTCUT_PATH%" /setowner "%TARGET_USER%" /c >nul 2>&1
) else (
    echo Warning: Desktop shortcut creation failed
)

exit /b

:CreateStartMenuShortcut
REM Ensure Start Menu directory exists
if not exist "%USER_START_MENU%" mkdir "%USER_START_MENU%" 2>nul

set "START_SHORTCUT_PATH=%USER_START_MENU%\Primus Implant Report Generator.lnk"

echo Creating Start Menu shortcut pointing to: %USER_EXE_PATH%

REM Create Start Menu shortcut using PowerShell (simpler version)
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%START_SHORTCUT_PATH%'); $Shortcut.TargetPath = '%USER_EXE_PATH%'; $Shortcut.WorkingDirectory = '%USER_DATA_DIR%'; $Shortcut.IconLocation = '%USER_EXE_PATH%'; $Shortcut.Description = 'Primus Implant Report Generator'; $Shortcut.Save()" 2>nul

if exist "%START_SHORTCUT_PATH%" (
    echo Start Menu shortcut created successfully
    REM Set ownership of the shortcut and directory
    icacls "%USER_START_MENU%" /setowner "%TARGET_USER%" /t /c >nul 2>&1
    icacls "%START_SHORTCUT_PATH%" /setowner "%TARGET_USER%" /c >nul 2>&1
) else (
    echo Warning: Start Menu shortcut creation failed
)

exit /b

:CreateUninstaller
REM Create uninstaller script in the user data directory
(
echo @echo off
echo echo Uninstalling Primus Implant Report Generator...
echo echo Installation was for user: %TARGET_USER%
echo echo.
echo.
echo REM Stop any running instances
echo taskkill /f /im "Primus Implant Report Generator 1.0.4.exe" 2^>nul
echo timeout /t 2 /nobreak ^>nul
echo.
echo REM Remove desktop shortcut
echo del "%USER_DESKTOP%\Primus Implant Report Generator.lnk" 2^>nul
echo.
echo REM Remove Start Menu shortcuts
echo rmdir /s /q "%USER_START_MENU%" 2^>nul
echo.
echo REM Remove user data directory ^(optional^)
echo echo.
echo echo Do you want to remove user data and settings? ^(Y/N^)
echo set /p "REMOVE_DATA="
echo if /i "%%REMOVE_DATA%%"=="Y" ^(
echo     echo Removing user data...
echo     rmdir /s /q "%USER_DATA_DIR%" 2^>nul
echo ^) else ^(
echo     echo User data preserved at: %USER_DATA_DIR%
echo ^)
echo.
echo REM Remove registry entries
echo reg delete "HKLM\SOFTWARE\CreoDent\PrimusImplant" /f 2^>nul
echo.
echo echo Uninstallation complete.
echo echo.
echo pause
) > "%USER_DATA_DIR%\uninstall.bat"

if exist "%USER_DATA_DIR%\uninstall.bat" (
    echo Uninstaller created successfully
    REM Set ownership of uninstaller
    icacls "%USER_DATA_DIR%\uninstall.bat" /setowner "%TARGET_USER%" /c >nul 2>&1
) else (
    echo Warning: Uninstaller creation failed
)

exit /b

:ErrorEnd
echo.
echo Installation encountered errors. Please check:
echo 1. Verify the target user account exists
echo 2. Ensure you're running as administrator
echo 3. Check Windows permissions for user directories
echo.
pause
exit /b 1

:End
echo Installation script completed.
timeout /t 3 /nobreak >nul