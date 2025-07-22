@echo off
REM post_install_setup.bat - Run after installation to set up user directories

REM Get the currently logged-in user (not the elevated admin)
for /f "tokens=3" %%i in ('query session ^| find "Active"') do set ACTIVE_USER=%%i

REM If we can't detect active user, try different method
if "%ACTIVE_USER%"=="" (
    for /f "skip=1 tokens=1" %%i in ('query user ^| find "Active"') do set ACTIVE_USER=%%i
)

REM Create user-specific directories for the logged-in user
if not "%ACTIVE_USER%"=="" (
    set "USER_DATA_DIR=C:\Users\%ACTIVE_USER%\AppData\Local\CreoDent\PrimusImplant"
    
    REM Create the directory structure
    mkdir "%USER_DATA_DIR%" 2>nul
    mkdir "%USER_DATA_DIR%\updates" 2>nul
    mkdir "%USER_DATA_DIR%\logs" 2>nul
    
    REM Copy data files to user directory
    copy "%~dp0*.csv" "%USER_DATA_DIR%\" >nul 2>&1
    copy "%~dp0*.png" "%USER_DATA_DIR%\" >nul 2>&1
    copy "%~dp0*.ico" "%USER_DATA_DIR%\" >nul 2>&1
    
    REM Create registry entry for user data path
    reg add "HKLM\SOFTWARE\CreoDent\PrimusImplant" /v "UserDataPath" /t REG_SZ /d "%%LOCALAPPDATA%%\CreoDent\PrimusImplant" /f >nul 2>&1
    
    REM Set permissions for the user
    icacls "%USER_DATA_DIR%" /grant "%ACTIVE_USER%:F" /t >nul 2>&1
)

REM Also create public version accessible to all users
set "PUBLIC_DATA_DIR=%ALLUSERSPROFILE%\CreoDent\PrimusImplant"
mkdir "%PUBLIC_DATA_DIR%" 2>nul

REM Copy data files to public directory as backup
copy "%~dp0*.csv" "%PUBLIC_DATA_DIR%\" >nul 2>&1
copy "%~dp0*.png" "%PUBLIC_DATA_DIR%\" >nul 2>&1
copy "%~dp0*.ico" "%PUBLIC_DATA_DIR%\" >nul 2>&1

REM Make sure Everyone can read these files
icacls "%PUBLIC_DATA_DIR%" /grant "Everyone:R" /t >nul 2>&1

echo User directory setup completed.
timeout /t 2 /nobreak >nul