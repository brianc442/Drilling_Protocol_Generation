@echo off
echo Building Primus Implant Report Generator...

cd C:\Users\bconn\PycharmProjects\Drilling Protocol Generation

python -m PyInstaller --onefile --windowed --icon=icon.ico --name "Primus Implant Report Generator 1.0.5" --add-data="icon.ico:." --add-data="inosys_logo.png:." --add-data="Primus Implant List - Primus Implant List.csv:." main.py

if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)

echo Build completed successfully!
echo Executable created in: dist\
pause