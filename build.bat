@echo off
echo ================================
echo  Personal Assistant Build Tool
echo ================================
echo.

echo [1/3] Installing dependencies...
python -m pip install pyinstaller Pillow
if %errorlevel% neq 0 (
    echo.
    echo ERROR: pip install failed.
    pause
    exit /b 1
)

echo.
echo [2/3] Building exe, please wait...
python -m PyInstaller --onefile --windowed --name "PersonalAssistant" --add-data "支付系统;支付系统" powermill_copier.py
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo ================================
echo  Build successful!
echo  Output: dist\PersonalAssistant.exe
echo ================================
explorer dist
pause
