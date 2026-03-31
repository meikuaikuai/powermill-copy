@echo off
chcp 65001 >nul
echo ================================
echo  Personal Assistant 打包工具
echo ================================
echo.

echo [1/3] 安装依赖...
python -m pip install pyinstaller Pillow
if %errorlevel% neq 0 (
    echo.
    echo 错误: pip 安装失败，请确认 Python 已安装并添加到系统 PATH
    echo 提示: 安装 Python 时勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo.
echo [2/3] 打包中，请稍候...
python -m PyInstaller --onefile --windowed --name "Personal Assistant" --add-data "支付系统;支付系统" powermill_copier.py
if %errorlevel% neq 0 (
    echo.
    echo 错误: 打包失败，请检查上方错误信息
    pause
    exit /b 1
)

echo.
echo ================================
echo  打包成功！
echo  文件位置: dist\Personal Assistant.exe
echo ================================
explorer dist
pause
