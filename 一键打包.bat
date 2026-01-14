@echo off
chcp 65001 >nul
echo ========================================
echo   库存表管理系统 - 一键打包工具
echo ========================================
echo.

echo [1/3] 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未检测到Python环境
    echo 请先安装Python 3.8或更高版本
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)
python --version
echo ✓ Python环境检查通过
echo.

echo [2/3] 安装依赖库...
pip install -r requirements.txt
if errorlevel 1 (
    echo ❌ 错误: 依赖安装失败
    pause
    exit /b 1
)
echo ✓ 依赖安装完成
echo.

echo [3/3] 开始打包...
pyinstaller --name=库存表管理系统 --onefile --windowed --clean improve_inventory_gui.py
if errorlevel 1 (
    echo ❌ 错误: 打包失败
    pause
    exit /b 1
)
echo ✓ 打包完成
echo.

echo ========================================
echo   打包成功!
echo ========================================
echo.
echo 可执行文件位置: dist\库存表管理系统.exe
echo.
echo 按任意键打开输出目录...
pause >nul
explorer dist
