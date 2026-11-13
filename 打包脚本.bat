@echo off
chcp 65001 >nul
echo ==========================================
echo  机制电价市场力分析工具 - 打包脚本
echo ==========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python 3.12+
    pause
    exit /b 1
)

REM 安装依赖
echo [1/3] 安装依赖...
pip install pandas openpyxl scikit-learn pyinstaller -q
if errorlevel 1 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

REM 执行打包
echo [2/3] 开始打包...
pyinstaller --onefile --windowed --name "机制电价市场力分析工具" gui_main.py
if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo [3/3] 打包完成！
echo.
echo 可执行文件位于: dist\机制电价市场力分析工具.exe
echo 请将该exe文件复制到任意Windows电脑即可运行。
echo.
pause
