@echo off
echo 正在启动银行账单批量处理工具...

:: 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到Python安装，请先安装Python 3.7或更高版本。
    echo 您可以从 https://www.python.org/downloads/ 下载并安装Python。
    pause
    exit /b 1
)

:: 检查依赖包是否安装
echo 检查依赖包...
python -c "import PyPDF2, pandas, pdfplumber" >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装依赖包，请稍候...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo 错误: 安装依赖包失败。
        pause
        exit /b 1
    )
)

:: 启动应用程序
echo 启动应用程序...
python main.py

exit /b 0