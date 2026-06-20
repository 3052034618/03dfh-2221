@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   冷链探头校准核对工具
echo ============================================
echo.

if not exist "venv" (
    echo [首次运行] 正在创建虚拟环境并安装依赖...
    python -m venv venv
    if errorlevel 1 (
        echo [错误] 无法创建虚拟环境，请确认已安装 Python 3.8+
        pause
        exit /b 1
    )
    call venv\Scripts\activate.bat
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败
        pause
        exit /b 1
    )
    echo.
    echo [完成] 环境初始化成功
    echo.
) else (
    call venv\Scripts\activate.bat
)

if "%~1"=="" (
    echo 可用命令:
    echo   import    导入车辆清单、标准读数、探头读数
    echo   check     执行校准偏差核对
    echo   summary   生成车队经理摘要报告
    echo.
    echo 示例:
    echo   coldchain import
    echo   coldchain check -c 冷冻
    echo   coldchain check -t 1.0 -k 2.0
    echo   coldchain summary -d 30
    echo.
    set /p cmd="请输入命令: "
    python coldchain_calib.py !cmd!
) else (
    python coldchain_calib.py %*
)

endlocal
echo.
pause
