@echo off
chcp 65001 >nul
echo 正在启动设备缺陷统计管理系统...
python "auto_fill_defects.py"
if %errorlevel% neq 0 (
    echo.
    echo 程序运行出错！
    echo 请确保已安装 Python 以及相关依赖库 (pandas, openpyxl, pywin32, matplotlib)。
    echo.
    pause
)
