@echo off
chcp 65001 >nul 2>&1
title 条形码批量生成工具
cd /d "%~dp0"

set "PYW=C:\Users\Administrator\AppData\Local\Programs\Python\Python313\pythonw.exe"
set "PY=C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe"

rem 优先使用无控制台窗口的解释器（界面更干净）
if exist "%PYW%" (
    start "" "%PYW%" "%~dp0barcode_gui.py" %*
    exit /b
)

rem 兜底：无 pythonw 则用普通解释器
if exist "%PY%" (
    "%PY%" "%~dp0barcode_gui.py" %*
    exit /b
)

rem 最后兜底：PATH 里的 python
where pythonw >nul 2>&1 && (start "" pythonw "%~dp0barcode_gui.py" %* & exit /b)
where python  >nul 2>&1 && ("python" "%~dp0barcode_gui.py" %* & exit /b)

echo.
echo   未找到 Python 运行环境，无法启动图形界面。
echo.
echo   可选方案：
echo     1. 安装 Python（python.org 下载，安装时勾选 Add to PATH）
echo     2. 使用免安装版：双击 "条形码工具（免安装）.exe"（若存在）
echo     3. 使用命令行版：把 Excel 文件拖到 excel_barcode.py 上
echo.
pause