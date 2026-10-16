@echo off
chcp 65001 >nul
title 供应链库存与物流看板
echo.
echo 正在准备运行环境，请稍候……
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set PYTHON_CMD=py
) else (
    set PYTHON_CMD=python
)

%PYTHON_CMD% --version >nul 2>nul
if errorlevel 1 (
    echo 未检测到 Python。
    echo 请先从 https://www.python.org/downloads/ 安装 Python 3.11，
    echo 并在安装时勾选 Add Python to PATH。
    pause
    exit /b 1
)

if not exist ".venv" (
    echo 正在创建虚拟环境……
    %PYTHON_CMD% -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo 正在打开网页……
streamlit run streamlit_app.py
pause
