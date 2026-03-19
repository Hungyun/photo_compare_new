@echo off
REM -------------------------------
REM AGPL-compliant build script for Windows with Eel support
REM -------------------------------

REM 1. 建立虛擬環境
python -m venv venv
call venv\Scripts\activate.bat

REM 2. 更新 pip
pip install --upgrade pip

REM 3. 安裝依賴
pip install -r requirements.txt
pip install pyinstaller

REM 4. 使用 PyInstaller 打包，包含 web 資料夾
pyinstaller main.py --onefile --add-data "web;web"

echo Build finished! Generated exe in dist\ folder
pause