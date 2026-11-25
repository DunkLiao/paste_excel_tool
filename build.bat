@echo off
REM 重新打包免安裝執行檔 -> dist\PasteExcelTool.exe
chcp 65001 >nul
cd /d "%~dp0"

python -m pip install -r requirements.txt
python -m pip install pyinstaller

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name PasteExcelTool ^
    paste_excel_tool.py

echo.
echo 完成: dist\PasteExcelTool.exe
pause
