@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在打包，请稍候（首次约 1～2 分钟）...
py -3 -m PyInstaller --noconfirm --clean InvoiceTaxTool.spec
if errorlevel 1 (
  echo 打包失败，请确认已安装: py -3 -m pip install pyinstaller -r requirements.txt
  pause
  exit /b 1
)
echo.
echo 完成。可执行文件: dist\InvoiceTaxTool.exe
pause
