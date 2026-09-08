@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 首次安裝 - 法規鑑別追蹤工具
echo ============================================
echo   正在安裝所需套件，請確認電腦已連接網路
echo ============================================
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo.
echo ============================================
echo   安裝完成！
echo   之後可直接雙擊「start_law_tool.bat」開啟工具
echo   （或使用桌面捷徑，作法請見「README.txt」）
echo ============================================
pause
