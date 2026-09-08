@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 法規鑑別追蹤工具
echo ============================================
echo   正在啟動法規鑑別追蹤工具...
echo   請稍候,瀏覽器視窗將自動開啟
echo   （使用時請勿關閉此視窗，關閉視窗會停止工具）
echo ============================================
echo.
python law_tool_app.py
if errorlevel 1 (
    echo.
    echo [錯誤] 工具無法啟動，請確認：
    echo   1. 電腦已安裝 Python（並在安裝時勾選 Add to PATH）
    echo   2. 已先雙擊執行過「first_time_setup.bat」
    echo.
    pause
)
