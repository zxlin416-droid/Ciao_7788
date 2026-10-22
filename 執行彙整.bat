@echo off
chcp 65001 >nul
cd /d "%~dp0"
node "彙整執課表.js"
echo.
pause
