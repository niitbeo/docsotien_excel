@echo off
chcp 65001 >nul
color 0A
echo =========================================
echo   CÀI ĐẶT DOCTIEN VÀO EXCEL TỰ ĐỘNG
echo =========================================
echo.
echo Đang chạy script PowerShell...
echo.

PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0CAI_DAT_TU_DONG.ps1"

pause
