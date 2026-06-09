@echo off
chcp 65001 >nul
echo 正在启动PDF编辑器...
cd /d "%~dp0"
if exist "dist\PDF编辑器.exe" (
    start "" "dist\PDF编辑器.exe"
) else if exist "PDF编辑器.exe" (
    start "" "PDF编辑器.exe"
) else (
    echo 错误：未找到PDF编辑器可执行文件
    pause
)