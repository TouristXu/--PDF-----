@echo off
cd /d "%~dp0"
echo Running PDF编辑器_debug.exe...
PDF编辑器_debug.exe
echo Exit code: %errorlevel%
pause