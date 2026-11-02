@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo 正在启动 JDE 土地使用税流程...
"%PY%" main.py
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
    echo 运行完成。
) else (
    echo 运行失败，退出码：%EXITCODE%
)
pause
exit /b %EXITCODE%
