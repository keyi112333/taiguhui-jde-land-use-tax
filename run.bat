@echo off
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo Starting JDE land use tax flow...
"%PY%" main.py
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
    echo Completed.
) else (
    echo Failed, exit code: %EXITCODE%
)
pause
exit /b %EXITCODE%
