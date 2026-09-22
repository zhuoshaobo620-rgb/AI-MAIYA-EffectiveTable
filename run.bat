@echo off
chcp 65001 >nul
echo ========================================
echo   Daily Report Tool
echo ========================================
echo.
echo Starting, please wait...
echo.

rem 优先使用已装好依赖的 Python 3.14（若不存在则回退到 PATH 中的 python）
set "PY_EXE=C:\Users\Administrator\AppData\Local\Programs\Python\Python314\python.exe"
if exist "%PY_EXE%" (
    "%PY_EXE%" "%~dp0daily_report.py"
) else (
    python "%~dp0daily_report.py"
)

echo.
echo ========================================
if %ERRORLEVEL% EQU 0 (
echo   Done!
) else (
echo   Error occurred. Check logs above.
)
echo ========================================
pause
