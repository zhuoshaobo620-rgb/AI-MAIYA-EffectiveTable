@echo off
chcp 65001 >nul
cd /d "%~dp0.."

REM 作用：定时静默运行有效表日报（经浏览器排队锁）。每次独立分卷日志，避免 >> 争锁导致次日任务挂起。
REM 何时改：Python 路径或锁任务名变化时。

if not exist "06_日志" mkdir "06_日志"
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "RUN_TAG=%%I"
set "LOG=06_日志\定时运行_%RUN_TAG%.log"
set "MASTER=06_日志\定时运行.log"

echo ============================================================>> "%MASTER%"
echo 开始 %date% %time% 分卷=%LOG%>> "%MASTER%"
echo ============================================================>> "%MASTER%"
echo ============================================================>> "%LOG%"
echo 开始 %date% %time%>> "%LOG%"
echo ============================================================>> "%LOG%"

set "PY_EXE=C:\Users\Administrator\AppData\Local\Programs\Python\Python314\python.exe"
if not exist "%PY_EXE%" (
    set "PY_EXE=python"
)

set "LOCK_RUNNER=E:\AI项目\_共享\定时任务排队\run_with_browser_lock.py"
REM 排队脚本必须用 python.exe，pythonw 无控制台会导致 >> 日志 收不到 daily_report 输出
set "CONTROL_PY=E:\AI项目\运营自动化管理后台\.venv\Scripts\python.exe"
if exist "%CONTROL_PY%" (
    set "LOCK_PY=%CONTROL_PY%"
) else (
    set "LOCK_PY=%PY_EXE%"
)

"%LOCK_PY%" "%LOCK_RUNNER%" --task 信息流有效表日报 --hide-window -- "%PY_EXE%" -u "%~dp0..\daily_report.py" >> "%LOG%" 2>&1
set ERR=%ERRORLEVEL%

echo.>> "%LOG%"
echo 结束 %date% %time% 退出码=%ERR%>> "%LOG%"
echo 分卷结束 %date% %time% 退出码=%ERR% 文件=%LOG%>> "%MASTER%"
exit /b %ERR%
