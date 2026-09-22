@echo off
chcp 65001 >nul
cd /d "%~dp0.."

REM 作用：在本机创建 Windows 计划任务「信息流有效表_每天15点」，每天 15:00 运行定时 bat。
REM 何时改：计划时间变化时改 /ST 参数。总后台只读取和启停该任务，不自动创建。

set "TASK_NAME=信息流有效表_每天15点"
set "OLD_TASK=信息流有效表_每天7点"
set "RUN_BAT=%~dp0定时静默运行.bat"

schtasks /Query /TN "%OLD_TASK%" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo 删除旧计划任务：%OLD_TASK%
    schtasks /Delete /TN "%OLD_TASK%" /F >nul 2>&1
)

schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo 更新已有计划任务时间为 15:00：%TASK_NAME%
    schtasks /Change /TN "%TASK_NAME%" /TR "\"%RUN_BAT%\"" /ST 15:00 >nul 2>&1
    if errorlevel 1 (
        echo 更新失败，尝试删除后重建…
        schtasks /Delete /TN "%TASK_NAME%" /F >nul 2>&1
    ) else (
        echo 已更新：每天 15:00 运行
        goto :done
    )
)

schtasks /Create /TN "%TASK_NAME%" /TR "\"%RUN_BAT%\"" /SC DAILY /ST 15:00 /RL HIGHEST /F
if errorlevel 1 (
    echo 创建失败。请以管理员身份运行本 bat，或手动在「任务计划程序」中创建。
    pause
    exit /b 1
)

:done
echo 计划任务：%TASK_NAME%
echo 每天 15:00 运行，命令：%RUN_BAT%
echo 可在总后台「定时任务」页查看和启停。
pause
