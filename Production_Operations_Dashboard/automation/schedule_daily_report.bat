@echo off
REM ============================================================================
REM 生产日报自动化 - Windows 计划任务配置
REM 配置此批处理文件作为 Windows 计划任务的任务触发程序
REM
REM 用法:
REM 1. 以管理员身份运行此脚本
REM 2. 设置任务计划程序在每天 17:00 (5 PM) 运行
REM ============================================================================

setlocal enabledelayedexpansion

REM 设置工作目录
cd /d "%~dp0"

REM 设置 Python 路径 (根据实际安装位置修改)
set PYTHON_PATH=python

REM 日志文件路径
set LOG_FILE=logs\schedule.log

REM 获取当前日期和时间
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c%%a%%b)
for /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set mytime=%%a%%b)

REM 执行 Python 脚本
echo [%mydate% %mytime%] 开始执行日报自动化... >> %LOG_FILE%

%PYTHON_PATH% daily_report_automation.py >> %LOG_FILE% 2>&1

if errorlevel 1 (
    echo [%mydate% %mytime%] 日报生成失败 (错误码: %errorlevel%) >> %LOG_FILE%
    exit /b 1
) else (
    echo [%mydate% %mytime%] 日报生成成功 >> %LOG_FILE%
    exit /b 0
)

endlocal
