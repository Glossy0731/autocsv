@echo off
setlocal

:: 获取脚本文件的路径
set "SCRIPT_PATH=%~dp0"

:: 执行 PowerShell 脚本并将输出重定向到日志文件中
powershell.exe -ExecutionPolicy Bypass -File "%SCRIPT_PATH%run.ps1" >> "%SCRIPT_PATH%log.txt" 2>&1