@echo off

REM 定义要转换的CSV文件夹路径
set folder="C:\CSV Files"

REM 定义PowerShell脚本内容
set script=^
$config = Get-Content .\config.txt ;^
$folder = $config.Trim() ;^
function ConvertTo-Excel { ... } ;^
while ($true) { ... }

REM 调用PowerShell执行脚本，并将输出记录到log.txt文件中
powershell.exe -ExecutionPolicy Bypass -Command "%script%" -Encoding utf8 >> "log.txt"

pause