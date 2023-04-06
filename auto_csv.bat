@echo off
powershell.exe -ExecutionPolicy Bypass -File "auto_csv.ps1" -Encoding utf8 >> "log.txt"
pause