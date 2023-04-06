@echo off
powershell.exe -ExecutionPolicy Bypass -Command "& { Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; & .\auto_csv.ps1 }" >> "log.txt"