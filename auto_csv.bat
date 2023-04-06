@echo off
setlocal

:: PowerShellスクリプトファイルのパスを取得する
set "SCRIPT_PATH=C:\Users\gloss\Downloads"

:: PowerShellスクリプトを実行し、出力をログファイルにリダイレクトする
powershell.exe -ExecutionPolicy Bypass -File "%SCRIPT_PATH%auto_csv.ps1" >> "%SCRIPT_PATH%log.txt" 2>&1