@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT=%SCRIPT_DIR%Elf-Symbols_plugin.pyw"
set "PYW=%SystemRoot%\pyw.exe"
set "PY=%SystemRoot%\py.exe"
if not exist "%SCRIPT%" (
    echo 找不到 %SCRIPT% ，請確認批次檔與 pyw 位於同一資料夾。
    pause
    exit /b 1
)
if exist "%PYW%" (
    start "" "%PYW%" -3 "%SCRIPT%"
) else if exist "%PY%" (
    start "" "%PY%" -3 "%SCRIPT%"
) else (
    echo 找不到 pyw.exe 或 py.exe ，請確認 Python Launcher 已安裝。
)
pause
