@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT=%SCRIPT_DIR%Elf-Symbols_plugin.pyw"
set "PY=%SystemRoot%\py.exe"
if not exist "%SCRIPT%" (
    echo 找不到 %SCRIPT% ，請確認批次檔與 pyw 位於同一資料夾。
    pause
    exit /b 1
)
if exist "%PY%" (
    start "" "%PY%" -3 "%SCRIPT%"
) else (
    echo 找不到 py.exe ，請確認 Python 已安裝。
)
pause
