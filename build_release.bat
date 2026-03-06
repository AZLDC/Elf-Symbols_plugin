@echo off
setlocal
cd /d "%~dp0"

rem -- Check PyInstaller --
py -3 -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller not found. Run: py -3 -m pip install pyinstaller
    pause
    exit /b 1
)

rem -- Build exe --
echo [1/4] Building exe ...
py -3 -m PyInstaller --noconfirm Elf-Symbols_plugin.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    pause
    exit /b 1
)

rem -- Create Release folder --
if not exist "Release" mkdir "Release"

rem -- Pack Excutable_APP.zip --
echo [2/4] Packing Excutable_APP.zip ...
if exist "Release\Excutable_APP.zip" del "Release\Excutable_APP.zip"
py -3 -c "import zipfile;z=zipfile.ZipFile('Release/Excutable_APP.zip','w',zipfile.ZIP_DEFLATED);z.write('dist/Elf-Symbols_plugin.exe','Elf-Symbols_plugin.exe');z.close()"
if errorlevel 1 (
    echo [ERROR] Failed to create Excutable_APP.zip
    pause
    exit /b 1
)

rem -- Pack Excutable_Script.zip (filenames read from .spec to avoid encoding issues) --
echo [3/4] Packing Excutable_Script.zip ...
if exist "Release\Excutable_Script.zip" del "Release\Excutable_Script.zip"
py -3 -c "import zipfile,re;s=open('Elf-Symbols_plugin.spec',encoding='utf-8').read();p=chr(39)+'([^'+chr(39)+']+[.](?:png|jpg))'+chr(39);imgs=list(set(re.findall(p,s)));z=zipfile.ZipFile('Release/Excutable_Script.zip','w',zipfile.ZIP_DEFLATED);[z.write(f) for f in ['Elf-Symbols_plugin.pyw','Cursors_FIX.py']+imgs];z.close()"
if errorlevel 1 (
    echo [ERROR] Failed to create Excutable_Script.zip
    pause
    exit /b 1
)

rem -- Cleanup --
echo [4/4] Cleaning up ...
if exist "build" rmdir /s /q "build"
if exist "dist"  rmdir /s /q "dist"

echo.
echo Done!
echo   Release\Excutable_APP.zip
echo   Release\Excutable_Script.zip
pause
