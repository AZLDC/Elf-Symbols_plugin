py -3 -m PyInstaller "新注音繁簡快速切換.spec"
move dist\*.exe .\
rd build /s/q
rd dist
pause