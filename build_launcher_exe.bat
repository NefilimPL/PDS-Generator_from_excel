@echo off
setlocal

set WORK_DIR=%~dp0build_env
set PY_URL=https://www.python.org/ftp/python/3.11.6/python-3.11.6-embed-amd64.zip

if exist "%WORK_DIR%" rmdir /S /Q "%WORK_DIR%"
mkdir "%WORK_DIR%"

echo Downloading portable Python...
powershell -Command "Invoke-WebRequest '%PY_URL%' -OutFile '%WORK_DIR%\python.zip'"

echo Extracting Python...
powershell -Command "Expand-Archive '%WORK_DIR%\python.zip' '%WORK_DIR%\python'"

powershell -Command "(Get-Content '%WORK_DIR%\python\python311._pth') -replace '#import site','import site' | Set-Content '%WORK_DIR%\python\python311._pth'"

echo Installing pip and PyInstaller...
"%WORK_DIR%\python\python.exe" -m ensurepip >nul
"%WORK_DIR%\python\python.exe" -m pip install --upgrade pip pyinstaller >nul

echo Building launcher.exe...
"%WORK_DIR%\python\python.exe" -m PyInstaller launcher.py --onefile --noconsole --name launcher

if exist dist\launcher.exe (
    copy dist\launcher.exe launcher.exe >nul
    echo launcher.exe created.
) else (
    echo Build failed.
)

echo Cleaning up...
rmdir /S /Q "%WORK_DIR%"
rmdir /S /Q build dist __pycache__ >nul 2>&1
del launcher.spec >nul 2>&1
echo Done.
