@echo off
setlocal

set WORK_DIR=%~dp0build_env
set PY_VERSION=3.11.6
set PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-amd64.exe

if exist "%WORK_DIR%" rmdir /S /Q "%WORK_DIR%"
mkdir "%WORK_DIR%"

echo Downloading Python installer...
powershell -Command "Invoke-WebRequest '%PY_URL%' -OutFile '%WORK_DIR%\python-installer.exe'"

echo Installing temporary Python...
"%WORK_DIR%\python-installer.exe" /quiet InstallAllUsers=0 Include_pip=1 Include_tcltk=1 PrependPath=0 Shortcuts=0 TargetDir="%WORK_DIR%\python" >nul

echo Installing PyInstaller...
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
