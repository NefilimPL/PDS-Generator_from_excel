@echo off
setlocal

set TEMP_DIR=%~dp0build_env
set PY_DIR=%~dp0python_runtime
set PY_VERSION=3.11.6
set PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-amd64.exe

if not exist "%PY_DIR%\python.exe" (
    if exist "%TEMP_DIR%" rmdir /S /Q "%TEMP_DIR%"
    mkdir "%TEMP_DIR%"
    echo Downloading Python installer...
    powershell -Command "Invoke-WebRequest '%PY_URL%' -OutFile '%TEMP_DIR%\python-installer.exe'"
    echo Installing Python into %PY_DIR%...
    "%TEMP_DIR%\python-installer.exe" /quiet InstallAllUsers=0 Include_pip=1 Include_tcltk=1 PrependPath=0 Shortcuts=0 TargetDir="%PY_DIR%" >nul
) else (
    echo Using existing Python in %PY_DIR%...
)

echo Installing PyInstaller...
"%PY_DIR%\python.exe" -m pip install --upgrade pip pyinstaller >nul

echo Building launcher.exe...
"%PY_DIR%\python.exe" -m PyInstaller launcher.py --onefile --noconsole --name launcher

if exist dist\launcher.exe (
    copy dist\launcher.exe launcher.exe >nul
    echo launcher.exe created.
) else (
    echo Build failed.
)

echo Cleaning up...
rmdir /S /Q "%TEMP_DIR%" >nul 2>&1
rmdir /S /Q build dist __pycache__ >nul 2>&1
del launcher.spec >nul 2>&1
echo Done.

