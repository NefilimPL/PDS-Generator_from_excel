@echo off
setlocal

set "BASE_DIR=%~dp0"
set "PYTHON_EXE=%BASE_DIR%python_runtime\python.exe"
set "PDS_LAUNCH_SOURCE=%~f0"

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" "%BASE_DIR%pds_headless.py" %*
) else (
  python "%BASE_DIR%pds_headless.py" %*
)

endlocal
