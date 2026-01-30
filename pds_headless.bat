@echo off
setlocal

set "BASE_DIR=%~dp0"
set "PYTHON_EXE=%BASE_DIR%python_runtime\python.exe"

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" "%BASE_DIR%pds_headless.py" %*
) else (
  python "%BASE_DIR%pds_headless.py" %*
)

endlocal
