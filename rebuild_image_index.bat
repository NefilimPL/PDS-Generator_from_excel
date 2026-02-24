@echo off
setlocal

set SCRIPT_DIR=%~dp0

if exist "%SCRIPT_DIR%python_runtime\python.exe" (
    "%SCRIPT_DIR%python_runtime\python.exe" "%SCRIPT_DIR%rebuild_image_index.py" %*
) else (
    py -3 "%SCRIPT_DIR%rebuild_image_index.py" %*
)

endlocal
