@echo off
setlocal

for %%I in ("%~dp0..") do set "PROJECT_DIR=%%~fI"
set "PYTHON_EXE=%PROJECT_DIR%\python_runtime\python.exe"
set "PDS_LAUNCH_SOURCE=%~f0"
set "PDS_DEFAULT_INTERVAL=1800"
set "PDS_LOG_DIR=%PROJECT_DIR%\logs\windows_services"

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" "%PROJECT_DIR%\rebuild_image_index.py" --service --interval-seconds %PDS_DEFAULT_INTERVAL% --log-dir "%PDS_LOG_DIR%" %*
) else (
  python "%PROJECT_DIR%\rebuild_image_index.py" --service --interval-seconds %PDS_DEFAULT_INTERVAL% --log-dir "%PDS_LOG_DIR%" %*
)

endlocal
