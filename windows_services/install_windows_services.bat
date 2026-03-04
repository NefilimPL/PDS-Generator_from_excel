@echo off
setlocal EnableExtensions EnableDelayedExpansion

for %%I in ("%~dp0..") do set "PROJECT_DIR=%%~fI"
set "SERVICES_DIR=%~dp0"
set "TARGET=both"
set "PDF_SERVICE_NAME=PDSGeneratorPdfService"
set "INDEX_SERVICE_NAME=PDSGeneratorImageIndexService"
set "PDF_DISPLAY_NAME=PDS Generator - PDF Generation Service"
set "INDEX_DISPLAY_NAME=PDS Generator - Image Index Service"
set "PDF_DESCRIPTION=Generates PDF/PDS files from the configured Excel workbook."
set "INDEX_DESCRIPTION=Rebuilds the shared local image index used by PDS Generator."
set "PDF_INTERVAL_SECONDS=300"
set "INDEX_INTERVAL_SECONDS=1800"
set "PDF_IMAGE_INDEX_MODE=cache-only"
set "CONFIG_PATH="
set "EXCEL_PATH="
set "INDEX_DIR_ARGS="
set "NSSM_EXE="
set "PYTHON_EXE="
set "RUN_USER="
set "RUN_PASSWORD="
set "START_MODE=SERVICE_AUTO_START"
set "PYTHON_VERSION=3.11.6"
set "PYTHON_INSTALLER_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe"
set "NSSM_VERSION=2.24"
set "NSSM_ZIP_URL=https://nssm.cc/release/nssm-%NSSM_VERSION%.zip"
set "PYTHON_INSTALLER_PATH=%PROJECT_DIR%\python_runtime\python-installer.exe"
set "PYTHON_DIR=%PROJECT_DIR%\python_runtime"
if not defined PYTHON_EXE set "PYTHON_EXE=%PYTHON_DIR%\python.exe"
set "REQUIREMENTS_FILE=%PROJECT_DIR%\requirements.txt"
set "LOG_DIR=%PROJECT_DIR%\logs\windows_services"
set "BOOTSTRAP_LOG=%LOG_DIR%\install_windows_services.log"

call :parse_args %*
if errorlevel 1 exit /b 1

call :ensure_elevated %*
if errorlevel 2 exit /b 0
if errorlevel 1 exit /b 1

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
echo ==== %DATE% %TIME% ==== > "%BOOTSTRAP_LOG%"
call :log Bootstrap started.

call :ensure_nssm
if errorlevel 1 exit /b 1

call :ensure_python_runtime
if errorlevel 1 exit /b 1

call :install_requirements
if errorlevel 1 exit /b 1

set "PDF_SCRIPT_PATH=%PROJECT_DIR%\pds_headless.py"
set "INDEX_SCRIPT_PATH=%PROJECT_DIR%\rebuild_image_index.py"

if not exist "%PDF_SCRIPT_PATH%" (
  call :fail Missing file "%PDF_SCRIPT_PATH%".
  exit /b 1
)
if not exist "%INDEX_SCRIPT_PATH%" (
  call :fail Missing file "%INDEX_SCRIPT_PATH%".
  exit /b 1
)

call :warn_if_present "PDSGenerator"
call :warn_if_present "PDSGeneratorPdf"
call :warn_if_present "PDSGeneratorImageIndex"

if /I not "%TARGET%"=="index" (
  set "CURRENT_SERVICE_NAME=%PDF_SERVICE_NAME%"
  set "CURRENT_DISPLAY_NAME=%PDF_DISPLAY_NAME%"
  set "CURRENT_DESCRIPTION=%PDF_DESCRIPTION%"
  set "CURRENT_STDOUT_LOG=%LOG_DIR%\pdf_generation_stdout.log"
  set "CURRENT_STDERR_LOG=%LOG_DIR%\pdf_generation_stderr.log"
  set "CURRENT_APP_PARAMETERS=""%PDF_SCRIPT_PATH%"" --service --interval-seconds %PDF_INTERVAL_SECONDS% --log-dir ""%LOG_DIR%"" --image-index-mode %PDF_IMAGE_INDEX_MODE%"
  if defined CONFIG_PATH set "CURRENT_APP_PARAMETERS=!CURRENT_APP_PARAMETERS! --config ""%CONFIG_PATH%"""
  if defined EXCEL_PATH set "CURRENT_APP_PARAMETERS=!CURRENT_APP_PARAMETERS! --excel ""%EXCEL_PATH%"""
  call :install_or_update_service
  if errorlevel 1 exit /b 1
)

if /I not "%TARGET%"=="pdf" (
  set "CURRENT_SERVICE_NAME=%INDEX_SERVICE_NAME%"
  set "CURRENT_DISPLAY_NAME=%INDEX_DISPLAY_NAME%"
  set "CURRENT_DESCRIPTION=%INDEX_DESCRIPTION%"
  set "CURRENT_STDOUT_LOG=%LOG_DIR%\image_index_stdout.log"
  set "CURRENT_STDERR_LOG=%LOG_DIR%\image_index_stderr.log"
  set "CURRENT_APP_PARAMETERS=""%INDEX_SCRIPT_PATH%"" --service --interval-seconds %INDEX_INTERVAL_SECONDS% --log-dir ""%LOG_DIR%"""
  if defined CONFIG_PATH set "CURRENT_APP_PARAMETERS=!CURRENT_APP_PARAMETERS! --config ""%CONFIG_PATH%"""
  if defined EXCEL_PATH set "CURRENT_APP_PARAMETERS=!CURRENT_APP_PARAMETERS! --excel ""%EXCEL_PATH%"""
  if defined INDEX_DIR_ARGS set "CURRENT_APP_PARAMETERS=!CURRENT_APP_PARAMETERS!!INDEX_DIR_ARGS!"
  call :install_or_update_service
  if errorlevel 1 exit /b 1
)

call :log Windows services are ready.
echo.
echo Windows services are ready.
if /I not "%TARGET%"=="index" (
  echo   [%PDF_SERVICE_NAME%] %PDF_DISPLAY_NAME%
  echo     Interval: %PDF_INTERVAL_SECONDS%s
  echo     Mode: %PDF_IMAGE_INDEX_MODE%
)
if /I not "%TARGET%"=="pdf" (
  echo   [%INDEX_SERVICE_NAME%] %INDEX_DISPLAY_NAME%
  echo     Interval: %INDEX_INTERVAL_SECONDS%s
)
echo   Logs: %LOG_DIR%
echo   Installer log: %BOOTSTRAP_LOG%
echo   Open services.msc to verify both services.
exit /b 0

:ensure_elevated
fltmc >nul 2>&1
if not errorlevel 1 exit /b 0
echo Requesting administrator privileges...
set "PDS_INSTALL_ARGS=%*"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process -FilePath '%~f0' -WorkingDirectory '%CD%' -Verb RunAs -ArgumentList $env:PDS_INSTALL_ARGS"
if errorlevel 1 (
  echo ERROR: Administrator privileges are required.
  exit /b 1
)
exit /b 2

:ensure_nssm
call :resolve_nssm
if not errorlevel 1 exit /b 0

call :log NSSM not found locally. Downloading...
if not exist "%SERVICES_DIR%tmp" mkdir "%SERVICES_DIR%tmp"
set "NSSM_ZIP_PATH=%SERVICES_DIR%tmp\nssm-%NSSM_VERSION%.zip"

call :download_file "%NSSM_ZIP_URL%" "%NSSM_ZIP_PATH%"
if errorlevel 1 exit /b 1

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$zip='%NSSM_ZIP_PATH%'; $dest='%SERVICES_DIR%tmp\nssm'; if(Test-Path $dest){Remove-Item $dest -Recurse -Force}; Expand-Archive -LiteralPath $zip -DestinationPath $dest -Force"
if errorlevel 1 (
  call :fail Failed to extract NSSM archive.
  exit /b 1
)

if exist "%SERVICES_DIR%tmp\nssm\nssm-%NSSM_VERSION%\win64\nssm.exe" (
  copy /Y "%SERVICES_DIR%tmp\nssm\nssm-%NSSM_VERSION%\win64\nssm.exe" "%SERVICES_DIR%nssm.exe" >nul
)
if not exist "%SERVICES_DIR%nssm.exe" if exist "%SERVICES_DIR%tmp\nssm\nssm-%NSSM_VERSION%\win32\nssm.exe" (
  copy /Y "%SERVICES_DIR%tmp\nssm\nssm-%NSSM_VERSION%\win32\nssm.exe" "%SERVICES_DIR%nssm.exe" >nul
)

set "NSSM_EXE=%SERVICES_DIR%nssm.exe"
if exist "%NSSM_EXE%" (
  call :log NSSM downloaded to "%NSSM_EXE%".
  exit /b 0
)

call :fail Failed to prepare nssm.exe.
exit /b 1

:ensure_python_runtime
if defined PYTHON_EXE if /I not "%PYTHON_EXE%"=="%PYTHON_DIR%\python.exe" (
  call :python_runtime_usable "%PYTHON_EXE%"
  if not errorlevel 1 (
    call :log Using custom Python executable "%PYTHON_EXE%".
    exit /b 0
  )
  call :fail Custom Python executable is missing or unusable: "%PYTHON_EXE%".
  exit /b 1
)

call :python_runtime_usable "%PYTHON_EXE%"
if not errorlevel 1 (
  call :log Using existing Python runtime "%PYTHON_EXE%".
  exit /b 0
)

call :log Python runtime not found. Preparing local runtime...
if not exist "%PYTHON_DIR%" mkdir "%PYTHON_DIR%"

set "PY_INSTALL_ATTEMPT=1"
:python_install_attempt
if not exist "%PYTHON_INSTALLER_PATH%" (
  call :log Python installer not found locally. Downloading...
  call :download_file "%PYTHON_INSTALLER_URL%" "%PYTHON_INSTALLER_PATH%"
  if errorlevel 1 exit /b 1
)

call :log Installing Python runtime into "%PYTHON_DIR%"...
call :run_python_installer "%PYTHON_INSTALLER_PATH%" "%PYTHON_DIR%"
if errorlevel 1 (
  if %PY_INSTALL_ATTEMPT% GEQ 2 (
    call :fail Python installer failed after retry. See "%BOOTSTRAP_LOG%".
    exit /b 1
  )
  call :log Python installer failed. Removing installer and retrying download...
  del /f /q "%PYTHON_INSTALLER_PATH%" >nul 2>&1
  set /a PY_INSTALL_ATTEMPT+=1
  goto :python_install_attempt
)

set /a WAIT_SECONDS=0
:wait_for_python
call :python_runtime_usable "%PYTHON_EXE%"
if not errorlevel 1 (
  call :log Python runtime installed.
  exit /b 0
)
if %WAIT_SECONDS% GEQ 180 (
  if %PY_INSTALL_ATTEMPT% GEQ 2 (
    call :fail python.exe did not appear after installation. See "%BOOTSTRAP_LOG%".
    exit /b 1
  )
  call :log python.exe did not appear after installation. Removing installer and retrying...
  del /f /q "%PYTHON_INSTALLER_PATH%" >nul 2>&1
  set /a PY_INSTALL_ATTEMPT+=1
  goto :python_install_attempt
)
timeout /t 1 /nobreak >nul
set /a WAIT_SECONDS+=1
goto :wait_for_python

:install_requirements
if not exist "%REQUIREMENTS_FILE%" (
  call :log requirements.txt not found, skipping dependency installation.
  exit /b 0
)

call :requirements_installed
if not errorlevel 1 (
  call :log Required Python packages are already installed.
  exit /b 0
)

call :log Installing Python packages from requirements.txt...
"%PYTHON_EXE%" -m ensurepip --upgrade >nul 2>&1
"%PYTHON_EXE%" -m pip install --disable-pip-version-check --upgrade pip >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 (
  call :fail Failed to upgrade pip. See "%BOOTSTRAP_LOG%".
  exit /b 1
)
"%PYTHON_EXE%" -m pip install --disable-pip-version-check -r "%REQUIREMENTS_FILE%" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 (
  call :fail Failed to install required Python packages. See "%BOOTSTRAP_LOG%".
  exit /b 1
)
call :log Python packages installed.
exit /b 0

:requirements_installed
"%PYTHON_EXE%" -c "import importlib.util, sys; modules=['pandas','PIL','reportlab','requests','openpyxl','cryptography']; missing=[m for m in modules if importlib.util.find_spec(m) is None]; print('All required modules already installed.' if not missing else 'Missing modules: ' + ', '.join(missing)); sys.exit(0 if not missing else 1)" >> "%BOOTSTRAP_LOG%" 2>&1
exit /b %errorlevel%

:install_or_update_service
sc query "%CURRENT_SERVICE_NAME%" >nul 2>&1
if errorlevel 1 (
  call :log Installing service "%CURRENT_SERVICE_NAME%".
  "%NSSM_EXE%" install "%CURRENT_SERVICE_NAME%" "%PYTHON_EXE%" >> "%BOOTSTRAP_LOG%" 2>&1
  if errorlevel 1 (
    call :fail Failed to install service "%CURRENT_SERVICE_NAME%".
    exit /b 1
  )
)

call :log Configuring service "%CURRENT_SERVICE_NAME%".
"%NSSM_EXE%" stop "%CURRENT_SERVICE_NAME%" >nul 2>&1
"%NSSM_EXE%" set "%CURRENT_SERVICE_NAME%" AppDirectory "%PROJECT_DIR%" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed
"%NSSM_EXE%" set "%CURRENT_SERVICE_NAME%" Application "%PYTHON_EXE%" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed
"%NSSM_EXE%" set "%CURRENT_SERVICE_NAME%" AppParameters %CURRENT_APP_PARAMETERS% >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed
"%NSSM_EXE%" set "%CURRENT_SERVICE_NAME%" AppStdout "%CURRENT_STDOUT_LOG%" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed
"%NSSM_EXE%" set "%CURRENT_SERVICE_NAME%" AppStderr "%CURRENT_STDERR_LOG%" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed
"%NSSM_EXE%" set "%CURRENT_SERVICE_NAME%" Start %START_MODE% >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed
"%NSSM_EXE%" set "%CURRENT_SERVICE_NAME%" AppExit Default Restart >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed
sc config "%CURRENT_SERVICE_NAME%" DisplayName= "%CURRENT_DISPLAY_NAME%" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed
sc description "%CURRENT_SERVICE_NAME%" "%CURRENT_DESCRIPTION%" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 goto :service_set_failed

if defined RUN_USER (
  "%NSSM_EXE%" set "%CURRENT_SERVICE_NAME%" ObjectName "%RUN_USER%" "%RUN_PASSWORD%" >> "%BOOTSTRAP_LOG%" 2>&1
  if errorlevel 1 (
    call :fail Failed to configure service account for "%CURRENT_SERVICE_NAME%".
    exit /b 1
  )
)

call :log Starting service "%CURRENT_SERVICE_NAME%".
"%NSSM_EXE%" start "%CURRENT_SERVICE_NAME%" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 (
  call :fail Failed to start service "%CURRENT_SERVICE_NAME%".
  exit /b 1
)
exit /b 0

:service_set_failed
call :fail Failed to configure service "%CURRENT_SERVICE_NAME%".
exit /b 1

:warn_if_present
sc query %~1 >nul 2>&1
if errorlevel 1 exit /b 0
call :log Existing service "%~1" detected.
exit /b 0

:resolve_nssm
if defined NSSM_EXE (
  if exist "%NSSM_EXE%" exit /b 0
  exit /b 1
)
if exist "%SERVICES_DIR%nssm.exe" (
  set "NSSM_EXE=%SERVICES_DIR%nssm.exe"
  exit /b 0
)
if exist "%PROJECT_DIR%\nssm.exe" (
  set "NSSM_EXE=%PROJECT_DIR%\nssm.exe"
  exit /b 0
)
for /f "delims=" %%I in ('where nssm.exe 2^>nul') do (
  set "NSSM_EXE=%%I"
  goto :resolve_nssm_done
)
:resolve_nssm_done
if defined NSSM_EXE exit /b 0
exit /b 1

:download_file
call :log Downloading "%~1"...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "& { param([string]$Url, [string]$Target) [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $Target }" ^
  "%~1" "%~2" >> "%BOOTSTRAP_LOG%" 2>&1
if not errorlevel 1 exit /b 0

call :log PowerShell download failed. Trying curl.exe...
curl.exe -L --fail --output "%~2" "%~1" >> "%BOOTSTRAP_LOG%" 2>&1
if errorlevel 1 (
  call :fail Download failed: "%~1".
  exit /b 1
)
exit /b 0

:run_python_installer
call :log Running Python installer "%~1".
start "" /wait "%~1" /quiet InstallAllUsers=0 Include_pip=1 Include_tcltk=1 PrependPath=0 Shortcuts=0 TargetDir="%~2" >> "%BOOTSTRAP_LOG%" 2>&1
exit /b %errorlevel%

:python_runtime_usable
if not exist "%~1" exit /b 1
"%~1" -c "import sys; print(sys.version)" >nul 2>&1
exit /b %errorlevel%

:log
set "LOG_MESSAGE=%*"
echo !LOG_MESSAGE!
>> "%BOOTSTRAP_LOG%" echo !LOG_MESSAGE!
exit /b 0

:fail
set "FAIL_MESSAGE=%*"
echo ERROR: !FAIL_MESSAGE!
>> "%BOOTSTRAP_LOG%" echo ERROR: !FAIL_MESSAGE!
exit /b 0

:parse_args
if "%~1"=="" goto :parse_done
if /I "%~1"=="--help" goto :usage
if /I "%~1"=="-h" goto :usage
if /I "%~1"=="--only" (
  set "TARGET=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--pdf-service-name" (
  set "PDF_SERVICE_NAME=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--index-service-name" (
  set "INDEX_SERVICE_NAME=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--pdf-display-name" (
  set "PDF_DISPLAY_NAME=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--index-display-name" (
  set "INDEX_DISPLAY_NAME=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--pdf-description" (
  set "PDF_DESCRIPTION=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--index-description" (
  set "INDEX_DESCRIPTION=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--pdf-interval-seconds" (
  set "PDF_INTERVAL_SECONDS=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--index-interval-seconds" (
  set "INDEX_INTERVAL_SECONDS=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--pdf-image-index-mode" (
  set "PDF_IMAGE_INDEX_MODE=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--config" (
  set "CONFIG_PATH=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--excel" (
  set "EXCEL_PATH=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--index-dir" (
  set "INDEX_DIR_ARGS=%INDEX_DIR_ARGS% --dir ""%~2"""
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--nssm" (
  set "NSSM_EXE=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--python" (
  set "PYTHON_EXE=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--run-user" (
  set "RUN_USER=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--run-password" (
  set "RUN_PASSWORD=%~2"
  shift
  shift
  goto :parse_args
)
if /I "%~1"=="--startup" (
  if /I "%~2"=="manual" (
    set "START_MODE=SERVICE_DEMAND_START"
  ) else (
    set "START_MODE=SERVICE_AUTO_START"
  )
  shift
  shift
  goto :parse_args
)
echo ERROR: Unknown argument %~1
exit /b 1

:parse_done
if defined RUN_USER if not defined RUN_PASSWORD (
  echo ERROR: --run-user requires --run-password.
  exit /b 1
)

if /I not "%TARGET%"=="both" if /I not "%TARGET%"=="pdf" if /I not "%TARGET%"=="index" (
  echo ERROR: --only must be one of: both, pdf, index.
  exit /b 1
)

if /I not "%PDF_IMAGE_INDEX_MODE%"=="refresh" if /I not "%PDF_IMAGE_INDEX_MODE%"=="cache-only" if /I not "%PDF_IMAGE_INDEX_MODE%"=="skip" (
  echo ERROR: --pdf-image-index-mode must be one of: refresh, cache-only, skip.
  exit /b 1
)

set /a _pdf_interval_test=%PDF_INTERVAL_SECONDS% >nul 2>&1
if errorlevel 1 (
  echo ERROR: --pdf-interval-seconds must be a positive integer.
  exit /b 1
)
if %PDF_INTERVAL_SECONDS% LEQ 0 (
  echo ERROR: --pdf-interval-seconds must be greater than 0.
  exit /b 1
)

set /a _index_interval_test=%INDEX_INTERVAL_SECONDS% >nul 2>&1
if errorlevel 1 (
  echo ERROR: --index-interval-seconds must be a positive integer.
  exit /b 1
)
if %INDEX_INTERVAL_SECONDS% LEQ 0 (
  echo ERROR: --index-interval-seconds must be greater than 0.
  exit /b 1
)
exit /b 0

:usage
echo Usage:
echo   install_windows_services.bat [options]
echo.
echo One run should:
echo   1. elevate to Administrator
echo   2. download NSSM if missing
echo   3. install local python_runtime if missing
echo   4. install Python packages from requirements.txt
echo   5. register and start Windows services
echo.
echo Default service names:
echo   %PDF_SERVICE_NAME%
echo   %INDEX_SERVICE_NAME%
echo.
echo Default display names in Services:
echo   %PDF_DISPLAY_NAME%
echo   %INDEX_DISPLAY_NAME%
echo.
echo Options:
echo   --only both^|pdf^|index
echo   --config "C:\path\config.json"
echo   --excel "C:\path\data.xlsx"
echo   --index-dir "C:\path\images"   ^(can be passed multiple times^)
echo   --pdf-interval-seconds 300
echo   --index-interval-seconds 1800
echo   --pdf-image-index-mode refresh^|cache-only^|skip
echo   --pdf-service-name PDSGeneratorPdfService
echo   --index-service-name PDSGeneratorImageIndexService
echo   --pdf-display-name "PDS Generator - PDF Generation Service"
echo   --index-display-name "PDS Generator - Image Index Service"
echo   --python "C:\path\python.exe"
echo   --nssm "C:\path\nssm.exe"
echo   --run-user ".\UserName" --run-password "secret"
echo   --startup auto^|manual
exit /b 1
