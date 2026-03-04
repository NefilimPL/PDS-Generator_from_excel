@echo off
setlocal EnableExtensions

for %%I in ("%~dp0..") do set "PROJECT_DIR=%%~fI"
set "SERVICES_DIR=%~dp0"
set "TARGET=both"
set "PDF_SERVICE_NAME=PDSGeneratorPdfService"
set "INDEX_SERVICE_NAME=PDSGeneratorImageIndexService"
set "NSSM_EXE="

call :parse_args %*
if errorlevel 1 exit /b 1

call :ensure_elevated %*
if errorlevel 2 exit /b 0
if errorlevel 1 exit /b 1

call :resolve_nssm
if errorlevel 1 exit /b 1

if /I not "%TARGET%"=="index" (
  call :remove_service "%PDF_SERVICE_NAME%"
  if errorlevel 1 exit /b 1
)

if /I not "%TARGET%"=="pdf" (
  call :remove_service "%INDEX_SERVICE_NAME%"
  if errorlevel 1 exit /b 1
)

echo Done.
exit /b 0

:remove_service
set "CURRENT_SERVICE_NAME=%~1"
sc query "%CURRENT_SERVICE_NAME%" >nul 2>&1
if errorlevel 1 (
  echo Service "%CURRENT_SERVICE_NAME%" is not installed.
  exit /b 0
)

echo Stopping service "%CURRENT_SERVICE_NAME%"...
"%NSSM_EXE%" stop "%CURRENT_SERVICE_NAME%" >nul 2>&1

echo Removing service "%CURRENT_SERVICE_NAME%"...
"%NSSM_EXE%" remove "%CURRENT_SERVICE_NAME%" confirm
if errorlevel 1 (
  echo ERROR: Failed to remove service "%CURRENT_SERVICE_NAME%".
  exit /b 1
)

echo Service "%CURRENT_SERVICE_NAME%" has been removed.
exit /b 0

:resolve_nssm
if defined NSSM_EXE (
  if exist "%NSSM_EXE%" exit /b 0
  echo ERROR: NSSM not found at "%NSSM_EXE%".
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
echo ERROR: Could not find nssm.exe. Put it in "%SERVICES_DIR%", in the project root, in PATH, or pass --nssm.
exit /b 1

:ensure_elevated
fltmc >nul 2>&1
if not errorlevel 1 exit /b 0
echo Requesting administrator privileges...
set "PDS_UNINSTALL_ARGS=%*"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process -FilePath '%~f0' -WorkingDirectory '%CD%' -Verb RunAs -ArgumentList $env:PDS_UNINSTALL_ARGS"
if errorlevel 1 (
  echo ERROR: Administrator privileges are required.
  exit /b 1
)
exit /b 2

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
if /I "%~1"=="--nssm" (
  set "NSSM_EXE=%~2"
  shift
  shift
  goto :parse_args
)
echo ERROR: Unknown argument %~1
exit /b 1

:parse_done
if /I not "%TARGET%"=="both" if /I not "%TARGET%"=="pdf" if /I not "%TARGET%"=="index" (
  echo ERROR: --only must be one of: both, pdf, index.
  exit /b 1
)
exit /b 0

:usage
echo Usage:
echo   uninstall_windows_services.bat [options]
echo.
echo Options:
echo   --only both^|pdf^|index
echo   --pdf-service-name PDSGeneratorPdfService
echo   --index-service-name PDSGeneratorImageIndexService
echo   --nssm "C:\path\nssm.exe"
exit /b 1
