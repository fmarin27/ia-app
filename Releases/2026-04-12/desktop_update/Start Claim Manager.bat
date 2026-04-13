@echo off
setlocal

set "APP_DIR=%~dp0"
set "PYTHON_EXE="

if exist "C:\Program Files\Python311\python.exe" set "PYTHON_EXE=C:\Program Files\Python311\python.exe"
if not defined PYTHON_EXE if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python311\python.exe"

pushd "%APP_DIR%"
if defined PYTHON_EXE (
  "%PYTHON_EXE%" "app.py"
) else (
  py -3.11 "app.py"
)
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
  echo.
  echo Claim Manager closed with exit code %EXIT_CODE%.
  pause
)

exit /b %EXIT_CODE%
