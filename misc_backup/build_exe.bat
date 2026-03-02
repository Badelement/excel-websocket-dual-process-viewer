@echo off
setlocal
chcp 65001 >nul

set "LOG=build.log"
set "OUT=python_research.exe"
set "VENV=.venv-build"
set "PY_CMD="

echo [INFO] Build started at %date% %time% > "%LOG%"
echo [INFO] Selecting Python version... >> "%LOG%"

py -3.12 -c "import sys; print(sys.version)" >> "%LOG%" 2>&1
if not errorlevel 1 (
  set "PY_CMD=py -3.12"
  set "VENV=.venv-build-312"
  echo [INFO] Using Python 3.12. >> "%LOG%"
) else (
  py -3.13 -c "import sys; print(sys.version)" >> "%LOG%" 2>&1
  if not errorlevel 1 (
    set "PY_CMD=py -3.13"
    set "VENV=.venv-build-313"
    echo [WARN] Python 3.12 not found, fallback to Python 3.13. >> "%LOG%"
  ) else (
    echo [ERROR] Neither Python 3.12 nor 3.13 is available. >> "%LOG%"
    echo [ERROR] Please install Python 3.12 ^(recommended^) or Python 3.13 ^(64-bit^), then run again.
    pause
    exit /b 1
  )
)

if "%PY_CMD%"=="" (
  echo [ERROR] Failed to select Python runtime. >> "%LOG%"
  pause
  exit /b 1
)

if not exist "%VENV%\Scripts\python.exe" (
  echo [INFO] Creating virtual environment: %VENV% >> "%LOG%"
  %PY_CMD% -m venv "%VENV%" >> "%LOG%" 2>&1
  if errorlevel 1 (
    echo [ERROR] Failed to create venv. >> "%LOG%"
    echo [ERROR] Failed to create venv, see %LOG%.
    pause
    exit /b 1
  )
)

echo [INFO] Installing build dependencies... >> "%LOG%"
"%VENV%\Scripts\python.exe" -m pip install --no-cache-dir -U pip >> "%LOG%" 2>&1
"%VENV%\Scripts\python.exe" -m pip install --no-cache-dir -U pyinstaller pyqt5 pandas openpyxl websockets >> "%LOG%" 2>&1
if errorlevel 1 (
  echo [ERROR] Failed to install build dependencies. >> "%LOG%"
  echo [ERROR] Failed to install build dependencies, see %LOG%.
  pause
  exit /b 1
)

echo [INFO] Running PyInstaller build... >> "%LOG%"
"%VENV%\Scripts\python.exe" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --noconsole ^
  --name python_research ^
  --add-data "data.xlsx;." ^
  app.py >> "%LOG%" 2>&1

if errorlevel 1 (
  echo [ERROR] Build failed, see %LOG%.
  echo.
  echo Last 80 lines of %LOG%:
  powershell -NoProfile -Command "Get-Content -LiteralPath '%cd%\%LOG%' -Tail 80"
  pause
  exit /b 1
)

if exist "dist\%OUT%" (
  copy /Y "dist\%OUT%" "%OUT%" >nul
)

echo [OK] Build success. Output: %cd%\%OUT%
echo [OK] Build success. Output: %cd%\%OUT% >> "%LOG%"
echo [INFO] Log file: %cd%\%LOG%
pause
endlocal
