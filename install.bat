@echo off
setlocal ENABLEDELAYEDEXPANSION

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo =============================================
echo ExcelSuite - Install dependencies (venv)
echo =============================================

set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY (
  where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
  echo [ERROR] Python 3.x not found on PATH.
  echo Python 3.9 이상을 설치한 후 다시 실행하세요.
  pause
  exit /b 1
)

echo Using Python: %PY%
%PY% -V

if not exist ".venv_suite" goto CREATE_VENV
echo [i] Using existing virtual environment: .venv_suite
goto ACTIVATE_VENV

:CREATE_VENV
echo [*] Creating virtual environment (.venv_suite)...
%PY% -m venv .venv_suite
if errorlevel 1 (
  echo [ERROR] Failed to create virtual environment.
  pause
  exit /b 1
)

:ACTIVATE_VENV
echo [*] Activating virtual environment...
call ".venv_suite\Scripts\activate.bat"
if errorlevel 1 (
  echo [ERROR] Failed to activate virtual environment.
  pause
  exit /b 1
)

echo [*] Upgrading pip/setuptools/wheel...
python -m pip install --upgrade pip wheel setuptools
if errorlevel 1 (
  echo [ERROR] Failed to upgrade pip/setuptools/wheel.
  pause
  exit /b 1
)

echo [*] Installing required packages (pillow, lxml, pyinstaller)...
python -m pip install pillow lxml pyinstaller
if errorlevel 1 (
  echo [ERROR] Failed to install required packages.
  pause
  exit /b 1
)

echo.
echo [OK] Environment ready.
echo - run.bat  : 가상환경으로 통합 도구 실행
echo - build.bat: 단일 EXE 빌드 (배포용)
echo.
pause
