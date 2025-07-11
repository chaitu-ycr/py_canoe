@echo off

REM ============================
REM Build Python Wheel Package
REM ============================

REM Set window title
title Building Python Wheel Package

REM Move to script directory and then project root
pushd %~dp0
cd ..

REM Set venv activation/deactivation commands
set "VENV_ACTIVATE=%CD%\.venv\Scripts\activate.bat"
set "VENV_DEACTIVATE=%CD%\.venv\Scripts\deactivate.bat"

REM ----------------------------
REM 1. Activate Virtual Environment
REM ----------------------------
call "%VENV_ACTIVATE%"
if %ERRORLEVEL% NEQ 0 goto ERROR

REM ----------------------------
REM 2. Build the Wheel Package
REM ----------------------------
uv build
if %ERRORLEVEL% NEQ 0 goto ERROR

REM ----------------------------
REM 3. Deactivate Virtual Environment and Cleanup
REM ----------------------------
call "%VENV_DEACTIVATE%"
popd
goto :EOF

REM ----------------------------
REM Error Handler
REM ----------------------------
:ERROR
title Failed to build wheel package due to error %ERRORLEVEL%
popd
pause
goto :EOF
