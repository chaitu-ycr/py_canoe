@echo off

REM ============================
REM Deploy MkDocs Documentation to GitHub Pages
REM ============================

REM Set window title
title Deploying Documentation to GitHub Pages

REM Save original directory and move to project root
set "ORIGIN_DIR=%CD%"
pushd %~dp0
cd ..
set "ROOT_DIR=%CD%"

REM Set venv activation/deactivation commands
set "VENV_ACTIVATE=%ROOT_DIR%\.venv\Scripts\activate.bat"
set "VENV_DEACTIVATE=%ROOT_DIR%\.venv\Scripts\deactivate.bat"

REM ----------------------------
REM 1. Activate Virtual Environment
REM ----------------------------
call "%VENV_ACTIVATE%"
if %ERRORLEVEL% NEQ 0 goto ERROR

REM ----------------------------
REM 2. Deploy Documentation
REM ----------------------------
uv run mkdocs gh-deploy
if %ERRORLEVEL% NEQ 0 goto ERROR

REM ----------------------------
REM 3. Deactivate Virtual Environment and Cleanup
REM ----------------------------
call "%VENV_DEACTIVATE%"
popd
cd "%ORIGIN_DIR%"
goto :EOF

REM ----------------------------
REM Error Handler
REM ----------------------------
:ERROR
echo Failed to deploy documentation due to error %ERRORLEVEL%
call "%VENV_DEACTIVATE%"
popd
cd "%ORIGIN_DIR%"
pause
goto :EOF