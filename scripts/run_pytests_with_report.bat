@echo off

REM ============================
REM Run Pytest with Coverage and HTML Report
REM ============================

REM Set window title
title Running Pytests with Report

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
REM 2. Sync Dependencies
REM ----------------------------
uv sync --link-mode=copy
if %ERRORLEVEL% NEQ 0 goto ERROR

REM ----------------------------
REM 3. Run Pytest with Reports
REM ----------------------------
uv run pytest tests/ ^
    --html=tests/report/test_reports/full_test_report.html --self-contained-html ^
    --cov=src ^
    --cov-report=html:tests/report/cov/htmlcov ^
    --cov-report=xml:tests/report/cov/coverage.xml ^
    --cov-report=json:tests/report/cov/coverage.json ^
    --maxfail=5 ^
    --tb=short ^
    -n auto
if %ERRORLEVEL% NEQ 0 goto ERROR

REM ----------------------------
REM 4. Deactivate Virtual Environment and Cleanup
REM ----------------------------
call "%VENV_DEACTIVATE%"
popd
goto :EOF

REM ----------------------------
REM Error Handler
REM ----------------------------
:ERROR
title Failed to run pytests due to error %ERRORLEVEL%
popd
pause
goto :EOF
