@echo off

REM ============================
REM Create or Update Python Virtual Environment and Install Dependencies
REM ============================

title Creating/Updating Tool Environment...

REM Move to script directory and then project root
pushd %~dp0
cd ..

REM Set venv path and python executable
set "VENV_PATH=%CD%\.venv"
set "PYTHON_EXE=%VENV_PATH%\Scripts\python.exe"

REM ----------------------------
REM 1. Check if venv exists, else create it
REM ----------------------------
:CHECK_VENV
if exist "%VENV_PATH%" (
    echo Using '%PYTHON_EXE%' python.
    "%PYTHON_EXE%" --version
    echo Upgrading pip, installing uv, and syncing dependencies...
    "%PYTHON_EXE%" -m pip install --upgrade pip
    "%PYTHON_EXE%" -m pip install --upgrade uv
    "%PYTHON_EXE%" -m uv sync --link-mode=copy
    echo Completed installing tool dependencies.
    popd
    goto :EOF
) else (
    echo.
    echo Virtual environment not found at '%VENV_PATH%'
    echo Creating virtual environment now...
    python --version
    python -m venv "%VENV_PATH%"
    if %ERRORLEVEL% NEQ 0 goto ERROR
    echo Completed venv creation.
    goto CHECK_VENV
)

REM ----------------------------
REM Error Handler
REM ----------------------------
:ERROR
echo Failed to setup virtual environment due to error %ERRORLEVEL%.
popd
pause
goto :EOF