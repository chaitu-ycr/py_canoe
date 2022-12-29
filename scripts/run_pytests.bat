::==========================================================================
::Script to execute before any other. This scripts installs and updates necessary info.
::==========================================================================
@echo off

title "uploading package to pypi"

set origin_dir=%CD%
set file_dir=%~dp0
cd %file_dir%
cd ..
set root_folder=%CD%
set working_dir=%root_folder%
set python_venv_path=%working_dir%\venv
set python_exe=%python_venv_path%\Scripts\python.exe
set cmd_venv_activate=%root_folder%\venv\Scripts\activate.bat
set cmd_venv_deactivate=%root_folder%\venv\Scripts\deactivate.bat

:: checking if Python path exists

:PYTHON_VENV
title "checking if Python virtual environment exists"
if exist %python_venv_path% (
	title installing pip dependencies...
	echo ----------------------------------- PIP VENV FOUND -----------------------------------
	echo "Python virtual env path '%python_exe%' exists on machine"
	echo.
	echo "Upgrading Python PIP module"
	%python_exe% -m pip install pip --upgrade
	echo.
    echo "installing pytest and pywin32"
	%python_exe% -m pip install pytest, pywin32
	echo "completed installing pytest and pywin32"
    echo.
    echo "activating virtual env"
    call %cmd_venv_activate%
    if %ERRORLEVEL% NEQ 0 (GOTO ERROR)
    echo "activated virtual env"
    echo.
    title "running pytest"
    echo "started running pytest"
    cd %root_folder%\tests
    pytest
    if %ERRORLEVEL% NEQ 0 (GOTO ERROR)
    echo "completed running pytest"
) else (
    GOTO :VENV_ERROR
)

if %ERRORLEVEL% NEQ 0 (GOTO ERROR)
call %cmd_venv_deactivate%
cd %origin_dir%
pause
GOTO :eof

:VENV_ERROR
echo.
echo ----------------------------------- PIP VENV NOT FOUND -----------------------------------
echo '%python_venv_path%' not found
echo Creating virtual environment now
python -m venv %python_venv_path%
GOTO :PYTHON_VENV

:ERROR
title "Failed due to error %ERRORLEVEL%"
cd %origin_dir%
pause
GOTO :eof