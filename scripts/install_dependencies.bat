::==========================================================================
::Script to execute before any other. This scripts installs and updates necessary info.
::==========================================================================
@echo off

set origin_dir=%CD%
set file_dir=%~dp0
cd %file_dir%
cd ..
set root_folder=%CD%
set working_dir=%root_folder%
set python_venv_path=%working_dir%\venv
set python_exe=%python_venv_path%\Scripts\python.exe

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
	echo "installing/upgrading pip dependencies"
	%python_exe% -m pip install  -r %root_folder%\requirements.txt --upgrade
	echo "completed installing ConTest and project pip dependencies..."
) else (
    GOTO :VENV_ERROR
)

if %ERRORLEVEL% NEQ 0 (GOTO ERROR)
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