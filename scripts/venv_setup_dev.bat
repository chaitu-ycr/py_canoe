@echo off

set file_dir=%~dp0
cd %file_dir%
cd ..
set root_folder=%CD%
set python_venv_path=%root_folder%\.venv
set python_exe=%python_venv_path%\Scripts\python.exe

:PYTHON_VENV
title "creating/updating tool environment..."
if exist %python_venv_path% (
	echo "using '%python_exe%' python."
	echo "upgrade python pip module, install poetry and install repo dependencies..."
	%python_exe% -m pip install pip --upgrade
	%python_exe% -m pip install poetry --upgrade
	%python_exe% -m poetry install
	cd %root_folder%
	echo "completed installing tool dependencies."
) else (
    GOTO :VENV_ERROR
)
GOTO :eof

:VENV_ERROR
echo.
echo '%python_venv_path%' not found
echo "Creating virtual environment now..."
python -m venv %python_venv_path%
echo "completed venv creation."
GOTO :PYTHON_VENV

:ERROR
echo "failed to run extract due to error %ERRORLEVEL%."
cd %file_dir%
pause
