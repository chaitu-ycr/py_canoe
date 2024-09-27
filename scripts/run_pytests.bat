@echo off

title "running pytest"

set origin_dir=%CD%
set file_dir=%~dp0
pushd %file_dir%
cd ..
set root_folder=%CD%
set python_venv_path=%root_folder%\.venv
set python_exe=%python_venv_path%\Scripts\python.exe
set poetry_exe=%python_venv_path%\Scripts\poetry.exe
set cmd_venv_activate=%python_venv_path%\Scripts\activate.bat
set cmd_venv_deactivate=%python_venv_path%\Scripts\deactivate.bat

cd %root_folder%

:POETRY_SETUP
%python_exe% -m poetry install --only main
if %ERRORLEVEL% NEQ 0 (GOTO ERROR)

:START_PYTESTS
%poetry_exe% run pytest
if %ERRORLEVEL% NEQ 0 (GOTO ERROR)

:END
call %cmd_venv_deactivate%
cd %origin_dir%
pause
GOTO :eof

:ERROR
title "Failed to run pytests due to error %ERRORLEVEL%"
cd %origin_dir%
pause
popd
GOTO :eof

popd
