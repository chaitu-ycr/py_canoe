@echo off

title "deploying document to GitHub Pages using mkdocs"

set origin_dir=%CD%
set file_dir=%~dp0
pushd %file_dir%
cd ..
set root_folder=%CD%
set cmd_venv_activate=%root_folder%\.venv\Scripts\activate.bat
set cmd_venv_deactivate=%root_folder%\.venv\Scripts\deactivate.bat

cd %root_folder%

:ACTIVATE_VENV
call %cmd_venv_activate%
if %ERRORLEVEL% NEQ 0 (GOTO ERROR)

:START_MKDOCS_SERVER
mkdocs gh-deploy
if %ERRORLEVEL% NEQ 0 (GOTO ERROR)

:END
call %cmd_venv_deactivate%
cd %origin_dir%
pause
GOTO :eof

:ERROR
title "Failed to run mkdocs due to error %ERRORLEVEL%"
cd %origin_dir%
pause
GOTO :eof

popd
