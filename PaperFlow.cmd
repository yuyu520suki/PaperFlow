@echo off
setlocal
if not defined PAPERFLOW_PYTHON if /I "%CONDA_DEFAULT_ENV%"=="paperflow" if exist "%CONDA_PREFIX%\python.exe" set "PAPERFLOW_PYTHON=%CONDA_PREFIX%\python.exe"
if not defined PAPERFLOW_PYTHON set "PAPERFLOW_PYTHON=%USERPROFILE%\.conda\envs\paperflow\python.exe"
if not exist "%PAPERFLOW_PYTHON%" (
  echo PaperFlow Conda environment was not found. See environment.yml and README.md. 1>&2
  exit /b 1
)
for %%I in ("%PAPERFLOW_PYTHON%") do set "PAPERFLOW_ENV=%%~dpI"
set "PATH=%PAPERFLOW_ENV%;%PAPERFLOW_ENV%Scripts;%PAPERFLOW_ENV%Library\bin;%PATH%"
set "PYTHONNOUSERSITE=1"
pushd "%~dp0" || exit /b 1
"%PAPERFLOW_PYTHON%" -s -X utf8 -m paperflow %*
set "PAPERFLOW_EXIT=%ERRORLEVEL%"
popd
exit /b %PAPERFLOW_EXIT%
