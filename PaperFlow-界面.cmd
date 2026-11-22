@echo off
setlocal
call "%~dp0PaperFlow.cmd" gui %*
set "PAPERFLOW_GUI_EXIT=%ERRORLEVEL%"
exit /b %PAPERFLOW_GUI_EXIT%
