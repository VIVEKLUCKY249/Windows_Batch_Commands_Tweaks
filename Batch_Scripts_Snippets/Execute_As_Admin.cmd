@ECHO OFF
NET SESSION
SET SCRIPT_DIR=%~dp0
IF %ERRORLEVEL% NEQ 0 GOTO ELEVATE
GOTO ADMINTASKS

:ELEVATE
CD /d %~dp0
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
EXIT

:ADMINTASKS
:: (Do whatever you need to do here)
CD /d %SCRIPT_DIR%
CMD /K
EXIT