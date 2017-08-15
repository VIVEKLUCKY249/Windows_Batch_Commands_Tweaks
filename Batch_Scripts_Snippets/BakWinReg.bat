@echo off
setlocal
if not exist "WinRegBak" md "%cd%\WinRegBak"
for %%k in (lm cu cr u cc) do call :ExpReg %%k
goto :eof
:ExpReg
::reg.exe export hk%1 hk%1.reg > nul
reg export hk%1 "%cd%\WinRegBak\hk%1.reg" > nul
if "%errorlevel%"=="1" (
	echo ^>^> Export --hk%1-- Failed.
) else (
	echo ^>^> Export --hk%1-- Fine.
)
pause > nul
::goto :eof
endlocal