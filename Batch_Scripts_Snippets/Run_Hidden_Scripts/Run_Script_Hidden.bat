@echo off
set curr_dir=%~dp0
set launch_dir=%cd%
echo %curr_dir%
:: Use the script directory for path
:: pushd "%~dp0"
cscript "%cd%\invisible.vbs" "%cd%\Set_All_Files_Dirs_Attr.cmd
wscript "%cd%\invisible.vbs" "%cd%\Set_All_Files_Dirs_Attr.bat
pause
