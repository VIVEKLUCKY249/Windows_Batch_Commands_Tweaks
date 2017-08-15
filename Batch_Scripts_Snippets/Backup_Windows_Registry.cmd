@echo off
setlocal
set BackupFolder=%cd%\WinRegistryBak
if not exist "%BackupFolder%" md "%BackupFolder%"
for %%a in (HKLM HKCU HKCR HKU HKCC) do (
	echo Exporting %%a to %BackupFolder%\%%a.reg ...
	%Systemroot%\system32\reg.exe export "%%a" "%BackupFolder%\%%a.reg" /y
)
pause > nul
exit /b