@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0\Launch-IntuneExplorer.ps1" -InstallMissing
pause