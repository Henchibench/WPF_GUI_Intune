@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0\IntuneExplorer.ps1" -InstallMissing
pause