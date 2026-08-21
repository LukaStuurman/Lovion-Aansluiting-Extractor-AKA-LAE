@echo off
setlocal
cd /d "%~dp0"
start "Lovion Coordinate Workbench" powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -File "%~dp0LovionCoordinateTool.Gui.ps1"
