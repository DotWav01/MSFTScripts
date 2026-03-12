@echo off
start /b /wait powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Install-TSScanServer.ps1"
