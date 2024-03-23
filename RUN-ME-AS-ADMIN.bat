@echo off
cd /d "%~dp0"

powershell -Command "Set-ExecutionPolicy Bypass"
powershell -File "./ShellNinja96SystemMetrics.ps1"