@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0hermes-sidecar.ps1" %*
