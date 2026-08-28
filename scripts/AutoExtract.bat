@echo off
chcp 65001 >nul
setlocal
title AutoExtract
rem Drag & drop files/folders onto this script, or double-click and type a path.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0AutoExtract.ps1" %*
echo.
pause
