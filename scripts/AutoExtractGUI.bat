@echo off
chcp 65001 >nul
title AutoExtract GUI
rem Double-click to launch the GUI version.
start "" powershell -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0AutoExtractGUI.ps1"
