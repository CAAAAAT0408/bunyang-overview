@echo off
chcp 65001 >nul
title 분양개요 자동작성
cd /d "%~dp0"
python "app\control.py" open
if errorlevel 1 pause
timeout /t 4 >nul
