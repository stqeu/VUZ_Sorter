@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)
rem Iny projekt:  OSTRO.bat --config config\rhc.local.ini
%PY% lta_triedenie.py --config config\demo.ini --ostro %*
pause
