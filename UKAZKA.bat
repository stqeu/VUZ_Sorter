@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)
rem Vygeneruje ukazka_vystupu\ z demo dat (README.md + PREHLAD.md pre GitHub)
%PY% lta_triedenie.py --config config\ukazka.ini --ostro %*
pause
