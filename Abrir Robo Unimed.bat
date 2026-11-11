@echo off
cd /d "%~dp0"
where python >/dev/null 2>nul
if %errorlevel%==0 (set PY=python) else (set PY=py)
%PY% interface_unimed.py
