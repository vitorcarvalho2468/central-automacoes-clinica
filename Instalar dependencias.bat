@echo off
cd /d "%~dp0"
echo Instalando as bibliotecas da Central de Automacoes e do Robo SulAmerica (so precisa fazer uma vez)...
echo.
where python >nul 2>nul
if %errorlevel%==0 (set PY=python) else (set PY=py)
%PY% -m pip install --upgrade -r requirements.txt
%PY% -m playwright install chromium
echo.
echo ===== Pronto! Agora abra o "Abrir Central de Automacoes.bat" =====
pause
