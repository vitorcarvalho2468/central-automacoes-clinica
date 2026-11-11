@echo off
cd /d "%~dp0"
echo Gerando o CentralAutomacoes.exe (demora alguns minutos)...
echo (as bibliotecas sao instaladas pelo "Instalar dependencias.bat", nao aqui)
echo.
where python >nul 2>nul
if %errorlevel%==0 (set PY=python) else (set PY=py)

%PY% -c "import PyInstaller, playwright, customtkinter, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Faltam bibliotecas para gerar o .exe.
    echo De dois cliques em "Instalar dependencias.bat", espere terminar e rode este arquivo de novo.
    pause
    exit /b 1
)

if exist build rmdir /s /q build
if exist dist\CentralAutomacoes.exe del /q dist\CentralAutomacoes.exe
%PY% -m PyInstaller --noconfirm --clean CentralAutomacoes.spec
if errorlevel 1 goto erro

echo.
echo ===== Pronto! O programa esta em: dist\CentralAutomacoes.exe =====
explorer dist
pause
exit /b 0

:erro
echo.
echo ===== Deu erro ao gerar o .exe. Veja a mensagem acima. =====
echo Se o erro for "PermissionError" ou "Acesso negado", feche o CentralAutomacoes.exe e tente de novo.
pause
exit /b 1
