"""
Onde fica o navegador (Chromium) que os robôs usam, e instalação automática se ele faltar.

Rodando pelo .py, o Playwright já procura o navegador na pasta certa do usuário.
Dentro do CentralAutomacoes.exe ele procuraria numa pasta temporária do .exe (onde não há
navegador nenhum) e dava o erro "Executable doesn't exist ... _MEI...". Aqui a gente aponta
sempre para a pasta normal do usuário:  %LOCALAPPDATA%\\ms-playwright
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def pasta_navegadores() -> Path:
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Caches"
    else:
        base = Path.home() / ".cache"
    return base / "ms-playwright"


def preparar():
    """Chamar antes de abrir o navegador. No .exe, força a pasta normal dos navegadores."""
    if getattr(sys, "frozen", False) or os.environ.get("PLAYWRIGHT_BROWSERS_PATH") == "0":
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(pasta_navegadores())


preparar()


def falta_navegador(erro: Exception) -> bool:
    txt = str(erro)
    return "Executable doesn't exist" in txt or "playwright install" in txt


def instalar_chromium(log=print):
    """Baixa o Chromium do Playwright (uma vez só; uns 150 MB). Funciona também dentro do .exe."""
    from playwright._impl._driver import compute_driver_executable, get_driver_env

    preparar()
    log("AVISO O navegador do robô não está instalado neste computador. Baixando agora (só da primeira "
        "vez, pode levar alguns minutos)...")
    node, cli = compute_driver_executable()
    flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
    proc = subprocess.run([node, cli, "install", "chromium"], env=get_driver_env(), capture_output=True,
                          text=True, creationflags=flags)
    if proc.returncode != 0:
        detalhe = (proc.stderr or proc.stdout or "").strip().splitlines()[-1:] or ["sem detalhes"]
        raise RuntimeError("Não consegui baixar o navegador do robô (" + detalhe[0][:150] + "). Confira a "
                           "internet ou rode o \"Instalar dependencias.bat\".")
    log("Navegador instalado.")
