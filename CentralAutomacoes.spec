# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []
for pacote in ('customtkinter', 'playwright', 'reportlab', 'pdfplumber', 'pdfminer'):
    tmp_ret = collect_all(pacote)
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Robô SulAmérica (importado só quando o botão é clicado)
hiddenimports += ['interface_robo', 'robo_sulamerica', 'gerador_agenda', 'conferencia_faturamento',
                  'pypdf', 'PIL.Image',
                  'interface_porto', 'robo_porto', 'guia_porto_pdf', 'navegador',
                  'interface_unimed', 'robo_unimed',
                  'interface_uniao', 'uniao', 'xlsx_seguro']


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='CentralAutomacoes',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
