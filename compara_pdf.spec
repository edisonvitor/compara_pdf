# spec file for PyInstaller

a = Analysis(
    ['compara_pdf.py'],
    pathex=['J:\\projetos\\scripts\\new-venv\\Lib\\site-packages'],  # Adicione o path das bibliotecas
    binaries=[],
    datas=[],
    hiddenimports=['PyMuPDF', 'PyMuPDFb'],  # Inclua as dependências necessárias
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
    name='compara_pdf',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
