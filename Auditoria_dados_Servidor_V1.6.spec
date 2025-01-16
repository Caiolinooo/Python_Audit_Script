# -*- mode: python ; coding: utf-8 -*-


a = Analysis(['Auditoria_dados_Servidor_V1.6.py'],
             pathex=[],
             binaries=[],
             datas=[(r'C:\Users\Cliente2\AppData\Local\Programs\Python\Python38\Lib\tkinter', 'tkinter')],
             hiddenimports=['xlsxwriter', 'tkinter', 'pandas', 'tqdm'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=None,
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
    name='Auditoria_dados_Servidor_V1.6',
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
