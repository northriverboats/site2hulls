# -*- mode: python ; coding: utf-8 -*-

# pyinstaller --onefile "site2hulls.fww.spec" site2hulls.py

block_cipher = None


a = Analysis(['site2hulls.py'],
             pathex=['/home/fwarren/.venv/site2hulls'],
             binaries=[],
             datas=[('.env-local','.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='site2hulls',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )
