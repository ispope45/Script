# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['TGtoBluemax.py'],
             pathex=['D:\\Python\\Script\\ExcelControl\\TGtoBLUEMAX'],
             binaries=[],
             datas=[('src/BLUEMAX_2.5.3_pol.xlsx','./src/'),
             ('src/obj.csv','./src/')],
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
          name='TGtoBluemax_v2.2',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )
