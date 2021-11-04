# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['RFID_TagData.py'],
             pathex=['D:\Python\Script\ExcelControl\RFID_data'],
             binaries=[],
             datas=[('Data/sampleTagData.xlsx','./Data/')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=True)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='RFID_TagData_v1.00',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )