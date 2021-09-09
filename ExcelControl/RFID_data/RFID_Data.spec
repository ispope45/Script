# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['RFID_xlstosheet_v1.21.py'],
             pathex=['D:\Python\Script\ExcelControl\RFID_data'],
             binaries=[],
             datas=[('Data/classData210909.xml','.'),
             ('Data/sampledata210909.xlsx','.')],
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
          name='RFID_data_v1.01',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )