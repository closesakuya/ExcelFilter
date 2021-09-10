# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['main.py','main_ui.py','imgs.py','Filter.py'],
             pathex=['D:\\share_dir\\product_env\\01.SVN\01.local_git\\ExcelFilter.git','res'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=['pandas','numpy','bs4'],
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
          name='表格条件筛选工具',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          icon='res\\main.ico')
