# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置
使用: pyinstaller build.spec
"""

import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['app_ctk.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 如果需要打包默认图片，取消下面注释并调整路径
        # ('图片程序测试/图片程序测试/*.jpg', '图片程序测试'),
        # ('图片程序测试/图片程序测试/*.png', '图片程序测试'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell._writer',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'jieba',
        'jieba.posseg',
        'jieba.analyse',
        'customtkinter',
        'matching',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter.test',
        'unittest',
        'email',
        'http',
        'urllib',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='物料图片匹配系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # True=显示控制台, False=仅GUI窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if Path('icon.ico').exists() else None,
)

# Mac .app 打包（macOS 专用）
app = BUNDLE(
    exe,
    name='物料图片匹配系统.app',
    icon='icon.icns' if Path('icon.icns').exists() else None,
    bundle_identifier='com.yundazhifu.material-image-matcher',
    info_plist={
        'NSHighResolutionCapable': True,
        'CFBundleDisplayName': '物料图片匹配系统',
        'CFBundleName': '物料图片匹配系统',
    },
)
