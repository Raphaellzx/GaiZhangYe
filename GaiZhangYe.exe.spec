# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['GaiZhangYe\\core\\entrypoints\\start_service.py'],
    pathex=[r'D:\\Coding\\Python\\GaiZhangYe'],
    binaries=[],
    datas=[('GaiZhangYe/web/static', 'static'), ('GaiZhangYe/web/templates', 'templates')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Exclude many large optional libraries to reduce bundle size.
    # Keep pywin32, pymupdf (fitz) and pillow which are required by this project.
    excludes=['PyQt5','PyQt6','PySide6','PySide2','matplotlib','numpy','scipy','pandas',
              'sklearn','skimage','notebook','jupyter','jupyterlab','IPython','sphinx',
              'pyarrow','numba','numexpr','plotly','bokeh','torch','tensorflow','cv2',
              'opencv_python','pyqtgraph','tornado','zmq','qtpy','nbconvert',
              'jupyter_client','ipykernel','jupyter_core','mpl_toolkits','xlrd','openpyxl'],
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
    name='GaiZhangYe.exe',
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
