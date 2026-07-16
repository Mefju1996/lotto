# lotto.spec
import sys
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

a = Analysis(
    ['scripts/lotto_generator_tkinter.py'],
    pathex=['.'],
    binaries=[
        (sys.executable, '.'),   # kopiuje python.exe obok LottoGenerator.exe
    ],
    datas=[
        ('scripts/generate_lotto_stats_final.py', 'scripts'),
        ('scripts/generate_lotto_stats.py',        'scripts'),
        ('scripts/scraper_megalotto.py',            'scripts'),
        *collect_data_files('openpyxl'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'openpyxl.styles.stylesheet',
        'sqlite3',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='LottoGenerator',
    debug=False,
    strip=False,
    upx=True,
    console=False,
    icon=None,
)
