# AutoTask.spec
# ─────────────────────────────────────────────────────────────────────────────
# Compilar con:  pyinstaller AutoTask.spec
#
# Genera: dist/AutoTask/AutoTask.exe  (modo --onedir, NO onefile)
#
# IMPORTANTE — modo onedir (carpeta) en lugar de onefile (.exe único):
#   - Streamlit necesita acceder a sus archivos estáticos en disco
#   - onefile extrae todo a %TEMP% en cada ejecución → lento y bloqueado por antivirus
#   - onedir es ~5x más rápido al arrancar y más confiable
#
# Para distribuir: comprime la carpeta dist/AutoTask/ completa en un .zip
# ─────────────────────────────────────────────────────────────────────────────

import sys
import os
from PyInstaller.utils.hooks import collect_all, collect_data_files, copy_metadata

# ── Archivos de tu proyecto ───────────────────────────────────────────────────
datas = [
    ('autotask_app.py',  '.'),
    ('analizador.py',    '.'),
    ('config_tareas.py', '.'),
    ('motor_docx.py',    '.'),
    ('formateadores.py', '.'),
    ('generadores',      'generadores'),
]

binaries = []
hiddenimports = []

# ── Streamlit (el más crítico — necesita todos sus archivos estáticos) ────────
tmp = collect_all('streamlit')
datas          += tmp[0]
binaries       += tmp[1]
hiddenimports  += tmp[2]

# Metadata que Streamlit lee en tiempo de ejecución
datas += copy_metadata('streamlit')
try:
    datas += copy_metadata('streamlit-aggrid')
except Exception:
    pass

# ── Google GenAI ──────────────────────────────────────────────────────────────
# google-genai usa namespaces (google.*) que PyInstaller no detecta solo
tmp = collect_all('google.genai')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]

tmp = collect_all('google.ai')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]

tmp = collect_all('google.api_core')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]

tmp = collect_all('google.auth')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]

try:
    datas += copy_metadata('google-genai')
except Exception:
    pass
try:
    datas += copy_metadata('google-api-core')
except Exception:
    pass

# Imports ocultos críticos de google-genai
hiddenimports += [
    'google.genai',
    'google.genai.types',
    'google.genai.client',
    'google.genai.models',
    'google.genai.chats',
    'google._upb._message',
    'google.protobuf',
    'google.protobuf.descriptor',
    'google.protobuf.descriptor_pool',
    'google.protobuf.reflection',
    'google.protobuf.symbol_database',
    'google.protobuf.message_factory',
    'grpc',
    'grpc._channel',
    'httpx',
    'httpx._transports.default',
    'httpcore',
    'httpcore._sync.http11',
    'httpcore._async.http11',
]

# ── Dependencias de red/HTTP ──────────────────────────────────────────────────
for pkg in ['httpx', 'httpcore', 'anyio', 'h11', 'certifi', 'charset_normalizer']:
    try:
        tmp = collect_all(pkg)
        datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]
    except Exception:
        pass

# ── pypdf ─────────────────────────────────────────────────────────────────────
tmp = collect_all('pypdf')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]
try:
    datas += copy_metadata('pypdf')
except Exception:
    pass

# ── python-docx ──────────────────────────────────────────────────────────────
tmp = collect_all('docx')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]
try:
    datas += copy_metadata('python-docx')
except Exception:
    pass
hiddenimports += ['docx', 'docx.shared', 'docx.oxml', 'docx.oxml.ns', 'docx.enum.text']

# ── lxml (para manipulación OMML de ecuaciones) ───────────────────────────────
tmp = collect_all('lxml')
datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]
try:
    datas += copy_metadata('lxml')
except Exception:
    pass
hiddenimports += ['lxml.etree', 'lxml._elementpath', 'lxml.builder']

# ── pandas / altair / pydeck (requeridos por Streamlit) ──────────────────────
for pkg in ['pandas', 'altair', 'pydeck', 'pyarrow', 'numpy']:
    try:
        tmp = collect_all(pkg)
        datas += tmp[0]; binaries += tmp[1]; hiddenimports += tmp[2]
        datas += copy_metadata(pkg)
    except Exception:
        pass

# ── Otros imports ocultos frecuentes con Streamlit ───────────────────────────
hiddenimports += [
    'streamlit.runtime.scriptrunner',
    'streamlit.runtime.scriptrunner.magic_funcs',
    'streamlit.web.cli',
    'streamlit.components.v1',
    'tornado',
    'tornado.web',
    'tornado.httpserver',
    'tornado.ioloop',
    'tornado.websocket',
    'click',
    'PIL',
    'PIL.Image',
    'pypdf',
    'charset_normalizer',
    'certifi',
    'idna',
    'packaging',
    'importlib_metadata',
    'zipp',
    'toml',
    'validators',
    'watchdog',
    'watchdog.observers',
]

# ── Metadata adicional ────────────────────────────────────────────────────────
for pkg in ['pandas', 'altair', 'pydeck', 'click', 'toml', 'packaging']:
    try:
        datas += copy_metadata(pkg)
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────────────────
a = Analysis(
    ['run.py'],
    pathex=['.'],           # directorio actual
    binaries=binaries,
    datas=datas,
    hiddenimports=list(set(hiddenimports)),   # deduplicar
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Excluir lo que NO necesitas para reducir el tamaño
        'matplotlib', 'scipy', 'sklearn', 'tensorflow', 'torch',
        'IPython', 'jupyter', 'notebook', 'pytest', 'sphinx',
        'tkinter', '_tkinter',
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

# ── MODO ONEDIR (carpeta) — NO onefile ───────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    [],                     # <-- vacío para onedir
    exclude_binaries=True,  # <-- clave para onedir
    name='AutoTask',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=['vcruntime140.dll', 'python*.dll'],
    console=True,           # True = ventana de consola visible (útil para errores)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# COLL es necesario para onedir
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=['vcruntime140.dll', 'python*.dll'],
    name='AutoTask',        # carpeta de salida: dist/AutoTask/
)
