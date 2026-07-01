# run.py — Launcher de AutoTask
# ─────────────────────────────────────────────────────────────────────────────
# Este archivo es el punto de entrada compilado por PyInstaller.
# Funciona en modo --onedir (carpeta _internal/).
#
# Qué hace:
#   1. Detecta si está corriendo como .exe (sys._MEIPASS) o como script normal
#   2. Agrega el directorio de recursos al sys.path
#   3. Abre el navegador automáticamente
#   4. Lanza streamlit run autotask_app.py con configuración silenciosa
# ─────────────────────────────────────────────────────────────────────────────

import os
import sys
import time
import threading
import webbrowser
import subprocess

def get_base_dir():
    """Devuelve la carpeta raíz de recursos, tanto en .exe como en script."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Corriendo como .exe compilado — recursos en _MEIPASS
        return sys._MEIPASS
    # Corriendo como script normal en desarrollo
    return os.path.dirname(os.path.abspath(__file__))

def get_python_exe():
    """Devuelve la ruta al intérprete Python a usar."""
    if getattr(sys, 'frozen', False):
        # Dentro del .exe, sys.executable es el propio .exe
        # Pero necesitamos el python.exe embebido en _internal/
        base = sys._MEIPASS
        # PyInstaller onedir pone python3.dll y pythonXX.dll aquí
        # El ejecutable python real está junto al .exe
        exe_dir = os.path.dirname(sys.executable)
        for candidate in [
            os.path.join(exe_dir, '_internal', 'python.exe'),
            os.path.join(exe_dir, 'python.exe'),
            sys.executable,
        ]:
            if os.path.exists(candidate):
                return candidate
        return sys.executable
    return sys.executable

def abrir_navegador(puerto: int, retraso: float = 2.5):
    """Abre el navegador después de un breve retraso."""
    def _abrir():
        time.sleep(retraso)
        webbrowser.open(f"http://localhost:{puerto}")
    t = threading.Thread(target=_abrir, daemon=True)
    t.start()

def main():
    base_dir = get_base_dir()
    app_path = os.path.join(base_dir, 'autotask_app.py')
    puerto   = 8501

    # Verificar que el app existe
    if not os.path.exists(app_path):
        print(f"ERROR: No se encontró autotask_app.py en:\n  {app_path}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

    # Agregar base_dir al path para que los imports funcionen
    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)

    # Configurar variable de entorno para que Streamlit encuentre sus archivos
    os.environ['STREAMLIT_BROWSER_GATHER_USAGE_STATS'] = 'false'
    os.environ['STREAMLIT_SERVER_HEADLESS']             = 'true'
    os.environ['STREAMLIT_SERVER_PORT']                 = str(puerto)
    os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'
    os.environ['STREAMLIT_GLOBAL_DEVELOPMENT_MODE']     = 'false'
    # Evitar que Streamlit busque config en ~\.streamlit cuando corre como .exe
    os.environ['STREAMLIT_CONFIG_DIR'] = base_dir

    # Abrir el navegador en paralelo
    abrir_navegador(puerto)

    # ── Estrategia de lanzamiento ────────────────────────────────────────────
    # Intentamos primero lanzar streamlit desde dentro del mismo proceso
    # (más limpio, sin subprocess). Si falla, usamos subprocess como respaldo.
    try:
        # Agregar rutas de paquetes empaquetados al path
        for subdir in ['Lib', 'lib', os.path.join('Lib', 'site-packages'),
                       os.path.join('lib', 'site-packages')]:
            candidate = os.path.join(base_dir, subdir)
            if os.path.isdir(candidate) and candidate not in sys.path:
                sys.path.insert(0, candidate)

        from streamlit.web import cli as stcli

        # Reemplazar argv para que Streamlit lo lea correctamente
        sys.argv = [
            'streamlit', 'run', app_path,
            f'--server.port={puerto}',
            '--server.headless=true',
            '--server.enableXsrfProtection=false',
            '--browser.gatherUsageStats=false',
            '--global.developmentMode=false',
        ]
        sys.exit(stcli.main())

    except ImportError:
        # Respaldo: lanzar como subprocess usando el Python del sistema
        print("Lanzando Streamlit como proceso externo...")
        python_exe = get_python_exe()
        cmd = [
            python_exe, '-m', 'streamlit', 'run', app_path,
            f'--server.port={puerto}',
            '--server.headless=true',
            '--server.enableXsrfProtection=false',
            '--browser.gatherUsageStats=false',
            '--global.developmentMode=false',
        ]
        proc = subprocess.Popen(
            cmd,
            cwd=base_dir,
            env={**os.environ, 'PYTHONPATH': base_dir},
        )
        try:
            proc.wait()
        except KeyboardInterrupt:
            proc.terminate()
        sys.exit(proc.returncode)

if __name__ == '__main__':
    main()
