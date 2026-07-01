# autotask_app.py
import streamlit as st
import os
import threading
import time
from analizador import (
    ProviderManager,
    crear_cliente_gemini,
    crear_cliente_groq,
    crear_cliente_cohere,
    ejecutar_flujo_completo,
    convertir_docx_a_pdf,
    convertir_imagen_a_pdf,
)

st.set_page_config(page_title="AUTO-TASK GENERADOR", layout="wide")
st.title("🤖 AUTO-TASK GENERADOR")

# ══════════════════════════════════════════════════════════════════════════════
# INICIALIZACIÓN DE ESTADOS
# ══════════════════════════════════════════════════════════════════════════════
for key, default in {
    # Gemini
    'api_key': "",
    'api_key_backup': "",  # <-- RESPALDO
    # Groq
    'groq_api_key': "",
    'groq_model': "llama-3.1-70b-versatile",
    # Cohere
    'cohere_api_key': "",
    'cohere_model': "command-r",
    # Comunes
    'modelo_actual': "gemini-3.1-flash-lite",
    'gemini_temperature': 0.7,          # <-- AÑADIDO
    'gemini_thinking_budget': 0,        # <-- AÑADIDO (0 significa desactivado)
    'debug_mode': True,
    'logs': [],
    'archivo_listo': None,
    'evaluacion': None,
    'generando': False,
    'stop_requested': False,
    'generar_excel': True,
    'generar_ppt': True,
    'hilo_activo': None,
    'resultado_hilo': None,
    'error_hilo': None,
    'hilo_terminado': False,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

def agregar_log(mensaje):
    st.session_state.logs.append(mensaje)
    if len(st.session_state.logs) > 100:
        st.session_state.logs.pop(0)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.subheader("⚙️ Configuración")

    # ──────────────── GEMINI ────────────────────────────────────────────────
    st.markdown("### 🔵 Gemini")
    api_key_input = st.text_input(
        "API Key",
        type="password",
        value=st.session_state.api_key,
        help="Obtén tu clave gratis en https://aistudio.google.com/app/apikey"
    )
    # CAMPO PARA API DE RESPALDO
    api_key_backup_input = st.text_input(
        "API Key (Respaldo)",
        type="password",
        value=st.session_state.api_key_backup,
        help="Clave secundaria de Gemini para evitar interrupciones por cuota."
    )
    modelo_gemini = st.selectbox(
        "Modelo",
        ('gemini-3.1-flash-lite', 'gemini-3.5-flash', 'gemini-2.5-flash'),
        index=('gemini-3.1-flash-lite', 'gemini-3.5-flash', 'gemini-2.5-flash').index(st.session_state.modelo_actual)
            if st.session_state.modelo_actual in ('gemini-3.1-flash-lite', 'gemini-3.5-flash', 'gemini-2.5-flash')
            else 1
    )
    # --- NUEVA SECCIÓN DE PARÁMETROS AVANZADOS ---
    st.markdown("#### 🎛️ Parámetros Avanzados")
    temp_val = st.slider(
        "Temperatura de Gemini",
        min_value=0.0,
        max_value=2.0,
        value=st.session_state.gemini_temperature,
        step=0.1,
        help="Valores bajos (0.1 - 0.4) para mayor precisión matemática y técnica. Valores altos (1.0 - 1.5) para mayor creatividad y soltura."
    )
    if temp_val != st.session_state.gemini_temperature:
        st.session_state.gemini_temperature = temp_val

    thinking_opt = st.selectbox(
        "Limitador de Razonamiento",
        options=["Desactivado", "Bajo (1024 tokens)", "Medio (2048 tokens)", "Alto (4096 tokens)"],
        index=["Desactivado", "Bajo (1024 tokens)", "Medio (2048 tokens)", "Alto (4096 tokens)"].index(
            "Desactivado" if st.session_state.gemini_thinking_budget == 0 else
            "Bajo (1024 tokens)" if st.session_state.gemini_thinking_budget == 1024 else
            "Medio (2048 tokens)" if st.session_state.gemini_thinking_budget == 2048 else
            "Alto (4096 tokens)"
        ),
        help="Habilita el modo de razonamiento lógico profundo para modelos de última generación compatibles."
    )

    # Mapeo de selección a tokens de presupuesto
    budget_mapping = {
        "Desactivado": 0,
        "Bajo (1024 tokens)": 1024,
        "Medio (2048 tokens)": 2048,
        "Alto (4096 tokens)": 4096
    }
    st.session_state.gemini_thinking_budget = budget_mapping[thinking_opt]
    # ----------------------------------------------
    if api_key_input != st.session_state.api_key:
        st.session_state.api_key = api_key_input
        st.session_state.stop_requested = False
    # ── FIX: antes este campo nunca se guardaba en session_state, por lo que
    # la clave de respaldo se perdía y el sistema abortaba al agotarse la
    # clave principal con "No se configuró ninguna clave de respaldo".
    if api_key_backup_input != st.session_state.api_key_backup:
        st.session_state.api_key_backup = api_key_backup_input
        st.session_state.stop_requested = False
    if modelo_gemini != st.session_state.modelo_actual:
        st.session_state.modelo_actual = modelo_gemini
        st.session_state.stop_requested = False

    st.divider()

    # ──────────────── SELECCIÓN DE ARCHIVOS GENERABLES ─────────────────────
    st.markdown("### 📁 Archivos a generar")
    generar_word_opt = st.checkbox("📄 Word (obligatorio)", value=True, disabled=True)
    generar_excel_opt = st.checkbox("📊 Excel", value=st.session_state.generar_excel)
    generar_ppt_opt = st.checkbox("🔶 PowerPoint", value=st.session_state.generar_ppt)

    # Actualizar estado
    if generar_excel_opt != st.session_state.generar_excel:
        st.session_state.generar_excel = generar_excel_opt
    if generar_ppt_opt != st.session_state.generar_ppt:
        st.session_state.generar_ppt = generar_ppt_opt

    st.divider()

    # ──────────────── GROQ (desplegable = interruptor) ──────────────────────
    # Al hacer clic en el título se expande/colapsa el panel, y ese mismo
    # estado (expandido/colapsado) es lo que activa o desactiva el proveedor.
    # Colapsado = Groq NO se usa aunque tenga una API Key guardada.
    with st.expander("🟠 Groq — clic para activar / desactivar", key="groq_expander_open"):
        groq_api_input = st.text_input(
            "API Key (Groq)",
            type="password",
            value=st.session_state.groq_api_key,
            help="Obtén tu clave en https://console.groq.com/keys"
        )
        groq_model_select = st.selectbox(
            "Modelo Groq",
            ('llama-3.1-70b-versatile', 'mixtral-8x7b-32768', 'gemma2-9b-it'),
            index=('llama-3.1-70b-versatile', 'mixtral-8x7b-32768', 'gemma2-9b-it').index(st.session_state.groq_model)
                if st.session_state.groq_model in ('llama-3.1-70b-versatile', 'mixtral-8x7b-32768', 'gemma2-9b-it')
                else 0
        )
        if groq_api_input != st.session_state.groq_api_key:
            st.session_state.groq_api_key = groq_api_input
            st.session_state.stop_requested = False
        if groq_model_select != st.session_state.groq_model:
            st.session_state.groq_model = groq_model_select
            st.session_state.stop_requested = False

    groq_activo = st.session_state.get("groq_expander_open", False)
    if groq_activo and st.session_state.groq_api_key:
        st.caption("✅ Groq activo")
    elif groq_activo and not st.session_state.groq_api_key:
        st.caption("⚠️ Groq abierto, pero falta la API Key")
    else:
        st.caption("⏸️ Groq desactivado (clic en el título para activar)")

    st.divider()

    # ──────────────── COHERE (desplegable = interruptor) ────────────────────
    with st.expander("🟣 Cohere — clic para activar / desactivar", key="cohere_expander_open"):
        cohere_api_input = st.text_input(
            "API Key (Cohere)",
            type="password",
            value=st.session_state.cohere_api_key,
            help="Obtén tu clave en https://dashboard.cohere.com/api-keys"
        )
        cohere_model_select = st.selectbox(
            "Modelo Cohere",
            ('command-r', 'command-r-plus', 'command-light'),
            index=('command-r', 'command-r-plus', 'command-light').index(st.session_state.cohere_model)
                if st.session_state.cohere_model in ('command-r', 'command-r-plus', 'command-light')
                else 0
        )
        if cohere_api_input != st.session_state.cohere_api_key:
            st.session_state.cohere_api_key = cohere_api_input
            st.session_state.stop_requested = False
        if cohere_model_select != st.session_state.cohere_model:
            st.session_state.cohere_model = cohere_model_select
            st.session_state.stop_requested = False

    cohere_activo = st.session_state.get("cohere_expander_open", False)
    if cohere_activo and st.session_state.cohere_api_key:
        st.caption("✅ Cohere activo")
    elif cohere_activo and not st.session_state.cohere_api_key:
        st.caption("⚠️ Cohere abierto, pero falta la API Key")
    else:
        st.caption("⏸️ Cohere desactivado (clic en el título para activar)")

    st.divider()

    # ──────────────── Configuración común ──────────────────────────────────
    st.session_state.debug_mode = st.checkbox(
        "🔍 Modo Depuración (ver logs)",
        value=st.session_state.debug_mode
    )
    if st.button("🗑️ Limpiar logs"):
        st.session_state.logs = []
        st.session_state.stop_requested = False

# ══════════════════════════════════════════════════════════════════════════════
# ÁREA PRINCIPAL — subida de archivos e instrucciones
# ══════════════════════════════════════════════════════════════════════════════
archivos_cargados = st.file_uploader(
    "Arrastra aquí tus guías, PDFs o imágenes (máx. 10 archivos)",
    type=['pdf', 'png', 'jpg', 'jpeg', 'webp', 'gif','docx', 'doc'],
    accept_multiple_files=True
)

# ─── SOPORTE PARA PEGAR IMÁGENES (Ctrl + V) ───
st.components.v1.html(
    """
    <script>
    const parentDoc = window.parent.document;

    // Evitamos registrar múltiples listeners idénticos en cada recarga de Streamlit
    if (!window.parent.__pasteListenerAttached) {
        window.parent.__pasteListenerAttached = true;

        parentDoc.addEventListener('paste', (e) => {
            const files = e.clipboardData.files;
            if (files && files.length > 0) {
                // Validamos si alguno de los elementos del portapapeles es una imagen
                let tieneImagen = false;
                for (let i = 0; i < files.length; i++) {
                    if (files[i].type.startsWith('image/')) {
                        tieneImagen = true;
                        break;
                    }
                }

                // Si contiene una imagen, la asignamos al st.file_uploader
                if (tieneImagen) {
                    const fileInput = parentDoc.querySelector('input[type="file"]');
                    if (fileInput) {
                        const dataTransfer = new DataTransfer();
                        for (let i = 0; i < files.length; i++) {
                            dataTransfer.items.add(files[i]);
                        }
                        fileInput.files = dataTransfer.files;

                        // Despachamos el evento para notificar a React/Streamlit
                        const event = new Event('change', { bubbles: true });
                        fileInput.dispatchEvent(event);
                    }
                }
            }
        });
    }
    </script>
    """,
    height=0,
    width=0,
)
# =======================control + V para pegar imágenes=======================

instrucciones = st.text_area("Instrucciones para la IA:", height=100)

# ══════════════════════════════════════════════════════════════════════════════
# COLUMNAS: Botón Generar + Botón Cancelar
# ══════════════════════════════════════════════════════════════════════════════
col1, col2 = st.columns(2)

with col1:
    generar_clicked = st.button("🚀 Generar Tarea", use_container_width=True, disabled=st.session_state.generando)

with col2:
    cancelar_clicked = st.button("⏹️ Cancelar", use_container_width=True, disabled=not st.session_state.generando)

# Si el usuario pide cancelar, activamos el flag (el hilo de fondo lo verá)
if cancelar_clicked:
    st.session_state.stop_requested = True
    agregar_log("⏹️ Cancelación solicitada por el usuario...")
    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN QUE CORRE EN HILO DE FONDO (no bloquea la UI de Streamlit)
# ══════════════════════════════════════════════════════════════════════════════
def _trabajo_en_segundo_plano(manager, instrucciones, modelo_actual, rutas_temp,
                               generar_excel, generar_ppt, logs_thread, stop_flag):
    """
    Ejecuta ejecutar_flujo_completo en un hilo separado.
    Como Streamlit no permite tocar st.session_state desde otro hilo de forma segura,
    los logs y el resultado se guardan en estructuras planas (listas/dicts) que el
    hilo principal de Streamlit lee periódicamente con st.rerun().
    """
    def log_en_vivo(mensaje):
        logs_thread.append(mensaje)

    def stop_check():
        if stop_flag["valor"]:
            raise Exception("CANCELADO_POR_USUARIO")

    try:
        resultado = ejecutar_flujo_completo(
            manager=manager,
            instrucciones=instrucciones,
            modelo_ia=modelo_actual,
            rutas_archivos=rutas_temp,
            log_callback=log_en_vivo,
            stop_check=stop_check,
            generar_excel=generar_excel,
            generar_ppt=generar_ppt,
        )
        logs_thread.append("__RESULTADO_OK__")
        stop_flag["resultado"] = resultado
    except Exception as e:
        if "CANCELADO_POR_USUARIO" in str(e):
            logs_thread.append("⏹️ Proceso cancelado por el usuario.")
            stop_flag["resultado"] = {"exito": False, "error": "CANCELADO_POR_USUARIO"}
        else:
            logs_thread.append(f"❌ Error inesperado: {e}")
            stop_flag["resultado"] = {"exito": False, "error": str(e)}
    finally:
        for r in rutas_temp:
            try: os.remove(r)
            except: pass
        stop_flag["terminado"] = True

# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA DE GENERACIÓN — lanza el hilo y NO bloquea la interfaz
# ══════════════════════════════════════════════════════════════════════════════
if generar_clicked:
    if not st.session_state.api_key:
        st.error("❌ Debes ingresar una API Key de Gemini en la barra lateral.")
    elif not archivos_cargados:
        st.warning("Por favor, sube al menos un archivo.")
    else:
        # Resetear estado
        st.session_state.logs = []
        st.session_state.evaluacion = None
        st.session_state.archivo_listo = None
        st.session_state.stop_requested = False
        st.session_state.generando = True
        st.session_state.hilo_terminado = False

        agregar_log("🚀 Iniciando generación de tarea...")
        agregar_log(f"📄 Archivos subidos: {len(archivos_cargados)}")
        for arch in archivos_cargados:
            agregar_log(f"   - {arch.name} ({arch.size} bytes)")
        agregar_log(f"📝 Instrucciones: {instrucciones[:100]}...")
# ── Guardar archivos temporalmente y normalizarlos a PDF ──
        rutas_temp = []
        for archivo in archivos_cargados:
            # Guardar el archivo original
            temp_path = f"temp_{archivo.name}"
            with open(temp_path, "wb") as f:
                f.write(archivo.getbuffer())

            ext_actual = archivo.name.lower()

            # CASO 1: Documentos Word (.docx o .doc)
            if ext_actual.endswith(('.docx', '.doc')):
                if 'convertir_docx_a_pdf' in globals():
                    carpeta_temp = os.path.dirname(temp_path)
                    pdf_path = convertir_docx_a_pdf(temp_path, carpeta_temp)
                    if pdf_path and os.path.exists(pdf_path):
                        rutas_temp.append(pdf_path)
                        agregar_log(f"📄 Word convertido a PDF: {os.path.basename(pdf_path)}")
                        try: os.remove(temp_path)
                        except: pass
                    else:
                        rutas_temp.append(temp_path)
                        agregar_log(f"⚠️ No se pudo convertir {archivo.name} a PDF, se usará el archivo original.")
                else:
                    rutas_temp.append(temp_path)
                    agregar_log(f"⚠️ Función de conversión no disponible, se usará {archivo.name} directamente.")

            # CASO 2: Imágenes (.png, .jpg, .jpeg, .webp, .gif)
            elif ext_actual.endswith(('.png', '.jpg', '.jpeg', '.webp', '.gif')):
                if 'convertir_imagen_a_pdf' in globals():
                    carpeta_temp = os.path.dirname(temp_path)
                    pdf_path = convertir_imagen_a_pdf(temp_path, carpeta_temp)
                    if pdf_path and os.path.exists(pdf_path):
                        rutas_temp.append(pdf_path)
                        agregar_log(f"🖼️ Imagen convertida a PDF: {os.path.basename(pdf_path)}")
                        try: os.remove(temp_path)
                        except: pass
                    else:
                        rutas_temp.append(temp_path)
                        agregar_log(f"⚠️ No se pudo convertir la imagen {archivo.name} a PDF, se usará el original.")
                else:
                    rutas_temp.append(temp_path)

            # CASO 3: Archivos PDF nativos u otros
            else:
                rutas_temp.append(temp_path)
                agregar_log(f"💾 Archivo PDF cargado directamente: {temp_path}")

        # ── Crear ProviderManager y registrar clientes ──
        manager = ProviderManager(cb=None)  # el cb se maneja vía logs_thread, no aquí

        cliente_gemini = crear_cliente_gemini(
            api_key=st.session_state.api_key,
            modelo=st.session_state.modelo_actual,
            cb=None,
            temperature=st.session_state.gemini_temperature,          # <-- ENVIAR TEMP
            thinking_budget=st.session_state.gemini_thinking_budget   # <-- ENVIAR BUDGET
        )
        manager.registrar_cliente(cliente_gemini)

        if st.session_state.api_key_backup:
            cliente_gemini_backup = crear_cliente_gemini(
                api_key=st.session_state.api_key_backup,
                modelo=st.session_state.modelo_actual,
                cb=None,
                temperature=st.session_state.gemini_temperature,        # <-- ENVIAR TEMP
                thinking_budget=st.session_state.gemini_thinking_budget # <-- ENVIAR BUDGET
            )
            cliente_gemini_backup.nombre = "Gemini_Respaldo"
            cliente_gemini_backup.es_respaldo = True
            manager.registrar_cliente(cliente_gemini_backup)
            agregar_log("✅ Gemini de respaldo registrado.")

        ##### Groq y Cohere: solo se registran si su panel está activo (abierto)
        # Y MUESTRA su API Key. Si el panel está colapsado, el proveedor se
        # ignora aunque la clave siga guardada en session_state.
        if st.session_state.get("groq_expander_open", False) and st.session_state.groq_api_key:
            cliente_groq = crear_cliente_groq(
                api_key=st.session_state.groq_api_key,
                modelo=st.session_state.groq_model,
                cb=None
            )
            manager.registrar_cliente(cliente_groq)
            agregar_log("✅ Groq activo y registrado.")
        elif st.session_state.get("groq_expander_open", False):
            agregar_log("⚠️ Panel de Groq abierto pero sin API Key: se omite.")

        if st.session_state.get("cohere_expander_open", False) and st.session_state.cohere_api_key:
            cliente_cohere = crear_cliente_cohere(
                api_key=st.session_state.cohere_api_key,
                modelo=st.session_state.cohere_model,
                cb=None
            )
            manager.registrar_cliente(cliente_cohere)
            agregar_log("✅ Cohere activo y registrado.")
        elif st.session_state.get("cohere_expander_open", False):
            agregar_log("⚠️ Panel de Cohere abierto pero sin API Key: se omite.")

        # ── Estructuras compartidas entre el hilo y la UI principal ──
        logs_thread = []
        stop_flag = {"valor": False, "terminado": False, "resultado": None}

        st.session_state._logs_thread = logs_thread
        st.session_state._stop_flag = stop_flag

        hilo = threading.Thread(
            target=_trabajo_en_segundo_plano,
            args=(manager, instrucciones, st.session_state.modelo_actual, rutas_temp,
                  st.session_state.generar_excel, st.session_state.generar_ppt,
                  logs_thread, stop_flag),
            daemon=True
        )
        hilo.start()
        st.session_state.hilo_activo = True
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# MONITOR DEL HILO — se ejecuta en cada rerun mientras "generando" es True.
# Esto permite que el botón Cancelar SIEMPRE responda, porque el script principal
# de Streamlit nunca queda bloqueado: solo revisa el estado del hilo y se refresca.
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.generando and st.session_state.get("hilo_activo"):
    stop_flag = st.session_state.get("_stop_flag")
    logs_thread = st.session_state.get("_logs_thread")

    # Propagar cancelación solicitada al hilo
    if st.session_state.stop_requested and stop_flag:
        stop_flag["valor"] = True

    # Volcar logs nuevos del hilo al historial visible
    if logs_thread:
        nuevos = [m for m in logs_thread if m not in st.session_state.logs]
        for m in nuevos:
            if m != "__RESULTADO_OK__":
                agregar_log(m)

    status = st.status("⚙️ Procesando en segundo plano...", expanded=True)
    if logs_thread:
        for m in logs_thread[-15:]:
            if m != "__RESULTADO_OK__":
                status.write(m)

    if stop_flag and stop_flag.get("terminado"):
        resultado = stop_flag.get("resultado") or {"exito": False, "error": "Sin resultado."}
        st.session_state.generando = False
        st.session_state.hilo_activo = False

        if resultado.get("error") == "CANCELADO_POR_USUARIO":
            status.update(label="⏹️ Generación cancelada", state="error")
            st.warning("Generación cancelada. No se ha guardado ningún archivo.")
        elif resultado.get("exito"):
            st.session_state.archivo_listo = resultado["ruta"]
            st.session_state.evaluacion = resultado.get("evaluacion")
            st.session_state.ruta_excel = resultado.get("ruta_excel")
            st.session_state.ruta_ppt = resultado.get("ruta_ppt")
            status.update(label="✅ Tarea generada y evaluada con éxito", state="complete")
            st.success("✅ ¡Tarea generada con éxito! Revisa la evaluación y descarga tus archivos.")
        else:
            status.update(label="❌ Error en la generación", state="error")
            st.error(f"❌ Error: {resultado.get('error', 'Desconocido')}")

        st.rerun()
    else:
        # Aún procesando: esperar un poco y refrescar para revisar de nuevo.
        # Este sleep es corto para que el botón Cancelar quede disponible casi
        # de inmediato en el siguiente rerun.
        time.sleep(1.2)
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# CUADRO DE EVALUACIÓN IA, BOTONES DE DESCARGA, REGISTRO HISTÓRICO
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.evaluacion:
    ev = st.session_state.evaluacion
    nota_str = ev.get("nota", "?")
    justificacion = ev.get("justificacion", "Sin justificación.")
    try:
        nota_num = float(nota_str)
        if nota_num >= 14:
            color_borde = "#28a745"; emoji_nota = "✅"
        elif nota_num >= 11:
            color_borde = "#ffc107"; emoji_nota = "⚠️"
        else:
            color_borde = "#dc3545"; emoji_nota = "❌"
    except ValueError:
        color_borde = "#6c757d"; emoji_nota = "❓"

    st.markdown("---")
    st.markdown("### 🎓 Evaluación del Docente IA")
    st.markdown(
        f"""
        <div style="border:2px solid {color_borde};border-radius:12px;padding:20px 24px;background-color:#1e1e1e;margin-bottom:16px;">
            <div style="font-size:2.2em;font-weight:bold;color:{color_borde};text-align:center;margin-bottom:12px;">
                {emoji_nota} NOTA: {nota_str} / 20
            </div>
            <hr style="border-color:{color_borde};opacity:0.3;margin:12px 0;">
            <div style="font-size:0.95em;color:#e0e0e0;line-height:1.7;white-space:pre-wrap;">{justificacion}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ── BOTONES DE DESCARGA (SOLO SI EXISTEN) ──
if st.session_state.get('archivo_listo') and os.path.exists(st.session_state.archivo_listo):
    # Word
    with open(st.session_state.archivo_listo, "rb") as file:
        st.download_button(
            label="📄 Descargar Word",
            data=file,
            file_name="Tarea_Realizada.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    # Excel (solo si fue generado)
    ruta_excel = st.session_state.archivo_listo.replace('.docx', '_datos.xlsx')
    if os.path.exists(ruta_excel):
        with open(ruta_excel, "rb") as file:
            st.download_button(
                label="📊 Descargar Excel con datos",
                data=file,
                file_name="Tarea_Realizada_datos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # PowerPoint (solo si fue generado)
    ruta_ppt = st.session_state.archivo_listo.replace('.docx', '.pptx')
    if os.path.exists(ruta_ppt):
        with open(ruta_ppt, "rb") as file:
            st.download_button(
                label="🔶 Descargar PowerPoint",
                data=file,
                file_name="Tarea_Realizada.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

# ── REGISTRO HISTÓRICO ──
st.markdown("---")
if st.session_state.debug_mode:
    st.subheader("📡 Registro histórico de la última ejecución")
    st.caption("Los logs se actualizan automáticamente al finalizar la tarea.")
    log_container = st.container(height=400)
    with log_container:
        for msg in st.session_state.logs:
            st.text(msg)
else:
    st.info("💡 Activa el 'Modo Depuración' en la barra lateral para ver los logs históricos.")
