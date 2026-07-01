# main.py
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
import shutil
import os
import json
import queue
import asyncio
from typing import List
from concurrent.futures import ThreadPoolExecutor

# Importar funciones existentes
from analizador import ProviderManager, crear_cliente_gemini, ejecutar_flujo_completo

app = FastAPI()

# Permitir que Next.js se conecte al Backend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/generar")
async def generar_tarea(
    api_key: str = Form(...),
    api_key_backup: str = Form(""),
    modelo: str = Form("gemini-3.1-flash-lite"),
    temperature: float = Form(0.7),
    thinking_budget: int = Form(0),
    instrucciones: str = Form(""),
    generar_excel: str = Form("true"),  # Capturar preferencia de la UI
    generar_ppt: str = Form("true"),    # Capturar preferencia de la UI
    files: List[UploadFile] = File(...)
):
    rutas_temp = []
    # 1. Guardar archivos subidos temporalmente
    for file in files:
        temp_path = f"temp_{file.filename}"
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        rutas_temp.append(temp_path)

    # Cola de comunicación segura entre el hilo de ejecución y el generador asíncrono
    log_queue = queue.Queue()

    def log_callback(mensaje: str):
        log_queue.put({"type": "log", "message": mensaje})

    def run_flow_thread():
        try:
            # 2. Configurar el gestor de proveedores
            manager = ProviderManager(cb=None)
            
            cliente_primary = crear_cliente_gemini(
                api_key=api_key,
                modelo=modelo,
                temperature=temperature,
                thinking_budget=thinking_budget
            )
            manager.registrar_cliente(cliente_primary)
            
            if api_key_backup:
                cliente_backup = crear_cliente_gemini(
                    api_key=api_key_backup,
                    modelo=modelo,
                    temperature=temperature,
                    thinking_budget=thinking_budget
                )
                cliente_backup.nombre = "Gemini_Respaldo"
                cliente_backup.es_respaldo = True
                manager.registrar_cliente(cliente_backup)

            # 3. Ejecutar flujo secuencial
            resultado = ejecutar_flujo_completo(
                manager=manager,
                instrucciones=instrucciones,
                modelo_ia=modelo,
                rutas_archivos=rutas_temp,
                log_callback=log_callback,  # Pasar el callback para capturar logs
                generar_word=True,
                generar_excel=(generar_excel.lower() == "true"),
                generar_ppt=(generar_ppt.lower() == "true")
            )

            if resultado.get("exito"):
                log_queue.put({
                    "type": "result",
                    "data": {
                        "exito": True,
                        "evaluacion": resultado.get("evaluacion"),
                        "archivos": {
                            "word": resultado.get("ruta"),
                            "excel": resultado.get("ruta_excel"),
                            "ppt": resultado.get("ruta_ppt")
                        }
                    }
                })
            else:
                log_queue.put({"type": "error", "message": resultado.get("error", "Error desconocido")})

        except Exception as e:
            log_queue.put({"type": "error", "message": str(e)})
        finally:
            log_queue.put({"type": "done"})

    # Ejecutar el flujo en un hilo secundario para no bloquear el bucle de eventos asíncrono
    loop = asyncio.get_running_loop()
    executor = ThreadPoolExecutor(max_workers=1)
    loop.run_in_executor(executor, run_flow_thread)

    async def event_generator():
        try:
            while True:
                try:
                    # Intento de lectura no bloqueante de la cola
                    item = log_queue.get_nowait()
                except queue.Empty:
                    await asyncio.sleep(0.2)
                    continue

                if item["type"] == "done":
                    break
                
                # Enviar datos formateados como eventos de transmisión
                yield f"data: {json.dumps(item)}\n\n"
                
        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"
        finally:
            # Limpieza de archivos temporales de entrada
            for r in rutas_temp:
                if os.path.exists(r):
                    try:
                        os.remove(r)
                    except:
                        pass

    return StreamingResponse(event_generator(), media_type="text/event-stream")

# Endpoint para descargar los archivos resultantes
@app.get("/descargar/{nombre_archivo}")
def descargar_archivo(nombre_archivo: str):
    if os.path.exists(nombre_archivo):
        return FileResponse(nombre_archivo, filename=nombre_archivo)
    raise HTTPException(status_code=404, detail="Archivo no encontrado")