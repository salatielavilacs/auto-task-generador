// auto-task-frontend/app/page.tsx
'use client';

import React, { useState, useEffect, useRef } from 'react';

interface Evaluacion {
  nota: string;
  justificacion: string;
}

interface Descargas {
  word: string | null;
  excel: string | null;
  ppt: string | null;
}

export default function Home() {
  // En desarrollo usará localhost:7860 de forma automática.
  // En producción (Vercel) leerá la variable de entorno que configuremos.
  const [backendUrl, setBackendUrl] = useState(
    process.env.NEXT_PUBLIC_BACKEND_URL || 'http://localhost:7860'
  );

  // Ya no se requiere el useEffect anterior que detectaba el hostname.


  // Estados de Configuración — Gemini
  const [apiKey, setApiKey] = useState('');
  const [apiKeyBackup, setApiKeyBackup] = useState('');
  const [modelo, setModelo] = useState('gemini-3.1-flash-lite');
  const [temperature, setTemperature] = useState(0.7);
  const [thinkingBudget, setThinkingBudget] = useState(0);
  const [generarExcel, setGenerarExcel] = useState(true);
  const [generarPpt, setGenerarPpt] = useState(true);
  const [showApiKey, setShowApiKey] = useState(false);
  const [showApiKeyBackup, setShowApiKeyBackup] = useState(false);

  // Estados de Configuración — Groq (panel colapsable = interruptor)
  const [groqActivo, setGroqActivo] = useState(false);
  const [groqApiKey, setGroqApiKey] = useState('');
  const [groqModelo, setGroqModelo] = useState('llama-3.1-70b-versatile');

  // Estados de Configuración — Cohere (mismo patrón que Groq)
  const [cohereActivo, setCohereActivo] = useState(false);
  const [cohereApiKey, setCohereApiKey] = useState('');
  const [cohereModelo, setCohereModelo] = useState('command-r');

  // Modo depuración (equivalente al checkbox de Streamlit)
  const [debugMode, setDebugMode] = useState(true);

  // Estados de la Aplicación
  const [archivos, setArchivos] = useState<File[]>([]);
  const [instrucciones, setInstrucciones] = useState('');
  const [generando, setGenerando] = useState(false);
  const [logs, setLogs] = useState<string[]>([]);
  const [evaluacion, setEvaluacion] = useState<Evaluacion | null>(null);
  const [descargas, setDescargas] = useState<Descargas | null>(null);
  
  // Estado para el diseño visual de la zona de arrastre
  const [isDragging, setIsDragging] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const abortControllerRef = useRef<AbortController | null>(null);

  // --- EFECTO: CAPTURAR CONTROL + V (PEGAR IMÁGENES DESDE EL PORTAPAPELES) ---
  useEffect(() => {
    const handlePaste = (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;

      const filesList: File[] = [];
      for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
          const file = items[i].getAsFile();
          if (file) {
            const nombreUnico = `captura_${Date.now()}.png`;
            const archivoRenombrado = new File([file], nombreUnico, { type: file.type });
            filesList.push(archivoRenombrado);
          }
        }
      }

      if (filesList.length > 0) {
        setArchivos((prev) => [...prev, ...filesList]);
        agregarLog(`📋 Imagen pegada desde el portapapeles.`);
      }
    };

    window.addEventListener('paste', handlePaste);
    return () => window.removeEventListener('paste', handlePaste);
  }, []);

  const agregarLog = (msg: string) => {
    setLogs((prev) => [...prev, `[${new Date().toLocaleTimeString()}] ${msg}`]);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const nuevos = Array.from(e.target.files);
      setArchivos((prev) => [...prev, ...nuevos]);
      nuevos.forEach(f => agregarLog(`💾 Archivo cargado: ${f.name}`));
    }
  };

  const removerArchivo = (idx: number) => {
    setArchivos((prev) => prev.filter((_, i) => i !== idx));
  };

  const limpiarLogs = () => {
    setLogs([]);
  };

  // --- CONTROLADORES DE EVENTOS DE ARRASTRE ---
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const nuevos = Array.from(e.dataTransfer.files);
      setArchivos((prev) => [...prev, ...nuevos]);
      nuevos.forEach(f => agregarLog(`💾 Archivo soltado: ${f.name}`));
    }
  };

  // --- FUNCIÓN PRINCIPAL: LLAMAR AL BACKEND CON SOPORTE DE STREAMING EN VIVO (MÉTODO ROBUSTO) ---
  const handleGenerar = async () => {
    if (!apiKey) {
      alert("❌ Debes ingresar tu API Key principal de Gemini.");
      return;
    }
    if (archivos.length === 0) {
      alert("⚠️ Por favor, sube o pega al menos un archivo.");
      return;
    }

    setGenerando(true);
    setEvaluacion(null);
    setDescargas(null);
    setLogs([]);
    
    agregarLog("🚀 Iniciando generación de tarea en el servidor...");
    agregarLog(`📄 Archivos subidos: ${archivos.length}`);
    archivos.forEach((f) => agregarLog(`   - ${f.name} (${(f.size / 1024).toFixed(0)} KB)`));

    if (groqActivo && groqApiKey) agregarLog("✅ Groq activo, se enviará al backend.");
    if (cohereActivo && cohereApiKey) agregarLog("✅ Cohere activo, se enviará al backend.");

    const formData = new FormData();
    formData.append("api_key", apiKey);
    formData.append("api_key_backup", apiKeyBackup);
    formData.append("modelo", modelo);
    formData.append("temperature", temperature.toString());
    formData.append("thinking_budget", thinkingBudget.toString());
    formData.append("instrucciones", instrucciones);

    formData.append("generar_excel", generarExcel.toString());
    formData.append("generar_ppt", generarPpt.toString());

    formData.append("groq_activo", (groqActivo && !!groqApiKey).toString());
    if (groqActivo && groqApiKey) {
      formData.append("groq_api_key", groqApiKey);
      formData.append("groq_modelo", groqModelo);
    }
    formData.append("cohere_activo", (cohereActivo && !!cohereApiKey).toString());
    if (cohereActivo && cohereApiKey) {
      formData.append("cohere_api_key", cohereApiKey);
      formData.append("cohere_modelo", cohereModelo);
    }

    archivos.forEach((file) => {
      formData.append("files", file);
    });

    const controller = new AbortController();
    abortControllerRef.current = controller;

    try {
      const response = await fetch(`${backendUrl}/generar`, {
        method: "POST",
        body: formData,
        signal: controller.signal,
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(errorText || "Error en el servidor");
      }

      const reader = response.body?.getReader();
      const decoder = new TextDecoder("utf-8");
      
      if (!reader) {
        throw new Error("No se pudo inicializar el lector de flujo.");
      }

      let buffer = "";

      while (true) {
        const { value, done } = await reader.read();
        if (done) break;

        // Decodificar el fragmento binario actual y sumarlo al acumulador
        buffer += decoder.decode(value, { stream: true });
        
        // Dividir por saltos de línea individuales (\n es más tolerante a variaciones de red)
        const lines = buffer.split("\n");
        buffer = lines.pop() || ""; // Mantener fragmentos incompletos en el buffer

        for (let line of lines) {
          line = line.trim();
          if (!line) continue;

          if (line.startsWith("data:")) {
            // Eliminar de forma segura el prefijo "data:" y espacios adicionales
            const rawData = line.replace(/^data:\s*/, "").trim();
            if (!rawData) continue;

            try {
              const parsed = JSON.parse(rawData);
              
              if (parsed.type === "log") {
                agregarLog(parsed.message);
              } else if (parsed.type === "error") {
                agregarLog(`❌ Error: ${parsed.message}`);
              } else if (parsed.type === "result") {
                const resultData = parsed.data;
                if (resultData.exito) {
                  agregarLog("✅ ¡Tarea generada y evaluada con éxito!");
                  setEvaluacion(resultData.evaluacion);
                  setDescargas(resultData.archivos);
                } else {
                  agregarLog(`❌ Error en el resultado: ${resultData.error}`);
                }
              }
            } catch (e) {
              // Si falla una línea por estar incompleta, el try-catch interno evita que se rompa el bucle principal
              console.warn("Fragmento de flujo ignorado por parseo incompleto:", line);
            }
          }
        }
      }

    } catch (error: any) {
      if (error.name === "AbortError") {
        agregarLog("⏹️ Generación cancelada por el usuario.");
      } else {
        agregarLog(`❌ Falló la conexión con el servidor: ${error.message}`);
      }
    } finally {
      setGenerando(false);
      abortControllerRef.current = null;
    }
  };

  const handleCancelar = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
      agregarLog("⏹️ Cancelación solicitada por el usuario...");
    }
  };

  return (
    <div className="flex min-h-screen bg-[#0d1117] text-gray-100 font-sans">

      {/* BARRA LATERAL (CONFIGURACIÓN) */}
      <aside className="w-80 bg-[#161b22] border-r border-[#30363d] p-6 flex flex-col gap-6 select-none overflow-y-auto">
        <div>
          <h2 className="text-xl font-bold flex items-center gap-2 text-blue-400">
            ⚙️ Configuración
          </h2>
          <p className="text-xs text-gray-400 mt-1">Ajusta los parámetros del motor</p>
        </div>

       {/* SECCIÓN GEMINI */}
        <div className="flex flex-col gap-4">
          <h3 className="text-sm font-semibold tracking-wider text-gray-400 uppercase">🔵 Gemini</h3>
          
          {/* API KEY PRINCIPAL */}
          <div>
            <label className="text-xs text-gray-400 block mb-1">API Key (Principal)</label>
            <div className="relative">
              <input
                type={showApiKey ? "text" : "password"}
                value={apiKey}
                onChange={e => setApiKey(e.target.value)}
                className="w-full bg-[#0d1117] border border-[#30363d] rounded p-2 pr-10 text-sm focus:outline-none focus:border-blue-500 text-gray-200"
                placeholder="AIzaSy..."
              />
              <button
                type="button"
                onClick={() => setShowApiKey(!showApiKey)}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-200 transition-colors focus:outline-none text-base select-none"
                title={showApiKey ? "Ocultar clave" : "Mostrar clave"}
              >
                {showApiKey ? "👁️" : "🙈"}
              </button>
            </div>
          </div>

          {/* API KEY RESPALDO */}
          <div>
            <label className="text-xs text-gray-400 block mb-1">API Key (Respaldo)</label>
            <div className="relative">
              <input
                type={showApiKeyBackup ? "text" : "password"}
                value={apiKeyBackup}
                onChange={e => setApiKeyBackup(e.target.value)}
                className="w-full bg-[#0d1117] border border-[#30363d] rounded p-2 pr-10 text-sm focus:outline-none focus:border-blue-500 text-gray-200"
                placeholder="Opcional..."
              />
              <button
                type="button"
                onClick={() => setShowApiKeyBackup(!showApiKeyBackup)}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-200 transition-colors focus:outline-none text-base select-none"
                title={showApiKeyBackup ? "Ocultar clave" : "Mostrar clave"}
              >
                {showApiKeyBackup ? "👁️" : "🙈"}
              </button>
            </div>
          </div>

          <div>
            <label className="text-xs text-gray-400 block mb-1">Modelo de IA</label>
            <select
              value={modelo}
              onChange={e => setModelo(e.target.value)}
              className="w-full bg-[#0d1117] border border-[#30363d] rounded p-2 text-sm focus:outline-none"
            >
              <option value="gemini-3.1-flash-lite">gemini-3.1-flash-lite</option>
              <option value="gemini-3.5-flash">gemini-3.5-flash</option>
              <option value="gemini-2.5-flash">gemini-2.5-flash</option>
            </select>
          </div>
        </div>

        {/* PARÁMETROS AVANZADOS */}
        <div className="flex flex-col gap-4 border-t border-[#30363d] pt-4">
          <h3 className="text-sm font-semibold tracking-wider text-gray-400 uppercase">🎛️ Avanzado</h3>
          <div>
            <div className="flex justify-between text-xs text-gray-400 mb-1">
              <span>Temperatura</span>
              <span className="text-blue-400 font-bold">{temperature}</span>
            </div>
            <input
              type="range"
              min="0" max="2" step="0.1"
              value={temperature}
              onChange={e => setTemperature(parseFloat(e.target.value))}
              className="w-full accent-blue-500 cursor-pointer"
            />
          </div>
          <div>
            <label className="text-xs text-gray-400 block mb-1">Limitador de Razonamiento (Thinking)</label>
            <select
              value={thinkingBudget}
              onChange={e => setThinkingBudget(parseInt(e.target.value))}
              className="w-full bg-[#0d1117] border border-[#30363d] rounded p-2 text-sm focus:outline-none"
            >
              <option value={0}>Desactivado</option>
              <option value={1024}>Bajo (1024 tokens)</option>
              <option value={2048}>Medio (2048 tokens)</option>
              <option value={4096}>Alto (4096 tokens)</option>
            </select>
          </div>
        </div>

        {/* ENTREGABLES */}
        <div className="flex flex-col gap-3 border-t border-[#30363d] pt-4">
          <h3 className="text-sm font-semibold tracking-wider text-gray-400 uppercase">📁 Entregables</h3>
          <label className="flex items-center gap-2 text-sm text-gray-400 cursor-not-allowed">
            <input type="checkbox" checked disabled className="rounded accent-blue-500" />
            📄 Word (Obligatorio)
          </label>
          <label className="flex items-center gap-2 text-sm cursor-pointer">
            <input type="checkbox" checked={generarExcel} onChange={e => setGenerarExcel(e.target.checked)} className="rounded accent-blue-500" />
            📊 Excel de Datos
          </label>
          <label className="flex items-center gap-2 text-sm cursor-pointer">
            <input type="checkbox" checked={generarPpt} onChange={e => setGenerarPpt(e.target.checked)} className="rounded accent-blue-500" />
            🔶 PowerPoint
          </label>
        </div>

        {/* GROQ — panel clic-para-activar/desactivar */}
        <div className="border-t border-[#30363d] pt-4">
          <button
            type="button"
            onClick={() => setGroqActivo(v => !v)}
            className="w-full flex items-center justify-between text-sm font-semibold tracking-wider text-gray-300 uppercase hover:text-orange-400 transition-colors"
          >
            <span className="flex items-center gap-2 normal-case text-sm">
              🟠 Groq
              {groqActivo && groqApiKey && (
                <span className="text-[10px] normal-case bg-green-600/20 text-green-400 px-2 py-0.5 rounded-full">Activo</span>
              )}
            </span>
            <span className={`transition-transform text-gray-500 ${groqActivo ? 'rotate-180' : ''}`}>▾</span>
          </button>

          {groqActivo ? (
            <div className="flex flex-col gap-3 mt-3">
              <div>
                <label className="text-xs text-gray-400 block mb-1">API Key (Groq)</label>
                <input
                  type="password"
                  value={groqApiKey}
                  onChange={e => setGroqApiKey(e.target.value)}
                  className="w-full bg-[#0d1117] border border-[#30363d] rounded p-2 text-sm focus:outline-none focus:border-orange-500"
                  placeholder="gsk_..."
                />
              </div>
              <div>
                <label className="text-xs text-gray-400 block mb-1">Modelo Groq</label>
                <select
                  value={groqModelo}
                  onChange={e => setGroqModelo(e.target.value)}
                  className="w-full bg-[#0d1117] border border-[#30363d] rounded p-2 text-sm focus:outline-none"
                >
                  <option value="llama-3.1-70b-versatile">llama-3.1-70b-versatile</option>
                  <option value="mixtral-8x7b-32768">mixtral-8x7b-32768</option>
                  <option value="gemma2-9b-it">gemma2-9b-it</option>
                </select>
              </div>
              {!groqApiKey && (
                <p className="text-[11px] text-yellow-500">⚠️ Panel abierto, pero falta la API Key.</p>
              )}
            </div>
          ) : (
            <p className="text-[11px] text-gray-500 mt-1">⏸️ Desactivado — clic en el título para activar</p>
          )}
        </div>

        {/* COHERE — panel clic-para-activar/desactivar */}
        <div className="border-t border-[#30363d] pt-4">
          <button
            type="button"
            onClick={() => setCohereActivo(v => !v)}
            className="w-full flex items-center justify-between text-sm font-semibold tracking-wider text-gray-300 uppercase hover:text-purple-400 transition-colors"
          >
            <span className="flex items-center gap-2 normal-case text-sm">
              🟣 Cohere
              {cohereActivo && cohereApiKey && (
                <span className="text-[10px] normal-case bg-green-600/20 text-green-400 px-2 py-0.5 rounded-full">Activo</span>
              )}
            </span>
            <span className={`transition-transform text-gray-500 ${cohereActivo ? 'rotate-180' : ''}`}>▾</span>
          </button>

          {cohereActivo ? (
            <div className="flex flex-col gap-3 mt-3">
              <div>
                <label className="text-xs text-gray-400 block mb-1">API Key (Cohere)</label>
                <input
                  type="password"
                  value={cohereApiKey}
                  onChange={e => setCohereApiKey(e.target.value)}
                  className="w-full bg-[#0d1117] border border-[#30363d] rounded p-2 text-sm focus:outline-none focus:border-purple-500"
                  placeholder="co-..."
                />
              </div>
              <div>
                <label className="text-xs text-gray-400 block mb-1">Modelo Cohere</label>
                <select
                  value={cohereModelo}
                  onChange={e => setCohereModelo(e.target.value)}
                  className="w-full bg-[#0d1117] border border-[#30363d] rounded p-2 text-sm focus:outline-none"
                >
                  <option value="command-r">command-r</option>
                  <option value="command-r-plus">command-r-plus</option>
                  <option value="command-light">command-light</option>
                </select>
              </div>
              {!cohereApiKey && (
                <p className="text-[11px] text-yellow-500">⚠️ Panel abierto, pero falta la API Key.</p>
              )}
            </div>
          ) : (
            <p className="text-[11px] text-gray-500 mt-1">⏸️ Desactivado — clic en el título para activar</p>
          )}
        </div>

        {/* MODO DEPURACIÓN + LIMPIAR LOGS */}
        <div className="flex flex-col gap-3 border-t border-[#30363d] pt-4">
          <label className="flex items-center gap-2 text-sm cursor-pointer">
            <input
              type="checkbox"
              checked={debugMode}
              onChange={e => setDebugMode(e.target.checked)}
              className="rounded accent-blue-500"
            />
            🔍 Modo Depuración (ver logs)
          </label>
          <button
            type="button"
            onClick={limpiarLogs}
            className="text-xs text-gray-400 hover:text-red-400 border border-[#30363d] hover:border-red-500/50 rounded-lg py-2 transition-colors"
          >
            🗑️ Limpiar logs
          </button>
        </div>
      </aside>

      {/* PANEL PRINCIPAL */}
      <main className="flex-1 p-8 flex flex-col gap-6 overflow-y-auto">
        <header className="flex justify-between items-center border-b border-[#30363d] pb-4">
          <div>
            <h1 className="text-3xl font-extrabold tracking-tight text-white flex items-center gap-2">
              🤖 AUTO-TASK GENERADOR
            </h1>
            <p className="text-sm text-gray-400 mt-1">Crea entregables académicos completos y profesionales</p>
          </div>
        </header>

        {/* ZONA DE ARRASTRE / PEGAR */}
        <section
          onClick={() => fileInputRef.current?.click()}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          className={`border-2 border-dashed transition-colors rounded-xl p-8 flex flex-col items-center justify-center gap-3 cursor-pointer ${
            isDragging 
              ? 'border-blue-500 bg-blue-500/10' 
              : 'border-[#30363d] hover:border-blue-500 bg-[#161b22]'
          }`}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            multiple
            className="hidden"
            // Se ha removido el atributo 'accept' para permitir la selección de cualquier archivo.
          />
          <div className="text-4xl">📥</div>
          <p className="text-center font-semibold text-gray-200">
            Arrastra aquí tus archivos o haz clic para buscarlos
          </p>
          <p className="text-xs text-gray-400 text-center">
            Soporta cualquier tipo de archivo (PDF, Word, Imágenes, Excel) o <strong className="text-blue-400">pega capturas directamente con Ctrl+V</strong>
          </p>
        </section>

        {/* LISTADO DE ARCHIVOS CARGADOS */}
        {archivos.length > 0 && (
          <section className="bg-[#161b22] border border-[#30363d] rounded-xl p-4 flex flex-col gap-2">
            <h3 className="text-xs font-semibold tracking-wider text-gray-400 uppercase">Archivos Listos ({archivos.length})</h3>
            <div className="flex flex-wrap gap-2">
              {archivos.map((file, idx) => (
                <div key={idx} className="bg-[#0d1117] border border-[#30363d] rounded-full px-3 py-1 text-xs flex items-center gap-2 text-gray-300">
                  <span>📄 {file.name} ({(file.size / 1024).toFixed(0)} KB)</span>
                  <button onClick={(e) => { e.stopPropagation(); removerArchivo(idx); }} className="hover:text-red-400 font-bold">×</button>
                </div>
              ))}
            </div>
          </section>
        )}

        {/* INSTRUCCIONES */}
        <section className="flex flex-col gap-2">
          <label className="text-xs font-semibold tracking-wider text-gray-400 uppercase">Instrucciones para la IA:</label>
          <textarea
            rows={3}
            value={instrucciones}
            onChange={e => setInstrucciones(e.target.value)}
            placeholder="Escribe indicaciones específicas, aclaraciones de rúbrica o el número de ejercicio a resolver..."
            className="w-full bg-[#161b22] border border-[#30363d] rounded-xl p-4 text-sm focus:outline-none focus:border-blue-500 text-gray-200"
          />
        </section>

        {/* BOTONES DE ACCIÓN */}
        <section className="flex gap-4">
          <button
            onClick={handleGenerar}
            disabled={generando}
            className="flex-1 bg-blue-600 hover:bg-blue-500 disabled:bg-blue-800 disabled:cursor-not-allowed transition-colors py-3 rounded-xl font-bold text-white shadow-lg text-center"
          >
            {generando ? "⚙️ Generando en segundo plano..." : "🚀 Generar Tarea"}
          </button>
          <button
            onClick={handleCancelar}
            disabled={!generando}
            className="flex-1 bg-red-600 hover:bg-red-500 disabled:bg-[#21262d] disabled:text-gray-600 disabled:cursor-not-allowed transition-colors py-3 rounded-xl font-bold text-white shadow-lg text-center"
          >
            ⏹️ Cancelar
          </button>
        </section>

        {/* LOGS EN TIEMPO REAL */}
        {debugMode && logs.length > 0 && (
          <section className="bg-[#161b22] border border-[#30363d] rounded-xl p-4 flex flex-col gap-2">
            <h3 className="text-xs font-semibold tracking-wider text-gray-400 uppercase">Registro de Ejecución</h3>
            <div className="bg-[#0d1117] border border-[#30363d] rounded-lg p-3 font-mono text-xs text-green-400 h-32 overflow-y-auto flex flex-col gap-1">
              {logs.map((log, i) => (
                <div key={i}>{log}</div>
              ))}
            </div>
          </section>
        )}
        {!debugMode && (
          <p className="text-xs text-gray-500 -mt-2">💡 Activa el "Modo Depuración" en la barra lateral para ver los logs.</p>
        )}

        {/* TARJETA DE EVALUACIÓN DE LA IA */}
        {evaluacion && (
          <section className="bg-[#161b22] border-2 border-green-500 rounded-xl p-6 flex flex-col gap-4">
            <h3 className="text-xl font-bold text-green-400 flex items-center gap-2">
              🎓 Evaluación del Docente IA
            </h3>
            <div className="text-3xl font-extrabold text-green-500 text-center">
              NOTA: {evaluacion.nota} / 20
            </div>
            <hr className="border-[#30363d]" />
            <div className="text-sm leading-relaxed text-gray-300 whitespace-pre-wrap">
              {evaluacion.justificacion}
            </div>
          </section>
        )}

        {/* BOTONES DE DESCARGA */}
        {descargas && (
          <section className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {descargas.word && (
              <a
                href={`${backendUrl}/descargar/${descargas.word}`}
                download
                className="bg-blue-600 hover:bg-blue-500 transition-colors py-3 rounded-xl font-bold text-white text-center shadow-lg block"
              >
                📄 Descargar Word
              </a>
            )}
            {descargas.excel && (
              <a
                href={`${backendUrl}/descargar/${descargas.excel}`}
                download
                className="bg-green-600 hover:bg-green-500 transition-colors py-3 rounded-xl font-bold text-white text-center shadow-lg block"
              >
                📊 Descargar Excel
              </a>
            )}
            {descargas.ppt && (
              <a
                href={`${backendUrl}/descargar/${descargas.ppt}`}
                download
                className="bg-orange-600 hover:bg-orange-500 transition-colors py-3 rounded-xl font-bold text-white text-center shadow-lg block"
              >
                🔶 Descargar PowerPoint
              </a>
            )}
          </section>
        )}
      </main>
    </div>
  );
}