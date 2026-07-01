# 🤖 AUTO-TASK GENERADOR

**AUTO-TASK GENERADOR** es una plataforma full-stack diseñada para la automatización, estructuración y compilación de entregables académicos e informes utilizando Inteligencia Artificial. El sistema procesa guías o consignas de estudio y genera de forma automatizada documentos estructurados en múltiples formatos profesionales (**Word (.docx)**, **Excel (.xlsx)** y **PowerPoint (.pptx)**), acompañados de una rúbrica de evaluación docente simulada por IA.

---

## 📢 Estado del Proyecto y Fase Experimental

Este proyecto se encuentra actualmente en **fase experimental (Beta)**. Es una prueba de concepto técnica orientada a explorar la automatización de flujos documentales.

* **Tareas Cortas**: El sistema muestra un desempeño óptimo y preciso en tareas de extensión corta (1 a 2 páginas) o resolución de ejercicios numéricos puntuales, logrando de forma consistente calificaciones simuladas por el docente IA **superiores a 17/20**.
* **Tareas Largas / Informes Complejos**: Para entregables de gran extensión, los archivos autogenerados pueden requerir refinamiento, revisión tipográfica o corrección manual de estilo por parte del usuario. Sin embargo, son sumamente útiles como un **punto de partida estructurado (borrador inicial)**, ahorrando horas en la maquetación y planteamiento del contenido.

---

## 🚀 Características Principales

* **Despliegue Full-Stack Desacoplado**: Frontend optimizado en Next.js y un Backend robusto desarrollado en FastAPI y empaquetado en contenedores Docker.
* **Transmisión de Logs en Tiempo Real**: Implementación de Server-Sent Events (SSE) para transmitir el progreso de la generación del servidor al navegador del cliente en vivo.
* **Generación Multi-formato**: Generación dinámica de informes en Word, hojas de cálculo en Excel con fórmulas nativas de validación, y presentaciones dinámicas en PowerPoint.
* **Clasificación Inteligente de Tareas**: Motor de enrutamiento automático que clasifica el flujo de trabajo según la tipología del entregable (Matemáticas/Estadística vs. Redacción General) para aplicar prompts específicos.
* **Conmutación por Falla (Hot Failover)**: Sistema de contingencia de API en caliente que conmuta automáticamente a una clave de respaldo (Backup Key) en caso de saturación o límite de cuota en el proveedor principal.

---

## 🛠️ Stack Tecnológico

* **Frontend**: Next.js 14 (React, TypeScript, Tailwind CSS, KaTeX para renderizado de ecuaciones).
* **Backend**: FastAPI (Python 3.14, Uvicorn).
* **Procesamiento de Documentos**: python-docx, openpyxl, python-pptx, PyPDF.
* **Renderizado Documental**: LibreOffice Headless (dentro de contenedor Linux) para la conversión limpia de formatos.
* **Modelos de IA**: Google Gemini API (vía SDK nativo de Google GenAI).
* **Infraestructura y Despliegue**: Docker, Vercel (Frontend) y Render (Backend).

---

## 📁 Estructura del Repositorio

El proyecto está organizado en un monorrepositorio estructurado de la siguiente manera:

* `auto-task-frontend/`: Aplicación cliente desarrollada en Next.js que consume la API del backend de forma asíncrona.
* `auto-task-backend/`: Servidor de FastAPI encargado de la lógica de negocio, integración con APIs de IA y generación de archivos. Contiene el `Dockerfile` y los requerimientos del sistema.

---

## 🔗 Demo en Vivo

Puedes probar la aplicación desplegada en producción a través del siguiente enlace:
👉 **[auto-task-generador.vercel.app](https://auto-task-generador.vercel.app/)**

*(Nota: Al tratarse de una demo en el plan gratuito de Render, el backend puede demorar unos 50 segundos en responder en la primera petición si el servidor se encuentra inactivo).*

## ⚠️ Descargo de Responsabilidad (Disclaimer)

Esta aplicación ha sido desarrollada únicamente con fines educativos y de demostración técnica de portafolio. No se promueve ni se incentiva el plagio académico; el uso del material autogenerado como entregable final sin revisión previa es responsabilidad exclusiva del usuario final.
