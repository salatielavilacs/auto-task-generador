# 1. Usamos una imagen ultra-ligera basada en Alpine o Slim (ahorra 300MB de entrada)
FROM python:3.14-slim

# 2. Variables de entorno para evitar archivos temporales de Python (.pyc)
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# 3. Instalar LibreOffice limpiando la caché de instalación inmediatamente
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-java-common \
    curl \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

# 4. Instalar librerías sin guardar la caché de pip (ahorra cientos de MB)
COPY requirements.txt .
RUN pip install --upgrade pip && pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 7860

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "7860"]
