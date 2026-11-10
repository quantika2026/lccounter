# LCCOUNTER en la nube (Render u otro servicio con Docker).
FROM python:3.13-slim

# LibreOffice: PDF, recálculo de fórmulas e imágenes de las diapositivas.
RUN apt-get update \
 && apt-get install -y --no-install-recommends libreoffice-calc libreoffice-impress libreoffice-writer \
      fonts-dejavu fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app/codigo
COPY codigo/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY codigo/ .

# El espacio de trabajo (clientes, usuarios, expedientes) vive en el disco persistente /data.
ENV REVISORIA_RAIZ=/data \
    LCCOUNTER_HOST=0.0.0.0 \
    LCCOUNTER_COOKIE_SEGURA=1 \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp
EXPOSE 10000
CMD ["sh", "-c", "python -m app.servidor --raiz /data --puerto ${PORT:-10000}"]
