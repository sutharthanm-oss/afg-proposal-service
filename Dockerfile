FROM python:3.11-slim

# LibreOffice is needed for PPTX -> PDF conversion (soffice --headless)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-crosextra-carlito \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Railway sets $PORT; default to 8000 for local runs
ENV PORT=8000
CMD uvicorn main:app --host 0.0.0.0 --port $PORT
