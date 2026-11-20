FROM python:3.12-slim

WORKDIR /app

# LibreOffice (conversão do termo .docx para PDF, exigida pela API do RevSign) e fontes
# metricamente compatíveis com as do Word (Calibri/Cambria/Arial/Times).
RUN apt-get update \
    && apt-get install -y --no-install-recommends tini libreoffice-writer-nogui fonts-dejavu fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Cria os diretórios de storage local (usado em dev; em produção use S3)
RUN mkdir -p storage/terms storage/termos_assinados storage/termos_devolucao \
    storage/termos_devolucao_assinados storage/notas_remocao

EXPOSE 8000

# tini como PID 1 recolhe processos órfãos (ex.: soffice.bin de uma conversão que estourou o tempo).
ENTRYPOINT ["/usr/bin/tini", "--"]
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "2"]
