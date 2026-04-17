FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Cria os diretórios de storage local (usado em dev; em produção use S3)
RUN mkdir -p storage/terms storage/termos_assinados storage/termos_devolucao \
    storage/termos_devolucao_assinados storage/notas_remocao

EXPOSE 8000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "2"]
