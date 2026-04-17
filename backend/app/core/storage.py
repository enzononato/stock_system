"""
Abstração de storage: LocalStorage (dev) e S3Storage (produção).
Troca transparente via variável de ambiente STORAGE_BACKEND.
"""
import os
import io
from typing import Protocol
from app.core.config import settings


class StorageBackend(Protocol):
    def save(self, category: str, filename: str, content: bytes) -> str:
        """Salva arquivo e retorna a chave (key) para recuperação posterior."""
        ...

    def read(self, key: str) -> bytes:
        """Lê arquivo pela chave retornada em save()."""
        ...

    def get_url(self, key: str) -> str:
        """Retorna URL de download (presigned para S3, path relativo para local)."""
        ...

    def exists(self, key: str) -> bool:
        ...


class LocalStorage:
    def __init__(self, base_path: str = None):
        self.base_path = base_path or settings.LOCAL_STORAGE_PATH
        # Garante que os subdiretórios existam
        for subdir in ["terms", "termos_assinados", "termos_devolucao",
                       "termos_devolucao_assinados", "notas_remocao"]:
            os.makedirs(os.path.join(self.base_path, subdir), exist_ok=True)

    def save(self, category: str, filename: str, content: bytes) -> str:
        path = os.path.join(self.base_path, category, filename)
        with open(path, "wb") as f:
            f.write(content)
        return f"{category}/{filename}"

    def read(self, key: str) -> bytes:
        path = os.path.join(self.base_path, key)
        with open(path, "rb") as f:
            return f.read()

    def get_url(self, key: str) -> str:
        # O frontend buscará via /api/documents/files/{key}
        return f"/api/documents/files/{key}"

    def exists(self, key: str) -> bool:
        return os.path.exists(os.path.join(self.base_path, key))


class S3Storage:
    def __init__(self):
        import boto3
        self.client = boto3.client(
            "s3",
            region_name=settings.S3_REGION,
            aws_access_key_id=settings.AWS_ACCESS_KEY_ID,
            aws_secret_access_key=settings.AWS_SECRET_ACCESS_KEY,
        )
        self.bucket = settings.S3_BUCKET

    def save(self, category: str, filename: str, content: bytes) -> str:
        key = f"{category}/{filename}"
        self.client.put_object(Bucket=self.bucket, Key=key, Body=content)
        return key

    def read(self, key: str) -> bytes:
        response = self.client.get_object(Bucket=self.bucket, Key=key)
        return response["Body"].read()

    def get_url(self, key: str) -> str:
        # URL presigned com 15 minutos de validade
        return self.client.generate_presigned_url(
            "get_object",
            Params={"Bucket": self.bucket, "Key": key},
            ExpiresIn=900,
        )

    def exists(self, key: str) -> bool:
        try:
            self.client.head_object(Bucket=self.bucket, Key=key)
            return True
        except Exception:
            return False


def get_storage() -> StorageBackend:
    if settings.STORAGE_BACKEND == "s3":
        return S3Storage()
    return LocalStorage()


# Instância global reutilizável
storage: StorageBackend = get_storage()
