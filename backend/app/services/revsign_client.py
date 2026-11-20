"""Cliente do RevSign (assinatura eletrônica), isolado atrás da interface `RevSignClient`.

O restante do sistema só conhece `RevSignClient`, `NovoDocumento`, `StatusDocumento` e
`EventoWebhook`; o contrato HTTP fica todo neste módulo. A integração só é usada com
`REVSIGN_ENABLED=true` (padrão `false`).

Contrato: API dedicada do RevSign para o Patrimônio
(`docs/integracoes/revsign-contrato-patrimonio.md`).

- `REVSIGN_BASE_URL` é a base completa (ex.: `.../api/revalle/patrimonio/v1`); o cliente só
  acrescenta `/termos…`.
- Autenticação `Authorization: Bearer <REVSIGN_API_TOKEN>`.
- `POST /termos` (multipart, PDF) cria e envia; `GET /termos/{id}` consulta;
  `POST /termos/{id}/cancelar` e `/reenviar`; `GET /termos/{id}/pdf` baixa o PDF selado.
- Erros: `{"error": {"code", "message"}}`; a mensagem mostrada ao usuário vem de um mapa fixo
  (a do servidor nunca é repassada).
- Webhook: HMAC-SHA256 (`sha256=<hex>`) de `timestamp + "." + corpo bruto` com o segredo
  compartilhado, nos cabeçalhos `X-RevSign-Timestamp` e `X-RevSign-Signature`, janela de 300 s.

Nenhuma mensagem de erro deste módulo inclui o token da API nem o segredo do webhook.
"""
import hashlib
import hmac
import json
import logging
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Literal, Mapping, Optional, Protocol
from urllib.parse import quote

import httpx

logger = logging.getLogger(__name__)

PDF_FAKE = b"%PDF-1.4\n% termo assinado (falso)\n"
CABECALHO_TIMESTAMP = "X-RevSign-Timestamp"
CABECALHO_ASSINATURA = "X-RevSign-Signature"
JANELA_WEBHOOK_SEGUNDOS = 300
MENSAGEM_CONVITE = "Revalle · Patrimônio: assine o termo do equipamento sob sua responsabilidade."

Situacao = Literal["pendente", "assinado", "recusado", "expirado", "cancelado"]

_SITUACAO_POR_STATUS: dict[str, Situacao] = {
    "PENDING": "pendente",
    "SIGNED": "assinado",
    "REJECTED": "recusado",
    "EXPIRED": "expirado",
    "CANCELLED": "cancelado",
}

_MENSAGEM_POR_CODIGO = {
    "UNAUTHORIZED": "Chave da integração com o RevSign recusada.",
    "NOT_CONFIGURED": "Integração desligada no RevSign.",
    "INVALID_CPF": "CPF do colaborador recusado pelo RevSign.",
    "COLABORADOR_NOT_FOUND": "Colaborador não está na base do RH do RevSign.",
    "HR_BASE_UNAVAILABLE": "Base do RH do RevSign fora do ar.",
    "MARKER_NOT_FOUND": "O termo não tem o marcador de assinatura.",
    "INVALID_FILE": "O RevSign recusou o PDF do termo.",
    "FILE_TOO_LARGE": "PDF do termo maior que 10 MB.",
    "RESEND_TOO_SOON": "Aguarde 10 minutos entre reenvios.",
    "NOT_PENDING": "O termo já não está pendente no RevSign.",
    "RATE_LIMITED": "RevSign ocupado; tente de novo em instantes.",
}


class RevSignError(Exception):
    """Falha ao falar com o RevSign (mensagem segura para log e para o usuário).

    `codigo` é o `error.code` do contrato, quando houver; `str(e)` é a mensagem em português.
    """

    def __init__(self, mensagem: str, codigo: Optional[str] = None, *, termo_id: Optional[str] = None):
        super().__init__(mensagem)
        self.codigo = codigo
        self.termo_id = termo_id  # `error.termoId` do 409 DUPLICATE_EXTERNAL_ID


class WebhookInvalido(RevSignError):
    """Webhook sem assinatura válida (ou fora da janela de tempo)."""


class WebhookMalformado(RevSignError):
    """Webhook autenticado, mas com corpo ilegível."""


@dataclass(frozen=True)
class NovoDocumento:
    titulo: str
    arquivo: bytes  # PDF
    nome_arquivo: str  # termina em .pdf
    signatario_nome: str
    signatario_email: str
    referencia_externa: str
    tipo: Literal["ENTREGA", "DEVOLUCAO"]
    signatario_cpf: str  # só dígitos


@dataclass(frozen=True)
class DocumentoCriado:
    documento_id: str


@dataclass(frozen=True)
class StatusDocumento:
    situacao: Situacao
    assinado_em: Optional[datetime] = None
    motivo: Optional[str] = None


@dataclass(frozen=True)
class EventoWebhook:
    tipo: str
    documento_id: Optional[str]  # `termoId`
    referencia_externa: Optional[str]  # `externalId`


def assinar_webhook(segredo: str, corpo: bytes, timestamp: int) -> dict[str, str]:
    """Cabeçalhos que o RevSign envia no webhook (usado pelos testes e pelo falso)."""
    mac = hmac.new(segredo.encode("utf-8"), str(timestamp).encode() + b"." + corpo, hashlib.sha256)
    return {CABECALHO_TIMESTAMP: str(timestamp), CABECALHO_ASSINATURA: "sha256=" + mac.hexdigest()}


class RevSignClient(Protocol):
    def criar_documento(self, doc: NovoDocumento) -> DocumentoCriado: ...

    def consultar(self, documento_id: str) -> StatusDocumento: ...

    def cancelar(self, documento_id: str, motivo: Optional[str] = None) -> None: ...

    def reenviar(self, documento_id: str) -> None: ...

    def baixar_assinado(self, documento_id: str) -> bytes: ...

    def validar_webhook(self, headers: Mapping[str, str], corpo: bytes) -> EventoWebhook: ...


class _ValidadorWebhook:
    """Validação do webhook, comum ao cliente HTTP e ao falso.

    Confere o HMAC-SHA256 de `timestamp + "." + corpo bruto` e a janela de 300 s. O corpo
    é dado não confiável: só `event`, `termoId` e `externalId` são lidos, e o status que ele
    eventualmente traga nunca é usado (a fonte da verdade é `consultar`).
    """

    def __init__(self, webhook_secret: str):
        self._webhook_secret = webhook_secret or ""

    @staticmethod
    def _cabecalho(headers: Mapping[str, str], nome: str) -> Optional[str]:
        for chave, valor in headers.items():
            if chave.lower() == nome.lower():
                return str(valor)
        return None

    def validar_webhook(self, headers: Mapping[str, str], corpo: bytes) -> EventoWebhook:
        if not self._webhook_secret:
            raise WebhookInvalido("Segredo do webhook não configurado.")
        timestamp = self._cabecalho(headers, CABECALHO_TIMESTAMP)
        assinatura = self._cabecalho(headers, CABECALHO_ASSINATURA)
        if timestamp is None or assinatura is None:
            raise WebhookInvalido("Assinatura do webhook ausente.")
        if not (timestamp.isascii() and timestamp.isdigit()):
            raise WebhookInvalido("Assinatura do webhook inválida.")
        esperado = assinar_webhook(self._webhook_secret, corpo, int(timestamp))[CABECALHO_ASSINATURA]
        # O HMAC é sempre conferido (tempo constante), com ou sem janela válida.
        assinatura_ok = hmac.compare_digest(esperado.encode("utf-8"), assinatura.encode("utf-8"))
        dentro_da_janela = abs(time.time() - int(timestamp)) <= JANELA_WEBHOOK_SEGUNDOS
        if not (assinatura_ok and dentro_da_janela):
            raise WebhookInvalido("Assinatura do webhook inválida.")
        try:
            dados = json.loads(corpo)
        except (ValueError, UnicodeDecodeError):
            raise WebhookMalformado("Corpo do webhook inválido.") from None
        if not isinstance(dados, dict) or not isinstance(dados.get("event"), str):
            raise WebhookMalformado("Corpo do webhook inválido.")
        termo_id = dados.get("termoId")
        externo = dados.get("externalId")
        return EventoWebhook(
            dados["event"],
            termo_id if isinstance(termo_id, str) and termo_id else None,
            externo if isinstance(externo, str) else None,
        )


# ── Implementação falsa (testes) ──────────────────────────────────────────────


class FakeRevSignClient(_ValidadorWebhook):
    def __init__(self, webhook_secret: str = ""):
        super().__init__(webhook_secret)
        self.criados: list[NovoDocumento] = []
        self.cancelados: list[str] = []
        self.reenviados: list[str] = []
        self.falhar_em: set[str] = set()
        self._docs: dict[str, dict[str, Any]] = {}
        self._proximo = 1001

    def _checar(self, metodo: str) -> None:
        if metodo in self.falhar_em:
            raise RevSignError(f"Falha simulada em {metodo}.")

    def _doc(self, documento_id: str) -> dict[str, Any]:
        doc = self._docs.get(str(documento_id))
        if doc is None:
            raise RevSignError("Documento não encontrado no RevSign.")
        return doc

    def criar_documento(self, doc: NovoDocumento) -> DocumentoCriado:
        self._checar("criar_documento")
        documento_id = str(self._proximo)
        self._proximo += 1
        self.criados.append(doc)
        self._docs[documento_id] = {
            "situacao": "pendente", "pdf": None, "assinado_em": None, "motivo": None,
        }
        return DocumentoCriado(documento_id)

    def consultar(self, documento_id: str) -> StatusDocumento:
        self._checar("consultar")
        d = self._doc(documento_id)
        return StatusDocumento(d["situacao"], d["assinado_em"], d["motivo"])

    def cancelar(self, documento_id: str, motivo: Optional[str] = None) -> None:
        self._checar("cancelar")
        self.cancelados.append(str(documento_id))

    def reenviar(self, documento_id: str) -> None:
        self._checar("reenviar")
        self.reenviados.append(str(documento_id))

    def baixar_assinado(self, documento_id: str) -> bytes:
        self._checar("baixar_assinado")
        d = self._doc(documento_id)
        if d["situacao"] != "assinado" or d["pdf"] is None:
            raise RevSignError("Documento ainda não assinado.")
        return d["pdf"]

    # ajudantes de teste
    def marcar_assinado(
        self, documento_id: str, pdf: bytes = PDF_FAKE, assinado_em: Optional[datetime] = None
    ) -> None:
        d = self._doc(documento_id)
        d["situacao"] = "assinado"
        d["pdf"] = pdf
        d["assinado_em"] = assinado_em or datetime.now(timezone.utc)

    def marcar(self, documento_id: str, situacao: Situacao, motivo: Optional[str] = None) -> None:
        d = self._doc(documento_id)
        d["situacao"] = situacao
        d["motivo"] = motivo


# ── Implementação HTTP ────────────────────────────────────────────────────────


def _parse_data(texto: Any) -> Optional[datetime]:
    if not isinstance(texto, str) or not texto:
        return None
    try:
        dt = datetime.fromisoformat(texto.replace("Z", "+00:00"))
    except ValueError:
        return None
    return dt if dt.tzinfo is not None else dt.replace(tzinfo=timezone.utc)


def _erro_da_resposta(resp: httpx.Response) -> RevSignError:
    """Converte uma resposta de erro do contrato em `RevSignError` (mensagem do nosso mapa)."""
    codigo: Optional[str] = None
    termo_id: Optional[str] = None
    try:
        corpo = resp.json()
    except ValueError:
        corpo = None
    erro = corpo.get("error") if isinstance(corpo, dict) else None
    if isinstance(erro, dict):
        if isinstance(erro.get("code"), str):
            codigo = erro["code"]
        if isinstance(erro.get("termoId"), str):
            termo_id = erro["termoId"]
    mensagem = _MENSAGEM_POR_CODIGO.get(codigo or "") or f"RevSign respondeu {resp.status_code}."
    return RevSignError(mensagem, codigo, termo_id=termo_id)


class HttpRevSignClient(_ValidadorWebhook):
    TIMEOUT = httpx.Timeout(15.0, connect=5.0)
    TIMEOUT_DOWNLOAD = httpx.Timeout(30.0, connect=5.0)
    MAX_PDF_BYTES = 20 * 1024 * 1024

    def __init__(
        self,
        base_url: str,
        api_token: str,
        webhook_secret: str,
        *,
        transport: Optional[httpx.BaseTransport] = None,
    ):
        super().__init__(webhook_secret)
        self._base_url = (base_url or "").rstrip("/")
        self._api_token = api_token or ""
        self._transport = transport

    # -- infraestrutura ------------------------------------------------------

    def _exigir_configuracao(self) -> None:
        if not self._base_url or not self._api_token:
            raise RevSignError("RevSign não configurado (REVSIGN_BASE_URL/REVSIGN_API_TOKEN).")

    def _requisitar(
        self,
        metodo: str,
        caminho: str,
        *,
        timeout: Optional[httpx.Timeout] = None,
        **kwargs: Any,
    ) -> httpx.Response:
        self._exigir_configuracao()
        url = f"{self._base_url}{caminho}"
        try:
            with httpx.Client(
                transport=self._transport,
                timeout=timeout or self.TIMEOUT,
                headers={"Authorization": f"Bearer {self._api_token}"},
            ) as http:
                resp = http.request(metodo, url, **kwargs)
        except httpx.HTTPError as e:
            # Só o tipo vai para o log/mensagem: o texto do erro pode conter a URL.
            logger.warning("RevSign indisponível (%s) em %s %s.", type(e).__name__, metodo, caminho)
            raise RevSignError("RevSign indisponível.") from None
        if resp.status_code >= 400:
            erro = _erro_da_resposta(resp)
            logger.warning(
                "RevSign respondeu %s (%s) em %s %s.", resp.status_code, erro.codigo or "sem código", metodo, caminho
            )
            raise erro
        return resp

    @staticmethod
    def _json(resp: httpx.Response) -> dict:
        try:
            dados = resp.json()
        except ValueError:
            raise RevSignError("Resposta inválida do RevSign.") from None
        if not isinstance(dados, dict):
            raise RevSignError("Resposta inválida do RevSign.")
        return dados

    @staticmethod
    def _termo(documento_id: str) -> str:
        return f"/termos/{quote(str(documento_id), safe='')}"

    # -- operações -----------------------------------------------------------

    def criar_documento(self, doc: NovoDocumento) -> DocumentoCriado:
        self._exigir_configuracao()
        try:
            resp = self._requisitar(
                "POST",
                "/termos",
                data={
                    "externalId": doc.referencia_externa,
                    "kind": doc.tipo,
                    "title": doc.titulo,
                    "signerName": doc.signatario_nome,
                    "signerEmail": doc.signatario_email,
                    "signerCpf": doc.signatario_cpf,
                    "message": MENSAGEM_CONVITE,
                },
                files={"file": (doc.nome_arquivo, doc.arquivo, "application/pdf")},
            )
        except RevSignError as e:
            # Criar é idempotente: depois de uma falha de rede o termo pode já existir.
            if e.codigo == "DUPLICATE_EXTERNAL_ID" and e.termo_id:
                return DocumentoCriado(e.termo_id)
            raise
        termo_id = self._json(resp).get("termoId")
        if not isinstance(termo_id, str) or not termo_id:
            raise RevSignError("Resposta inválida do RevSign.")
        return DocumentoCriado(termo_id)

    def consultar(self, documento_id: str) -> StatusDocumento:
        dados = self._json(self._requisitar("GET", self._termo(documento_id)))
        situacao = _SITUACAO_POR_STATUS.get(str(dados.get("status") or "").upper())
        if situacao is None:
            raise RevSignError("Resposta inválida do RevSign.")
        signatario = dados.get("signer") if isinstance(dados.get("signer"), dict) else {}
        if situacao == "assinado":
            assinado_em = _parse_data(signatario.get("signedAt")) or _parse_data(dados.get("completedAt"))
            return StatusDocumento("assinado", assinado_em=assinado_em)
        if situacao == "recusado":
            motivo = signatario.get("rejectionReason")
            return StatusDocumento("recusado", motivo=motivo if isinstance(motivo, str) and motivo else None)
        return StatusDocumento(situacao)

    def cancelar(self, documento_id: str, motivo: Optional[str] = None) -> None:
        kwargs: dict[str, Any] = {"json": {"reason": motivo[:500]}} if motivo else {}
        try:
            self._requisitar("POST", f"{self._termo(documento_id)}/cancelar", **kwargs)
        except RevSignError as e:
            if e.codigo != "NOT_FOUND":  # termo que já não existe = cancelado
                raise

    def reenviar(self, documento_id: str) -> None:
        self._requisitar("POST", f"{self._termo(documento_id)}/reenviar")

    def baixar_assinado(self, documento_id: str) -> bytes:
        resp = self._requisitar("GET", f"{self._termo(documento_id)}/pdf", timeout=self.TIMEOUT_DOWNLOAD)
        conteudo = resp.content
        if len(conteudo) > self.MAX_PDF_BYTES:
            raise RevSignError("PDF assinado maior que o limite permitido.")
        if not conteudo.startswith(b"%PDF"):
            raise RevSignError("O RevSign não devolveu um PDF.")
        return conteudo


# ── Fábrica ───────────────────────────────────────────────────────────────────

_cliente_de_teste: Optional[RevSignClient] = None


def usar_cliente(cliente: Optional[RevSignClient]) -> None:
    """Substitui o cliente devolvido por `get_revsign_client` (só testes). `None` restaura."""
    global _cliente_de_teste
    _cliente_de_teste = cliente


def get_revsign_client() -> RevSignClient:
    if _cliente_de_teste is not None:
        return _cliente_de_teste
    # Lido no momento da chamada: testes trocam `app.core.config.settings`.
    from app.core import config as _config

    cfg = _config.settings
    return HttpRevSignClient(cfg.REVSIGN_BASE_URL, cfg.REVSIGN_API_TOKEN, cfg.REVSIGN_WEBHOOK_SECRET)
