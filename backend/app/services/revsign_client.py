"""Cliente do RevSign (assinatura eletrônica), isolado atrás da interface `RevSignClient`.

O restante do sistema só conhece `RevSignClient`, `NovoDocumento`, `StatusDocumento` e
`EventoWebhook`; o contrato HTTP fica todo neste módulo. A integração só é usada com
`REVSIGN_ENABLED=true` (padrão `false`).

PREMISSAS DO CONTRATO (a confirmar/ajustar quando a sessão do RevSign entregar o contrato
definitivo; tudo que depende delas está concentrado em `HttpRevSignClient`):

1. API v2 do Documenso v2.21.1, rotas "document" (marcadas como deprecated em favor da
   Envelope API).
2. `POST /api/v2/document/create` multipart (`payload` JSON + `file` .docx, convertido pelo
   Gotenberg do RevSign) cria o rascunho e `POST /api/v2/document/distribute` envia o e-mail.
3. Autenticação `Authorization: Bearer <REVSIGN_API_TOKEN>`.
4. Campo de assinatura obrigatório (o Documenso recusa enviar sem campo `SIGNATURE`): posição
   fixa `CAMPO_ASSINATURA` na página 1 até o contrato dizer outra coisa.
5. Signatário com e-mail pessoal depende da alteração do RevSign (spec, seção 2).
6. `GET /api/v2/document/{id}` traz `status` (`DRAFT|PENDING|COMPLETED|REJECTED|CANCELLED`),
   `completedAt` e `recipients[]` (`id`, `signingStatus`, `signedAt`, `rejectionReason`,
   `expiresAt`).
7. `GET /api/v2/document/{id}/download` devolve o PDF assinado.
8. `POST /api/v2/document/delete` cancela/apaga; `POST /api/v2/document/redistribute` reenvia.
9. Webhook com corpo `{event, payload: {id, externalId, ...}, createdAt, webhookEndpoint}` e
   cabeçalho `X-Documenso-Secret`; eventos úteis `DOCUMENT_COMPLETED`, `DOCUMENT_REJECTED`,
   `DOCUMENT_CANCELLED`, `RECIPIENT_EXPIRED`.
10. O Documenso recusa URL de webhook privada: cadastrar a URL pública
    `<APP_PUBLIC_URL>/api/integrations/revsign/webhook`.

Nenhuma mensagem de erro deste módulo inclui o token da API.
"""
import hmac
import json
import logging
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Literal, Mapping, Optional, Protocol

import httpx

logger = logging.getLogger(__name__)

PDF_FAKE = b"%PDF-1.4\n% termo assinado (falso)\n"
CABECALHO_SEGREDO = "X-Documenso-Secret"
API_PREFIXO = "/api/v2"
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
CAMPO_ASSINATURA = {"type": "SIGNATURE", "pageNumber": 1, "pageX": 10, "pageY": 80, "width": 35, "height": 8}

Situacao = Literal["pendente", "assinado", "recusado", "expirado", "cancelado"]


class RevSignError(Exception):
    """Falha ao falar com o RevSign (mensagem segura para log e para o usuário)."""


class WebhookInvalido(RevSignError):
    """Webhook sem o segredo correto."""


class WebhookMalformado(RevSignError):
    """Webhook autenticado, mas com corpo ilegível."""


@dataclass(frozen=True)
class NovoDocumento:
    titulo: str
    arquivo: bytes
    nome_arquivo: str
    signatario_nome: str
    signatario_email: str
    referencia_externa: str


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
    documento_id: Optional[str]
    referencia_externa: Optional[str]


class RevSignClient(Protocol):
    def criar_documento(self, doc: NovoDocumento) -> DocumentoCriado: ...

    def consultar(self, documento_id: str) -> StatusDocumento: ...

    def cancelar(self, documento_id: str) -> None: ...

    def reenviar(self, documento_id: str) -> None: ...

    def baixar_assinado(self, documento_id: str) -> bytes: ...

    def validar_webhook(self, headers: Mapping[str, str], corpo: bytes) -> EventoWebhook: ...


class _ValidadorWebhook:
    """Validação do webhook, comum ao cliente HTTP e ao falso.

    O corpo é dado não confiável: só `event`, `payload.id` e `payload.externalId` são lidos,
    e o status que ele eventualmente traga nunca é usado (a fonte da verdade é `consultar`).
    """

    def __init__(self, webhook_secret: str):
        self._webhook_secret = webhook_secret or ""

    def validar_webhook(self, headers: Mapping[str, str], corpo: bytes) -> EventoWebhook:
        if not self._webhook_secret:
            raise WebhookInvalido("Segredo do webhook não configurado.")
        recebido = None
        for nome, valor in headers.items():
            if nome.lower() == CABECALHO_SEGREDO.lower():
                recebido = valor
                break
        if recebido is None or not hmac.compare_digest(
            str(recebido).encode("utf-8"), self._webhook_secret.encode("utf-8")
        ):
            raise WebhookInvalido("Segredo do webhook inválido.")
        try:
            dados = json.loads(corpo)
        except (ValueError, UnicodeDecodeError):
            raise WebhookMalformado("Corpo do webhook inválido.") from None
        if not isinstance(dados, dict) or not isinstance(dados.get("event"), str):
            raise WebhookMalformado("Corpo do webhook inválido.")
        payload = dados.get("payload")
        documento_id: Optional[str] = None
        referencia: Optional[str] = None
        if isinstance(payload, dict):
            bruto = payload.get("id")
            if isinstance(bruto, (int, str)) and not isinstance(bruto, bool):
                documento_id = str(bruto)
            externo = payload.get("externalId")
            if isinstance(externo, str):
                referencia = externo
        return EventoWebhook(dados["event"], documento_id, referencia)


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

    def cancelar(self, documento_id: str) -> None:
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


def _como_id(documento_id: str) -> "int | str":
    texto = str(documento_id)
    return int(texto) if texto.isdigit() else texto


def _parse_data(texto: Any) -> Optional[datetime]:
    if not isinstance(texto, str) or not texto:
        return None
    try:
        dt = datetime.fromisoformat(texto.replace("Z", "+00:00"))
    except ValueError:
        return None
    return dt if dt.tzinfo is not None else dt.replace(tzinfo=timezone.utc)


class HttpRevSignClient(_ValidadorWebhook):
    API_PREFIXO = API_PREFIXO
    TIMEOUT = httpx.Timeout(15.0, connect=5.0)
    TIMEOUT_DOWNLOAD = httpx.Timeout(30.0, connect=5.0)
    MAX_PDF_BYTES = 20 * 1024 * 1024
    MIME_DOCX = MIME_DOCX
    CAMPO_ASSINATURA = CAMPO_ASSINATURA

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
        aceitar_404: bool = False,
        timeout: Optional[httpx.Timeout] = None,
        **kwargs: Any,
    ) -> httpx.Response:
        self._exigir_configuracao()
        url = f"{self._base_url}{self.API_PREFIXO}{caminho}"
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
        if resp.status_code == 404 and aceitar_404:
            return resp
        if resp.status_code >= 400:
            logger.warning("RevSign respondeu %s em %s %s.", resp.status_code, metodo, caminho)
            raise RevSignError(f"RevSign respondeu {resp.status_code}")
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

    # -- operações -----------------------------------------------------------

    def criar_documento(self, doc: NovoDocumento) -> DocumentoCriado:
        self._exigir_configuracao()
        payload = {
            "title": doc.titulo,
            "externalId": doc.referencia_externa,
            "recipients": [
                {
                    "email": doc.signatario_email,
                    "name": doc.signatario_nome,
                    "role": "SIGNER",
                    "signingOrder": 1,
                    "fields": [self.CAMPO_ASSINATURA],
                }
            ],
            "meta": {
                "subject": f"Assinatura: {doc.titulo}",
                "message": "Revalle · Patrimônio: assine o termo do equipamento sob sua responsabilidade.",
                "timezone": "America/Sao_Paulo",
                "language": "pt-BR",
            },
        }
        resp = self._requisitar(
            "POST",
            "/document/create",
            data={"payload": json.dumps(payload, ensure_ascii=False)},
            files={"file": (doc.nome_arquivo, doc.arquivo, self.MIME_DOCX)},
        )
        bruto = self._json(resp).get("id")
        if not isinstance(bruto, (int, str)) or isinstance(bruto, bool) or str(bruto) == "":
            raise RevSignError("Resposta inválida do RevSign.")
        documento_id = str(bruto)
        try:
            self._requisitar("POST", "/document/distribute", json={"documentId": _como_id(documento_id)})
        except RevSignError:
            # Sem distribuir, o rascunho ficaria órfão no RevSign: apaga em melhor esforço.
            try:
                self._requisitar(
                    "POST", "/document/delete", json={"documentId": _como_id(documento_id)}, aceitar_404=True
                )
            except RevSignError:
                logger.warning("Não foi possível apagar o rascunho %s no RevSign.", documento_id)
            raise
        return DocumentoCriado(documento_id)

    def _obter(self, documento_id: str) -> dict:
        return self._json(self._requisitar("GET", f"/document/{documento_id}"))

    def consultar(self, documento_id: str) -> StatusDocumento:
        dados = self._obter(documento_id)
        status = str(dados.get("status") or "").upper()
        destinatarios = [r for r in (dados.get("recipients") or []) if isinstance(r, dict)]
        if status == "COMPLETED":
            return StatusDocumento("assinado", assinado_em=_parse_data(dados.get("completedAt")))
        if status == "REJECTED":
            motivo = next(
                (r.get("rejectionReason") for r in destinatarios if r.get("rejectionReason")), None
            )
            return StatusDocumento("recusado", motivo=motivo)
        if status == "CANCELLED":
            return StatusDocumento("cancelado")
        agora = datetime.now(timezone.utc)
        for r in destinatarios:
            validade = _parse_data(r.get("expiresAt"))
            if r.get("signingStatus") == "NOT_SIGNED" and validade is not None and validade < agora:
                return StatusDocumento("expirado")
        return StatusDocumento("pendente")

    def cancelar(self, documento_id: str) -> None:
        self._requisitar(
            "POST", "/document/delete", json={"documentId": _como_id(documento_id)}, aceitar_404=True
        )

    def reenviar(self, documento_id: str) -> None:
        dados = self._obter(documento_id)
        pendentes = [
            r["id"]
            for r in (dados.get("recipients") or [])
            if isinstance(r, dict) and r.get("signingStatus") == "NOT_SIGNED" and "id" in r
        ]
        if not pendentes:
            raise RevSignError("Nenhum destinatário pendente de assinatura.")
        self._requisitar(
            "POST",
            "/document/redistribute",
            json={"documentId": _como_id(documento_id), "recipients": pendentes},
        )

    def baixar_assinado(self, documento_id: str) -> bytes:
        resp = self._requisitar(
            "GET", f"/document/{documento_id}/download", timeout=self.TIMEOUT_DOWNLOAD
        )
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
