# RevSign — adequar o Patrimônio ao contrato da integração — design

Data: 2026-11-20. Aprovado na conversa: termo de uso vai como `ENTREGA` ("pode seguir com a recomendação ali da opção A"); conta remetente `patrimonio@revalle.com.br`; entregar pronto para deploy.

## Intenção
O Patrimônio já tem a integração com o RevSign (F3), mas o cliente HTTP foi escrito com premissas da API genérica do Documenso, antes do contrato existir. O RevSign depois implementou uma API própria para o Patrimônio, documentada em `revalle/docs/patrimonio.md` do repositório `enzononato/revsign` (lido em 19/11). Ligar `REVSIGN_ENABLED=true` hoje não funciona. Esta entrega troca o lado do Patrimônio para o contrato real, sem mudar o fluxo de termos que já existe (solicitar, enviar, agendar, cancelar, reenviar, receber assinado, cair no manual quando falha).

## Diferenças entre o código atual e o contrato

| | Hoje (premissa) | Contrato |
|---|---|---|
| Base | `REVSIGN_BASE_URL` + `/api/v2/document/…` | `REVSIGN_BASE_URL` já é `…/api/revalle/patrimonio/v1`; rotas `/termos…` |
| Criar | multipart `payload` JSON + `.docx`, depois `distribute` | `POST /termos` multipart: `file` (PDF ≤10 MB), `externalId`, `kind` (`ENTREGA`/`DEVOLUCAO`), `title` (≤200), `signerName`, `signerEmail`, `signerCpf`, `message` (opcional, ≤2000). Resposta 201 com `termoId` |
| Repetição | — | 409 `DUPLICATE_EXTERNAL_ID` traz o `termoId` existente: tratar como sucesso |
| Assinatura | campo em posição fixa | marcador `{{assinatura}}` no PDF; sem ele 422 `MARKER_NOT_FOUND` |
| Estado | `status` DRAFT/PENDING/COMPLETED… | `GET /termos/{id}`: `PENDING`, `SIGNED`, `REJECTED` (`signer.rejectionReason`), `EXPIRED`, `CANCELLED`; `completedAt`, `signer.signedAt` |
| PDF | `/document/{id}/download` | `GET /termos/{id}/pdf` (só `SIGNED`), cabeçalho `X-RevSign-SHA256` |
| Cancelar | `POST /document/delete` | `POST /termos/{id}/cancelar` `{"reason"}`; 409 `NOT_PENDING` |
| Reenviar | `redistribute` com destinatários | `POST /termos/{id}/reenviar`; 429 `RESEND_TOO_SOON` |
| Webhook | cabeçalho `X-Documenso-Secret` = segredo | `X-RevSign-Signature: sha256=<hex>` = HMAC-SHA256(segredo, `timestamp + "." + corpo bruto`), `X-RevSign-Timestamp` (segundos; recusar se |agora − ts| > 300); corpo `{eventId, event, termoId, externalId, status, …}`; eventos `termo.sent/signed/rejected/expired/cancelled`; pode chegar repetido |
| Autenticação | Bearer token | Bearer com a chave `NEXT_PRIVATE_REVALLE_PATRIMONIO_API_KEY` do RevSign |

## Objetivos
- `HttpRevSignClient` fala o contrato acima. A interface `RevSignClient` (`criar_documento`, `consultar`, `cancelar`, `reenviar`, `baixar_assinado`, `validar_webhook`) continua a mesma para o resto do sistema; `NovoDocumento` ganha `tipo` (`ENTREGA`/`DEVOLUCAO`) e `signatario_cpf`.
- **Termo de uso vai como `ENTREGA`**; o título continua dizendo "Termo de uso" e o Patrimônio guarda o tipo real.
- **PDF**: o termo continua sendo gerado em `.docx` (download manual igual a hoje); para o RevSign, é convertido para PDF com LibreOffice headless dentro do container do backend. Falha ou demora (>60 s) na conversão → solicitação vai para `manual` com motivo claro.
- **Marcador**: todos os modelos em `backend/modelos/` ganham `{{assinatura}}` em fonte branca no lugar onde o colaborador assina, com espaço livre de ~6,5 × 2 cm à direita e abaixo, sem cobrir texto. A posição é conferida olhando o PDF gerado de cada modelo. O preenchimento atual não pode apagar nem trocar esse marcador, e o `.docx` baixado para assinatura manual não mostra nada visível (fonte branca).
- **CPF** do colaborador vai no envio (só dígitos). Sem CPF válido → `manual` com motivo "Colaborador sem CPF válido: assinatura manual."
- Erros do contrato viram motivos legíveis no termo (que já cai em `manual`): `COLABORADOR_NOT_FOUND` → "Colaborador não está na base do RH do RevSign: assinatura manual."; `INVALID_CPF`, `MARKER_NOT_FOUND`, `INVALID_FILE`, `FILE_TOO_LARGE`, `HR_BASE_UNAVAILABLE`, `NOT_CONFIGURED`, 401 com mensagens equivalentes. Nunca o token nos logs.
- Webhook com a assinatura HMAC nova. Evento repetido é inofensivo porque o processamento já consulta o estado no RevSign (`sincronizar`) e é idempotente; não precisa guardar `eventId`.
- O motivo de recusa (`rejectionReason`) continua indo para o termo.

## Fora
- Mudar o RevSign. Tipo `USO` no RevSign (opção B, recusada).
- Verificar o SHA-256 do PDF contra `/verificar` (o cabeçalho é guardado no log, não validado).

## O que tornaria isso um fracasso
- Um termo ficar sem caminho: se o RevSign recusar, o técnico tem que ver o motivo e conseguir seguir pelo manual, como hoje.
- Termo assinado no RevSign que não conclui o empréstimo/devolução no Patrimônio.
- O marcador aparecer visível no termo impresso, ou o campo de assinatura cobrir texto.
- Aceitar webhook forjado ou antigo.

## Abordagens consideradas
1. **Escolhida: reescrever só `HttpRevSignClient` + conversão para PDF + marcador nos modelos**, mantendo a interface. O resto (termos.py, termos_retorno.py, jobs) muda só no necessário (`tipo`, CPF, motivos).
2. Pedir ao RevSign para aceitar `.docx` e a premissa antiga: descartada, o contrato está fechado e foi feito para o Patrimônio.
3. Gerar o PDF direto (sem `.docx`, com fpdf2): descartada, duplicaria os 15 modelos de termo e mudaria o documento que já é usado.

## Configuração e deploy
- Variáveis do Patrimônio (nomes atuais mantidos): `REVSIGN_ENABLED`, `REVSIGN_BASE_URL` (passa a ser a base completa do contrato), `REVSIGN_API_TOKEN` (a chave do contrato), `REVSIGN_WEBHOOK_SECRET`. README e `.env.example` atualizados com os valores esperados.
- Rede interna do EasyPanel (recomendado pelo contrato): `REVSIGN_BASE_URL=http://agente_revalle_revsign:3000/api/revalle/patrimonio/v1` e, no RevSign, `NEXT_PRIVATE_REVALLE_PATRIMONIO_WEBHOOK_URL=http://agente_revalle_stock-sys-backend:8000/api/integrations/revsign/webhook`. Nomes de serviço confirmados no deploy.
- `backend/Dockerfile` instala `libreoffice-writer-nogui` e fontes (`fonts-dejavu`, `fonts-liberation`, `fonts-crosextra-carlito`, `fonts-crosextra-caladea`, para Calibri/Cambria).
- Ligar em produção depois do deploy, com um termo de teste mandado ao e-mail do responsável.

## Decidido por mim
- Timeout de conversão 60 s; perfil do LibreOffice em diretório temporário por conversão (permite duas conversões simultâneas nos 2 workers).
- 429 `RESEND_TOO_SOON` no reenvio vira mensagem "Aguarde 10 minutos entre reenvios." para o técnico.
- Cancelar termo que não existe no RevSign (404) conta como cancelado remotamente, como hoje. 409 `NOT_PENDING` (já assinado ou em selagem) é erro: o cancelamento remoto não aconteceu, e a consulta/webhook trazem o estado real.
- `message` do convite: "Revalle · Patrimônio: assine o termo do equipamento sob sua responsabilidade."
