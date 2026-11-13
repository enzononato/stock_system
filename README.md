# Controle de Estoque Revalle — Sistema Web

Aplicação web de controle de estoque de TI da Revalle: **FastAPI** (backend) +
**React** (frontend), substituindo o antigo sistema desktop em Tkinter. O
banco MySQL existente é reaproveitado sem alterações de schema.

---

## Estrutura

```
/
├── backend/           ← API FastAPI (Python)
│   ├── app/            ← Código da aplicação
│   ├── modelos/         ← Modelos .docx dos termos (versionados no git)
│   └── requirements*.txt
├── frontend/          ← SPA React (TypeScript + Vite + Shadcn/ui)
├── docs/              ← Documentação complementar
│   └── REMOCAO_HISTORICO_GIT.md
└── docker-compose.yml
```

---

## Como Rodar (Desenvolvimento Local)

### Backend

```bash
cd backend

# Crie o .env a partir do exemplo e preencha os valores reais
cp .env.example .env
# Edite o .env — cada variável está documentada com um comentário no próprio
# arquivo. DB_HOST, DB_USER, DB_PASSWORD, DB_NAME e JWT_SECRET são
# obrigatórios: o backend recusa subir sem eles (mensagem de erro clara).

# Instale as dependências (recomenda-se virtualenv)
pip install -r requirements.txt
# Para rodar testes, use o arquivo de dependências de desenvolvimento:
# pip install -r requirements-dev.txt

# Inicie o servidor
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Os modelos de termo (.docx) já estão versionados em `backend/modelos/` — não é
necessário copiá-los de nenhum outro lugar.

A documentação automática da API estará em: **http://localhost:8000/api/docs**

### Frontend

```bash
cd frontend

# Instale as dependências
npm install

# Inicie o servidor de desenvolvimento (proxy /api/ → localhost:8000)
npm run dev
```

Acesse: **http://localhost:5173**

---

## Deploy em Produção (Internet)

### Pré-requisitos
- Servidor Linux com Docker + Docker Compose
- Domínio apontando para o IP do servidor

### Passo a passo

**1. Clone o repositório no servidor**
```bash
git clone <seu-repo> /opt/revalle
cd /opt/revalle
```

**2. Configure o backend**
```bash
cp backend/.env.example backend/.env
nano backend/.env
```

Preencha pelo menos:
```env
DB_HOST=<endereço do seu servidor MySQL>
DB_USER=<usuário>
DB_PASSWORD=<senha>
DB_NAME=<banco>

# Gere com: python -c "import secrets; print(secrets.token_urlsafe(48))"
JWT_SECRET=<chave-aleatória-de-32+-caracteres>

ENVIRONMENT=production

# Para produção, prefira S3 em vez de storage local
STORAGE_BACKEND=s3
S3_BUCKET=<seu-bucket>
AWS_ACCESS_KEY_ID=<sua-key>
AWS_SECRET_ACCESS_KEY=<seu-secret>

# Domínio do frontend
CORS_ORIGINS=https://estoque.revalle.com.br
```

> Veja `backend/.env.example` para a lista completa de variáveis (pool de
> conexões do banco, rate limiting do login, paginação, cookies etc.) — cada
> uma está documentada com um comentário no próprio arquivo.

**3. Suba os containers**
```bash
docker compose up -d --build
```

O backend não publica porta no host (`expose`, não `ports`): só é acessível
através do proxy reverso do Nginx no serviço `frontend`.

**4. Configure HTTPS com Caddy (recomendado — TLS automático)**

Instale Caddy no servidor:
```bash
apt install -y debian-keyring debian-archive-keyring apt-transport-https
curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' | tee /etc/apt/sources.list.d/caddy-stable.list
apt update && apt install caddy
```

Crie `/etc/caddy/Caddyfile`:
```
estoque.revalle.com.br {
    reverse_proxy localhost:80
}
```

```bash
systemctl enable --now caddy
```

---

## Variáveis de ambiente novas

Acrescentadas nesta rodada de melhorias. A lista completa, com comentários, está em
`backend/.env.example`; aqui fica o resumo do que cada uma faz. Nenhum valor real de
segredo vai para o repositório: senhas, tokens e chaves só existem no `.env` do servidor.

| Variável | Para que serve |
|----------|----------------|
| `APP_PUBLIC_URL` | URL pública do frontend (ex.: `https://patrimonio.revalle.com.br`, sem `/` no fim). Monta os links dos e-mails (redefinição de senha, termos) **e é a base do QR code das etiquetas** (`APP_PUBLIC_URL/e/{id}`): com valor errado, as etiquetas impressas apontam para lugar nenhum e precisam ser reimpressas. |
| `MAIL_BACKEND` | `smtp` envia de verdade; `memory` só guarda em memória (testes). |
| `SMTP_HOST`, `SMTP_PORT` | Servidor de e-mail (`mail.revalle.com.br`, porta `465`, SSL implícito). |
| `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM` | Conta que envia os e-mails; a senha só vai no `.env` do servidor. `SMTP_FROM` vazio usa `SMTP_USER`. |
| `SMTP_VERIFY_TLS` | Verifica o certificado do servidor SMTP (padrão `true`; só desligue se o certificado não bater com o nome do host). |
| `COLABORADORES_DATABASE_URL` | Conexão (somente leitura) com o Postgres da base de RH. Use um usuário de leitura. |
| `COLABORADORES_TABLE`, `COLABORADORES_SCHEMA` | Tabela e schema dos colaboradores na base de RH. |
| `COLABORADORES_COLUMNS` | Mapeamento opcional de colunas; vazio usa o padrão. |
| `COLABORADORES_SYNC_HORA` | Horário (`HH:MM`) da sincronização diária de colaboradores. |
| `SCHEDULER_ENABLED` | `False` desliga o agendador de tarefas (sincronização, envio agendado de termos, consulta ao RevSign). |
| `TOTP_ENCRYPTION_KEY` | Chave Fernet que cifra o segredo do 2FA no banco. Gere com `python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"`. **Vazia = 2FA desligado.** Trocar a chave invalida os autenticadores já cadastrados. Com a chave definida, todo Gestor é levado a cadastrar o 2FA no primeiro login. |
| `REVSIGN_ENABLED` | Padrão `False`: a integração com o RevSign fica desligada até o contrato chegar, e os termos seguem o fluxo manual (baixar, assinar, anexar o PDF). |
| `REVSIGN_BASE_URL`, `REVSIGN_API_TOKEN` | Endereço e token da API do RevSign (token é segredo). |
| `REVSIGN_WEBHOOK_SECRET` | Segredo que valida os webhooks recebidos. URL do webhook a cadastrar no RevSign: `<APP_PUBLIC_URL>/api/integrations/revsign/webhook`. |
| `ALERTA_GARANTIA_DIAS` | Garantia que vence em até N dias entra nos alertas do painel (padrão `30`). |
| `ALERTA_TERMO_PENDENTE_DIAS` | Termo sem assinatura há mais de N dias (padrão `3`). |
| `ALERTA_DEVOLUCAO_PENDENTE_DIAS` | Devolução pendente há mais de N dias (padrão `7`). |
| `ALERTA_MANUTENCAO_DIAS` | Manutenção parada há mais de N dias (padrão `15`). |

Os alertas do painel são calculados na hora, a cada consulta; mudar um `ALERTA_*` no `.env`
vale depois de reiniciar o backend e não grava nada no banco.

---

## Deploy no EasyPanel

- **Rede**: o container do backend precisa estar na mesma rede do EasyPanel em que o host
  `agente_revalle_postgres` resolve (base de RH, somente leitura). Use um usuário de leitura
  em `COLABORADORES_DATABASE_URL`.
- **HTTPS**: a leitura de QR code com a câmera do celular (inventário) exige HTTPS. O proxy com TLS
  já atende; confira antes de usar no celular.
- **E-mail**: SMTP `mail.revalle.com.br:465`, com a senha (`SMTP_PASS`) só no `.env` do servidor.
- **Fuso horário**: o `docker-compose.yml` define `TZ=America/Sao_Paulo` no serviço do backend.
  A partir deste deploy, os registros gravados com o relógio do servidor (por exemplo
  `history.data_operacao`, escrito com `datetime.now()`) passam a ficar em horário de Brasília;
  antes, dentro do container, eram gravados em UTC. Os termos de responsabilidade já gravavam em
  horário de Brasília. Registros antigos não são alterados: os anteriores ao deploy continuam como
  foram gravados (UTC). **Se o backend for publicado fora deste compose (por exemplo, direto no
  EasyPanel), defina `TZ=America/Sao_Paulo` nas variáveis de ambiente do serviço.** Sem isso,
  a linha do tempo pode mostrar eventos de termo com 3 horas de diferença em relação aos demais.

---

## Migrações

As migrações do banco rodam sozinhas na subida do backend (com trava, para os 2 workers não
migrarem ao mesmo tempo) e também por linha de comando, de dentro do container:

```bash
docker compose exec backend python -m scripts.migrate pendentes   # lista o que falta aplicar
docker compose exec backend python -m scripts.migrate             # aplica as pendentes
docker compose exec backend python -m scripts.migrate reverter    # reverte a última aplicada
```

As convenções (nomes, faixas por fase, pares `.sql`/`.down.sql`) estão em
`backend/migrations/README.md`.

---

## 2FA: Gestor sem acesso

Se um Gestor perdeu o autenticador e os códigos de recuperação, outro administrador com acesso
ao servidor zera o 2FA dele. O comando grava a ação na auditoria:

```bash
docker compose exec backend python -m scripts.reset_2fa <usuario>
```

No próximo login o Gestor cadastra um autenticador novo.

---

## Antes de publicar (checklist)

- [ ] **1. Backup completo do MySQL**: `mysqldump --single-transaction --routines <banco> > backup.sql`.
- [ ] **2. Ensaio em uma cópia**: restaure o backup num banco separado, rode `python -m scripts.migrate`
  apontando para a cópia e confira o sistema nela antes de tocar na produção.
- [ ] **3. Conferir o ENUM de status**: em produção, `SHOW CREATE TABLE items` deve mostrar a coluna
  `status` como `enum('Disponível','Indisponível','Pendente','Pendente Devolução') DEFAULT 'Disponível'`.
  A migração 0402 redefine a coluna com essa base mais `'Em Manutenção'`; se a definição for diferente
  (por exemplo `NOT NULL` ou outro padrão), ajuste a 0402 antes do deploy.
  Confira também collation e engine das tabelas:
  `SELECT @@collation_database, table_name, table_collation, engine FROM information_schema.tables WHERE table_schema = DATABASE();`
  — `engine` precisa ser `InnoDB` em `usuarios`, `items`, `history` e `unidades` (as tabelas novas têm
  chave estrangeira para elas; com MyISAM a migração falha). As tabelas novas são criadas com o padrão
  do banco (`@@collation_database`), como as antigas.
- [ ] **4. Janela de manutenção**: as migrações 0004 e 0401 adicionam colunas com chave estrangeira e
  reconstroem a tabela `items`, o que bloqueia escritas durante a cópia. Agende o horário.
- [ ] **5. Preencher o `.env`** com as variáveis novas (tabela acima) e conferir `TZ` se não for usar este compose.
  Confira com atenção o `APP_PUBLIC_URL`: é a base do QR code das etiquetas (`APP_PUBLIC_URL/e/{id}`).
  Antes de imprimir etiquetas em lote, imprima uma, leia o QR com o celular e confirme que abre o equipamento.
- [ ] **6. Confirmar a rede do EasyPanel** (host `agente_revalle_postgres`) e o HTTPS.
- [ ] **7. Depois da subida (OBRIGATÓRIO)**: rode `python -m scripts.migrate pendentes` e confirme que
  responde "nenhuma". A API sobe mesmo se uma migração falhar (o erro só vai para o log), então a
  tela abrir não prova que o banco está atualizado. Se houver pendente, leia o log, corrija e rode
  `python -m scripts.migrate` antes de liberar o uso. Depois sincronize os colaboradores pela tela e
  confira o painel inicial.

---

## Roles e Permissões

| Role | Acesso |
|------|--------|
| **Gestor** | Tudo o que o Técnico faz, mais: Baixa de ativos, Estorno de lançamentos, Usuários, Unidades, Auditoria, Sistemas e Pacotes por função (acessos), sincronização manual de colaboradores |
| **Técnico** | Painel (indicadores e alertas), Estoque, Cadastrar/Editar, Periféricos, Vincular, Empréstimos, Devoluções, Termos, Histórico (e a aba Histórico do equipamento), Relatórios, Gráficos, Manutenções, Colaboradores, Locais, Etiquetas, Inventários, Linhas (com edição e relatório), Eventos de acesso |
| **Jovem Aprendiz** | Somente leitura: indicadores do Painel, Estoque (e o detalhe do equipamento, sem a aba Histórico), Gráficos e Linhas. Na lista de colaboradores recebe o CPF mascarado (`***.456.789-**`) e sem e-mail/telefone; não abre o cadastro de colaborador |

---

## Endpoints da API

Documentação interativa (Swagger): `http://localhost:8000/api/docs`

| Grupo | Base |
|-------|------|
| Autenticação | `/api/auth/` |
| Itens/Equipamentos | `/api/items/` |
| Periféricos | `/api/peripherals/` |
| Empréstimos | `/api/loans/` |
| Documentos | `/api/documents/` |
| Histórico | `/api/history/` |
| Relatórios | `/api/reports/` |
| Usuários | `/api/users/` |
| Constantes | `/api/constants/` |
| Painel: indicadores | `/api/painel/resumo` (todos os papéis) |
| Painel: alertas | `/api/painel/alertas` (Gestor e Técnico) |
| Exportações Excel/PDF | `/api/exportacoes/*` (`estoque`, `historico`, `manutencoes`, `linhas`, `termos-pendentes`; `relatorio-mensal` só Gestor) |

---

## Segurança

- **Nunca** commite o arquivo `backend/.env` — ele contém credenciais reais.
  Use `backend/.env.example` como modelo (já está no `.gitignore`).
- Este repositório **não publica usuário ou senha padrão**. O primeiro
  usuário Gestor deve ser criado diretamente no banco (veja "Schema herdado"
  abaixo), com hash bcrypt de uma senha própria — nunca em texto plano.
- Se alguma credencial (senha do banco, chave JWT, chave AWS) já foi
  commitada no histórico do git em algum momento — mesmo que depois removida
  do código —, trate-a como comprometida e **rotacione-a no servidor**.
  Trocar o valor no código não desfaz a exposição anterior. Veja
  `docs/REMOCAO_HISTORICO_GIT.md`.

---

## Schema herdado

A tabela `usuarios` é criada/reaproveitada pelo backend com o mesmo schema do
sistema desktop original, para não exigir migração do banco existente:

```sql
CREATE TABLE IF NOT EXISTS usuarios (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    role ENUM('Gestor', 'Técnico', 'Jovem Aprendiz') NOT NULL
)
```

`password` guarda sempre o hash bcrypt, nunca a senha em texto plano.
