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

## Roles e Permissões

| Role | Acesso |
|------|--------|
| **Gestor** | Tudo, incluindo Remover itens, Usuários, Estorno |
| **Técnico** | Estoque, Cadastrar, Editar, Periféricos, Vincular, Emprestar, Devolver, Histórico, Relatório, Termos |
| **Jovem Aprendiz** | Estoque, Gráficos (somente leitura) |

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
