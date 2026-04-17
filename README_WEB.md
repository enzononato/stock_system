# Controle de Estoque Revalle — Versão Web

Aplicação web completa com **FastAPI** (backend) + **React** (frontend), substituindo o sistema Tkinter original. O banco MySQL existente é reutilizado sem alterações de schema.

---

## Estrutura

```
/
├── backend/          ← API FastAPI (Python)
├── frontend/         ← SPA React (TypeScript + Vite + Shadcn/ui)
├── docker-compose.yml
└── README_WEB.md
```

---

## Como Rodar (Desenvolvimento Local)

### Backend

```bash
cd backend

# Crie o .env a partir do exemplo
cp .env.example .env
# Edite o .env com seus valores

# Instale as dependências (recomenda-se virtualenv)
pip install -r requirements.txt

# Copie os modelos de documento
# (os arquivos .docx já estão em ../modelos/ — crie um link ou copie)
cp -r ../modelos ./modelos

# Inicie o servidor
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

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

Preencha obrigatoriamente:
```env
JWT_SECRET=<string-aleatória-longa-32+-caracteres>
DB_HOST=72.61.53.20
DB_PASSWORD=<senha-real>

# Para produção, use S3
STORAGE_BACKEND=s3
S3_BUCKET=revalle-stock-docs
AWS_ACCESS_KEY_ID=<sua-key>
AWS_SECRET_ACCESS_KEY=<seu-secret>

# Domínio do frontend
CORS_ORIGINS=https://estoque.revalle.com.br
```

**3. Copie os modelos de documento**
```bash
cp -r modelos backend/modelos
```

**4. Suba os containers**
```bash
docker compose up -d --build
```

**5. Configure HTTPS com Caddy (recomendado — TLS automático)**

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

## Usuário padrão

```
Usuário: mãe
Senha:   Revalle@123
Role:    Gestor
```

> **Importante:** Altere a senha após o primeiro login!
