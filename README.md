# DocsPoint

Guia passo a passo para instalar e executar o **DocsPoint** em um servidor local utilizando Docker e Docker Compose.

---

## 1. Pré-requisitos

Antes de começar, certifique-se de ter instalado no servidor/máquina local:

- **Git**
- **Docker** (Docker Engine ou Docker Desktop)
- **Docker Compose v2** (`docker compose`)
- Portas livres no servidor: `3000` (Frontend), `5000` (Backend API), `5432` (PostgreSQL), `6379` (Redis), `7700` (Meilisearch), `9000` e `9001` (MinIO S3)

---

## 2. Clonar o Repositório

```bash
git clone https://github.com/Forc1pedro/DocsPoint.git
cd DocsPoint
```

---

## 3. Configurar Variáveis de Ambiente (Opcional)

O arquivo `docker-compose.yml` já possui valores padrão prontos para rodar localmente. Caso queira personalizar portas, senhas ou chaves de produção, crie um arquivo `.env` na raiz do projeto antes de subir os containers:

```env
# PostgreSQL
POSTGRES_USER=sharepoint
POSTGRES_PASSWORD=sua_senha_segura_postgres
POSTGRES_DB=sharepoint_db
POSTGRES_PORT=5432

# MinIO (S3 Object Storage)
MINIO_ROOT_USER=minioadmin
MINIO_ROOT_PASSWORD=sua_senha_segura_minio
MINIO_API_PORT=9000
MINIO_CONSOLE_PORT=9001
S3_BUCKET=sp-documents

# Redis & Meilisearch
REDIS_PORT=6379
MEILISEARCH_PORT=7700
MEILI_MASTER_KEY=sua_chave_meilisearch

# Aplicação & Autenticação
BACKEND_PORT=5000
FRONTEND_PORT=3000
JWT_SECRET=substitua-por-uma-chave-secreta-forte
```

---

## 4. Construir e Subir os Containers

Na raiz do projeto, execute:

```bash
docker compose up -d --build
```

Na primeira execução, o Docker fará o build do frontend e do backend e inicializara automaticamente:
- **PostgreSQL 16** (com criação automática de schemas e dados iniciais)
- **MinIO S3** (bucket de documentos)
- **Redis 7** (cache)
- **Meilisearch** (motor de busca)
- **Backend API** (`:5000`)
- **Frontend Web** (`:3000`)

---

## 5. Verificar se Todos os Serviços Estão Saudáveis

Execute o comando abaixo para conferir o status dos 6 containers:

```bash
docker compose ps
```

Todos os serviços devem aparecer com status `Up` / `(healthy)`. Você também pode validar o healthcheck da API:

```bash
curl http://localhost:5000/api/v1/health
```

---

## 6. Acessar a Aplicação

Abra no navegador:

- **Portal DocsPoint:** `http://localhost:3000`
- **API Backend (Health):** `http://localhost:5000/api/v1/health`
- **Console MinIO (S3):** `http://localhost:9001`

### Acesso Inicial no Portal (`/login`)
Ao acessar `http://localhost:3000`, você será redirecionado automaticamente para a tela de login (`/login`).

- **Administrador:** `admin@sharepoint.local`
- **Membro:** `carlos.silva@empresa.com.br`
- **Visitante:** `visitante@parceiro.com.br`
- **Senha inicial:** `SharePoint@2026`

> **Dica para servidor na rede local (LAN):** caso acesse de outra máquina na rede usando o IP do servidor, substitua `localhost` pelo IP do servidor no navegador (ex: `http://192.168.1.50:3000`).

---

## 7. Comandos Úteis de Manutenção

### Parar os containers (mantendo todos os dados salvos)
```bash
docker compose down
```

### Iniciar novamente após reiniciar o servidor
```bash
docker compose up -d
```

### Atualizar após baixar novidades do Git
```bash
git pull origin main
docker compose up -d --build
```

### Visualizar logs em tempo real
```bash
docker compose logs -f backend frontend
```

### Reset completo (apaga banco de dados e arquivos salvos nos volumes)
```bash
docker compose down -v
docker compose up -d --build
```
