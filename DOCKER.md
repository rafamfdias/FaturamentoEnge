# 🚀 Como usar o Docker

## Opção 1: Usando Docker Compose (Recomendado)

```bash
# Construir e iniciar o container
docker-compose up -d

# Ver logs
docker-compose logs -f

# Parar o container
docker-compose down
```

## Opção 2: Usando Docker direto

```bash
# Construir a imagem
docker build -t faturamento-enge .

# Executar o container
docker run -d \
  --name faturamento-enge \
  -p 3000:3000 \
  -v $(pwd)/data:/app/data \
  -v $(pwd)/uploads:/app/uploads \
  faturamento-enge

# Ver logs
docker logs -f faturamento-enge

# Parar o container
docker stop faturamento-enge

# Remover o container
docker rm faturamento-enge
```

## 📍 Acessar a aplicação

Após iniciar o container, acesse: http://localhost:3000

## 💾 Persistência de dados

Os dados ficam salvos nas pastas:
- `./data` - Banco de dados SQLite
- `./uploads` - Arquivos enviados

Essas pastas são montadas como volumes, então os dados **não são perdidos** quando você para o container.

## 🔄 Atualizar a aplicação

```bash
# Parar o container
docker-compose down

# Atualizar o código (git pull, etc)

# Reconstruir e iniciar
docker-compose up -d --build
```
