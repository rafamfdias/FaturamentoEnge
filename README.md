# Sistema de Faturamento - Upload de Planilhas

Sistema web para processamento de planilhas Excel com informações de funcionários e contratos.

## 🚀 Tecnologias

- **Backend:** Node.js com TypeScript
- **Framework:** Express.js
- **Banco de Dados:** PostgreSQL
- **Upload:** Multer
- **Processamento Excel:** xlsx
- **Frontend:** HTML, CSS, JavaScript

## 📋 Estrutura da Planilha

A planilha deve conter as seguintes colunas:

- CONTRATO
- COMUNIDADE
- TIME(BRE)
- GERENTE
- PREPOSTO
- NOME (obrigatório)
- MATRICULA
- POSTO
- GRUPO
- VALOR PROPORCIONAL

## 🔧 Instalação

1. Instale as dependências:
```bash
npm install
```

2. Configure o banco de dados PostgreSQL:
   - Crie um banco de dados chamado `faturamento_db`
   - Ajuste as credenciais no arquivo `.env`

3. Copie o arquivo de exemplo de variáveis de ambiente:
```bash
cp .env.example .env
```

4. Edite o arquivo `.env` com suas configurações:
```env
PORT=3000
DB_HOST=localhost
DB_PORT=5432
DB_USER=postgres
DB_PASSWORD=sua_senha_aqui
DB_NAME=faturamento_db
```

## 🎯 Como Usar

### Desenvolvimento

```bash
npm run dev
```

O servidor estará disponível em `http://localhost:3000`

### Produção

```bash
npm run build
npm start
```

## 📡 API Endpoints

### Upload de Planilha
```
POST /api/upload
Content-Type: multipart/form-data
Body: planilha (arquivo Excel)
```

### Listar Funcionários
```
GET /api/funcionarios
```

### Limpar Dados
```
DELETE /api/funcionarios
```

## 💾 Banco de Dados

O sistema cria automaticamente a tabela necessária:

```sql
CREATE TABLE funcionarios (
  id SERIAL PRIMARY KEY,
  contrato VARCHAR(255),
  comunidade VARCHAR(255),
  time_bre VARCHAR(255),
  gerente VARCHAR(255),
  preposto VARCHAR(255),
  nome VARCHAR(255) NOT NULL,
  matricula VARCHAR(100),
  posto VARCHAR(255),
  grupo VARCHAR(255),
  valor_proporcional DECIMAL(10, 2),
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

## 🌟 Funcionalidades

- ✅ Upload de planilhas Excel (.xlsx, .xls) e CSV
- ✅ Processamento automático de dados
- ✅ Validação de campos obrigatórios
- ✅ Relatório de erros por linha
- ✅ Interface web intuitiva com drag & drop
- ✅ Visualização dos dados cadastrados
- ✅ Limpeza de dados
- ✅ Estatísticas em tempo real

## 📝 Observações

- O campo NOME é obrigatório
- Arquivos são limitados a 10MB
- Os dados são armazenados em PostgreSQL
- A aplicação detecta automaticamente variações nos nomes das colunas (maiúsculas/minúsculas)

## 🛠️ Próximos Passos

Você mencionou que há uma segunda planilha. Quando estiver pronto, me informe quais são os campos dessa segunda planilha para que eu possa adicionar ao sistema!
