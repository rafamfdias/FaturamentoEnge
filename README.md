# Sistema de Faturamento - Upload de Planilhas

Sistema web para processamento de planilhas Excel com informações de funcionários e contratos.

## 🚀 Tecnologias

- **Backend:** Node.js com TypeScript
- **Framework:** Express.js
- **Banco de Dados:** SQLite
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
- Os dados são armazenados em SQLite
- A aplicação detecta automaticamente variações nos nomes das colunas (maiúsculas/minúsculas)
