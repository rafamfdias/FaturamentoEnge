# Usar imagem Node.js LTS
FROM node:18-alpine

# Definir diretório de trabalho
WORKDIR /app

# Copiar package.json e package-lock.json
COPY package*.json ./

# Instalar dependências
RUN npm install

# Copiar o código da aplicação
COPY . .

# Compilar TypeScript
RUN npm run build

# Criar diretório para uploads e banco de dados
RUN mkdir -p uploads
RUN mkdir -p data

# Expor a porta da aplicação
EXPOSE 8888

# Comando para iniciar a aplicação
CMD ["node", "dist/server.js"]
