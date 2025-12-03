import express from 'express';
import cors from 'cors';
import path from 'path';
import dotenv from 'dotenv';
import { createTables } from './database/init';
import planilhaRoutes from './routes/planilhaRoutes';
import empenhoRoutes from './routes/empenhoRoutes';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 8888;

// Middlewares
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Servir arquivos estáticos
app.use(express.static(path.join(__dirname, '../public')));

// Rotas da API
app.use('/api', planilhaRoutes);
app.use('/api/empenho', empenhoRoutes);

// Rota principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '../public/index.html'));
});

// Inicializar banco de dados e servidor
const startServer = async () => {
  try {
    await createTables();
    
    app.listen(PORT, () => {
      console.log(`🚀 Servidor rodando na porta ${PORT}`);
      console.log(`📊 Acesse: http://localhost:${PORT}`);
    });
  } catch (error) {
    console.error('❌ Erro ao iniciar servidor:', error);
    process.exit(1);
  }
};

startServer();
