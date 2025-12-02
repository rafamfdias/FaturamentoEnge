import { Router, Request, Response } from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import { processarPlanilhaEmpenho, listarEmpenhos, obterTotalEmpenho } from '../services/empenhoService';
import { obterTotalValorProporcional } from '../services/planilhaService';

const router = Router();

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, '../../uploads');
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, 'empenho-' + uniqueSuffix + path.extname(file.originalname));
  }
});

const upload = multer({
  storage: storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['.xlsx', '.xls', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Apenas arquivos .xlsx, .xls e .csv são permitidos'));
    }
  }
});

// Upload de planilha de empenho
router.post('/upload', upload.single('planilha'), async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Nenhum arquivo foi enviado' });
    }

    const resultado = await processarPlanilhaEmpenho(req.file.path);

    // Remover arquivo após processamento
    fs.unlinkSync(req.file.path);

    res.json({
      message: 'Planilha de empenho processada com sucesso',
      ...resultado
    });
  } catch (error: any) {
    console.error('Erro no upload de empenho:', error);
    res.status(500).json({ error: error.message });
  }
});

// Listar todos os empenhos
router.get('/', async (req: Request, res: Response) => {
  try {
    const empenhos = await listarEmpenhos();
    res.json(empenhos);
  } catch (error: any) {
    console.error('Erro ao listar empenhos:', error);
    res.status(500).json({ error: error.message });
  }
});

// Obter relatório comparativo
router.get('/relatorio', async (req: Request, res: Response) => {
  try {
    const totalFuncionarios = await obterTotalValorProporcional();
    const totalEmpenhos = await obterTotalEmpenho();
    const diferenca = totalEmpenhos - totalFuncionarios;

    res.json({
      totalFuncionarios,
      totalEmpenhos,
      diferenca
    });
  } catch (error: any) {
    console.error('Erro ao gerar relatório:', error);
    res.status(500).json({ error: error.message });
  }
});

export default router;
