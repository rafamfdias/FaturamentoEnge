import { Router, Request, Response } from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import { processarPlanilhaEmpenho, listarEmpenhos, obterTotalEmpenho, listarMembrosEmpenho, analisarFuncionariosForaEmpenho, listarMesesDisponiveisEmpenho } from '../services/empenhoService';
import { PlanilhaService } from '../services/planilhaService';

const router = Router();
const planilhaService = new PlanilhaService();

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

    const mesReferencia = req.body.mesReferencia || null;
    const resultado = await processarPlanilhaEmpenho(req.file.path, mesReferencia, req.file.originalname);

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
    const { mesReferencia } = req.query;
    const empenhos = await listarEmpenhos(mesReferencia as string);
    res.json(empenhos);
  } catch (error: any) {
    console.error('Erro ao listar empenhos:', error);
    res.status(500).json({ error: error.message });
  }
});

// Obter relatório comparativo
router.get('/relatorio', async (req: Request, res: Response) => {
  try {
    const { mesReferencia } = req.query;
    const totalFuncionarios = await planilhaService.obterTotalValorProporcional(mesReferencia as string);
    const totalEmpenhos = await obterTotalEmpenho(mesReferencia as string);
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

// Listar meses disponíveis
router.get('/meses', async (req: Request, res: Response) => {
  try {
    const meses = await listarMesesDisponiveisEmpenho();
    res.json(meses);
  } catch (error: any) {
    console.error('Erro ao listar meses:', error);
    res.status(500).json({ error: error.message });
  }
});

// Listar membros do empenho por equipe
router.get('/membros/:equipe?', async (req: Request, res: Response) => {
  try {
    const { equipe } = req.params;
    const membros = await listarMembrosEmpenho(equipe);
    res.json(membros);
  } catch (error: any) {
    console.error('Erro ao listar membros do empenho:', error);
    res.status(500).json({ error: error.message });
  }
});

// Analisar funcionários que não estão no empenho
router.get('/analise/fora-empenho', async (req: Request, res: Response) => {
  try {
    const resultado = await analisarFuncionariosForaEmpenho();
    res.json(resultado);
  } catch (error: any) {
    console.error('Erro ao analisar funcionários fora do empenho:', error);
    res.status(500).json({ error: error.message });
  }
});

export default router;
