import { Router, Request, Response } from 'express';
import { upload } from '../config/upload';
import { PlanilhaService } from '../services/planilhaService';
import fs from 'fs';

const router = Router();
const planilhaService = new PlanilhaService();

// Upload e processar planilha
router.post('/upload', upload.single('planilha'), async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Nenhum arquivo enviado' });
    }

    const mesReferencia = req.body.mesReferencia || null;
    const resultado = await planilhaService.processarPlanilha(req.file.path, mesReferencia, req.file.originalname);

    // Remover arquivo após processamento
    fs.unlinkSync(req.file.path);

    res.json({
      message: 'Planilha processada com sucesso',
      registros_importados: resultado.sucesso,
      erros: resultado.erros.length > 0 ? resultado.erros : undefined
    });
  } catch (error: any) {
    // Remover arquivo em caso de erro
    if (req.file) {
      fs.unlinkSync(req.file.path);
    }
    
    res.status(500).json({ error: error.message });
  }
});

// Listar todos os funcionários
router.get('/funcionarios', async (req: Request, res: Response) => {
  try {
    const { mesReferencia } = req.query;
    const funcionarios = await planilhaService.listarFuncionarios(mesReferencia as string);
    res.json(funcionarios);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

// Obter total do valor proporcional
router.get('/total', async (req: Request, res: Response) => {
  try {
    const { mesReferencia } = req.query;
    const total = await planilhaService.obterTotalValorProporcional(mesReferencia as string);
    res.json({ total });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

// Limpar todos os dados
router.delete('/funcionarios', async (req: Request, res: Response) => {
  try {
    await planilhaService.limparDados();
    res.json({ message: 'Todos os dados foram removidos' });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

// Listar meses disponíveis
router.get('/meses', async (req: Request, res: Response) => {
  try {
    const meses = await planilhaService.listarMesesDisponiveis();
    res.json(meses);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

export default router;
