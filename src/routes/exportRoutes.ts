import { Router, Request, Response } from 'express';
import * as XLSX from 'xlsx';
import { pool } from '../config/database';

const router = Router();

// Exportar dados para Excel - Relatório Comparativo
router.get('/excel', async (req: Request, res: Response) => {
  try {
    const { mesReferencia } = req.query;
    
    console.log('📊 Gerando planilha Excel do Relatório Comparativo...');
    
    // Query principal: Relatório Comparativo Detalhado
    const queryRelatorio = `
      SELECT 
        f.contrato as Contrato,
        f.comunidade as Comunidade,
        f.gerente as Gerente,
        f.preposto as Preposto,
        f.time_bre as "Equipe (BRE)",
        COUNT(DISTINCT f.id) as "Qtd Funcionários",
        ROUND(SUM(f.valor_proporcional), 2) as "Valor Posto Proporcional",
        COALESCE(e.quantidade_membros, 0) as "Qtd Empenho",
        ROUND(COALESCE(e.valor_liquido, 0), 2) as "Valor Empenho",
        (COUNT(DISTINCT f.id) - COALESCE(e.quantidade_membros, 0)) as "Diferença Qtd",
        ROUND((COALESCE(e.valor_liquido, 0) - SUM(f.valor_proporcional)), 2) as "Diferença Valor",
        CASE 
          WHEN (COUNT(DISTINCT f.id) - COALESCE(e.quantidade_membros, 0)) = 0 
            AND ABS(COALESCE(e.valor_liquido, 0) - SUM(f.valor_proporcional)) < 0.01 
          THEN '✅ OK'
          WHEN (COUNT(DISTINCT f.id) - COALESCE(e.quantidade_membros, 0)) > 0 
          THEN '⚠️ Mais Funcionários'
          WHEN (COUNT(DISTINCT f.id) - COALESCE(e.quantidade_membros, 0)) < 0 
          THEN '⚠️ Menos Funcionários'
          ELSE '⚠️ Divergência'
        END as Status
      FROM funcionarios f
      LEFT JOIN empenhos e ON f.time_bre = e.equipe ${mesReferencia ? 'AND f.mes_referencia = e.mes_referencia' : ''}
      ${mesReferencia ? 'WHERE f.mes_referencia = ?' : ''}
      GROUP BY f.contrato, f.comunidade, f.gerente, f.preposto, f.time_bre, e.quantidade_membros, e.valor_liquido
      ORDER BY f.comunidade, f.time_bre
    `;
    
    const relatorio = mesReferencia 
      ? await pool.prepare(queryRelatorio).all(mesReferencia)
      : await pool.prepare(queryRelatorio).all();
    
    // Query para Totais
    const queryTotais = `
      SELECT 
        'TOTAL GERAL' as Descrição,
        COUNT(DISTINCT f.id) as "Total Funcionários",
        ROUND(SUM(f.valor_proporcional), 2) as "Total Valor Funcionários",
        SUM(COALESCE(e.quantidade_membros, 0)) as "Total Qtd Empenho",
        ROUND(SUM(COALESCE(e.valor_liquido, 0)), 2) as "Total Valor Empenho",
        ROUND((SUM(COALESCE(e.valor_liquido, 0)) - SUM(f.valor_proporcional)), 2) as "Diferença Total"
      FROM funcionarios f
      LEFT JOIN empenhos e ON f.time_bre = e.equipe ${mesReferencia ? 'AND f.mes_referencia = e.mes_referencia' : ''}
      ${mesReferencia ? 'WHERE f.mes_referencia = ?' : ''}
    `;
    
    const totais = mesReferencia 
      ? await pool.prepare(queryTotais).all(mesReferencia)
      : await pool.prepare(queryTotais).all();
    
    // Criar workbook
    const workbook = XLSX.utils.book_new();
    
    // Aba 1: Relatório Comparativo (principal)
    const wsRelatorio = XLSX.utils.json_to_sheet(relatorio);
    
    // Ajustar largura das colunas
    wsRelatorio['!cols'] = [
      { wch: 12 }, // Contrato
      { wch: 35 }, // Comunidade
      { wch: 15 }, // Gerente
      { wch: 15 }, // Preposto
      { wch: 15 }, // Equipe
      { wch: 18 }, // Qtd Funcionários
      { wch: 22 }, // Valor Posto
      { wch: 15 }, // Qtd Empenho
      { wch: 18 }, // Valor Empenho
      { wch: 15 }, // Diferença Qtd
      { wch: 18 }, // Diferença Valor
      { wch: 22 }  // Status
    ];
    
    XLSX.utils.book_append_sheet(workbook, wsRelatorio, 'Relatório Comparativo');
    
    // Aba 2: Totais
    const wsTotais = XLSX.utils.json_to_sheet(totais);
    XLSX.utils.book_append_sheet(workbook, wsTotais, 'Resumo Totais');
    
    // Gerar buffer
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    
    // Definir nome do arquivo
    const nomeArquivo = mesReferencia 
      ? `relatorio_comparativo_${mesReferencia}.xlsx`
      : `relatorio_comparativo_completo.xlsx`;
    
    console.log(`✅ Planilha gerada: ${nomeArquivo}`);
    console.log(`📊 ${relatorio.length} equipes no relatório`);
    
    // Enviar arquivo
    res.setHeader('Content-Disposition', `attachment; filename="${nomeArquivo}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
    
  } catch (error: any) {
    console.error('❌ Erro ao gerar Excel:', error);
    res.status(500).json({ error: error.message });
  }
});

export default router;
