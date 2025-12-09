import { Router, Request, Response } from 'express';
import * as XLSX from 'xlsx';
import { pool } from '../config/database';

const router = Router();

// Exportar dados para Excel - Relatório Comparativo
router.get('/excel', async (req: Request, res: Response) => {
  try {
    const { mesReferencia } = req.query;
    
    console.log('📊 Gerando planilha Excel do Relatório Comparativo...');
    
    // Query 1: Relatório Comparativo Detalhado (UNIÃO de funcionários E empenhos)
    const queryRelatorio = `
      SELECT 
        COALESCE(f.comunidade, e.comunidade) as Comunidade,
        COALESCE(f.time_bre, e.equipe) as "Equipe (BRE)",
        COALESCE(f.gerente, 'N/A') as Gerente,
        COALESCE(f.preposto, 'N/A') as Preposto,
        COALESCE(COUNT(DISTINCT f.id), 0) as "Qtd Funcionários",
        ROUND(COALESCE(SUM(f.valor_proporcional), 0), 2) as "Valor Posto Proporcional",
        COALESCE(MAX(e.quantidade_membros), 0) as "Qtd Empenho",
        ROUND(COALESCE(MAX(e.valor_liquido), 0), 2) as "Valor Empenho",
        ROUND((COALESCE(MAX(e.valor_liquido), 0) - COALESCE(SUM(f.valor_proporcional), 0)), 2) as "Diferença Valor",
        CASE 
          WHEN COALESCE(MAX(e.valor_liquido), 0) - COALESCE(SUM(f.valor_proporcional), 0) > 0 
          THEN '💰 Sobra no Empenho'
          WHEN COALESCE(MAX(e.valor_liquido), 0) - COALESCE(SUM(f.valor_proporcional), 0) < 0 
          THEN '⚠️ Falta no Empenho'
          ELSE '✅ Valores Batem'
        END as Status
      FROM (
        SELECT DISTINCT time_bre as equipe FROM funcionarios ${mesReferencia ? 'WHERE mes_referencia = ?' : ''}
        UNION
        SELECT DISTINCT equipe FROM empenhos ${mesReferencia ? 'WHERE mes_referencia = ?' : ''}
      ) todas_equipes
      LEFT JOIN funcionarios f ON f.time_bre = todas_equipes.equipe ${mesReferencia ? 'AND f.mes_referencia = ?' : ''}
      LEFT JOIN empenhos e ON e.equipe = todas_equipes.equipe ${mesReferencia ? 'AND e.mes_referencia = ?' : ''}
      GROUP BY todas_equipes.equipe, f.comunidade, e.comunidade, f.gerente, f.preposto
      ORDER BY Comunidade, "Equipe (BRE)"
    `;
    
    const params = mesReferencia ? [mesReferencia, mesReferencia, mesReferencia, mesReferencia] : [];
    const relatorio = await pool.prepare(queryRelatorio).all(...params);
    
    // Query 2: Totais
    const queryTotais = `
      SELECT 
        'TOTAL GERAL' as Descrição,
        COUNT(DISTINCT f.id) as "Total Funcionários",
        ROUND(SUM(f.valor_proporcional), 2) as "Total Valor Funcionários",
        (SELECT ROUND(SUM(quantidade_membros), 2) FROM empenhos ${mesReferencia ? 'WHERE mes_referencia = ?' : ''}) as "Total Qtd Empenho",
        (SELECT ROUND(SUM(valor_liquido), 2) FROM empenhos ${mesReferencia ? 'WHERE mes_referencia = ?' : ''}) as "Total Valor Empenho"
      FROM funcionarios f
      ${mesReferencia ? 'WHERE f.mes_referencia = ?' : ''}
    `;
    
    const totaisParams = mesReferencia ? [mesReferencia, mesReferencia, mesReferencia] : [];
    const totais = await pool.prepare(queryTotais).all(...totaisParams);
    
    // Query 3: Lista completa de funcionários
    const queryFuncionarios = `
      SELECT 
        comunidade as Comunidade,
        time_bre as "Equipe (BRE)",
        nome as Nome,
        posto as Posto,
        ROUND(valor_proporcional, 2) as "Valor Proporcional",
        gerente as Gerente,
        preposto as Preposto
      FROM funcionarios
      ${mesReferencia ? 'WHERE mes_referencia = ?' : ''}
      ORDER BY comunidade, time_bre, nome
    `;
    
    const funcionarios = mesReferencia 
      ? await pool.prepare(queryFuncionarios).all(mesReferencia)
      : await pool.prepare(queryFuncionarios).all();
    
    // Query 4: Lista de membros do empenho (sem duplicados)
    const queryMembros = `
      SELECT 
        MAX(e.comunidade) as Comunidade,
        m.equipe as Equipe,
        m.nome as "Nome do Membro",
        m.matricula as "Matrícula do Membro"
      FROM membros_empenho m
      LEFT JOIN empenhos e ON m.equipe = e.equipe AND m.mes_referencia = e.mes_referencia
      ${mesReferencia ? 'WHERE m.mes_referencia = ? AND m.nome IS NOT NULL AND m.nome != "" AND m.nome NOT LIKE "%Equipe%" AND m.nome NOT LIKE "%Nome%"' : 'WHERE m.nome IS NOT NULL AND m.nome != "" AND m.nome NOT LIKE "%Equipe%" AND m.nome NOT LIKE "%Nome%"'}
      GROUP BY m.equipe, m.nome, m.matricula
      ORDER BY Comunidade, m.equipe, m.nome
    `;
    
    const membros = mesReferencia 
      ? await pool.prepare(queryMembros).all(mesReferencia)
      : await pool.prepare(queryMembros).all();
    
    // Query 5: Funcionários que estão faltando no empenho (não estão na lista de membros) - sem duplicados
    const queryFaltandoEmpenho = `
      SELECT 
        MAX(f.comunidade) as Comunidade,
        f.time_bre as "Equipe (BRE)",
        f.nome as "Nome do Funcionário",
        f.matricula as Matrícula,
        MAX(f.posto) as Posto,
        ROUND(MAX(f.valor_proporcional), 2) as "Valor Proporcional",
        MAX(f.gerente) as Gerente,
        MAX(f.preposto) as Preposto
      FROM funcionarios f
      WHERE ${mesReferencia ? 'f.mes_referencia = ? AND' : ''} f.time_bre IS NOT NULL
        AND NOT EXISTS (
          SELECT 1 FROM membros_empenho m 
          WHERE UPPER(TRIM(m.matricula)) = UPPER(TRIM(f.matricula))
            ${mesReferencia ? 'AND m.mes_referencia = ?' : ''}
        )
      GROUP BY f.matricula, f.time_bre, f.nome
      ORDER BY Comunidade, "Equipe (BRE)", f.nome
    `;
    
    const faltandoParams = mesReferencia ? [mesReferencia, mesReferencia] : [];
    const faltandoEmpenho = await pool.prepare(queryFaltandoEmpenho).all(...faltandoParams);
    
    // Criar workbook
    const workbook = XLSX.utils.book_new();
    
    // Aba 1: Relatório Comparativo (principal)
    const wsRelatorio = XLSX.utils.json_to_sheet(relatorio);
    wsRelatorio['!cols'] = [
      { wch: 35 }, // Comunidade
      { wch: 15 }, // Equipe
      { wch: 15 }, // Gerente
      { wch: 15 }, // Preposto
      { wch: 18 }, // Qtd Funcionários
      { wch: 22 }, // Valor Posto
      { wch: 15 }, // Qtd Empenho
      { wch: 18 }, // Valor Empenho
      { wch: 18 }, // Diferença Valor
      { wch: 22 }  // Status
    ];
    
    // Aplicar formato de moeda nas colunas de valores (Relatório)
    const range = XLSX.utils.decode_range(wsRelatorio['!ref'] || 'A1');
    for (let row = 1; row <= range.e.r; row++) {
      // Coluna F (índice 5): Valor Posto Proporcional
      const cellF = XLSX.utils.encode_cell({ r: row, c: 5 });
      if (wsRelatorio[cellF] && typeof wsRelatorio[cellF].v === 'number') {
        wsRelatorio[cellF].z = 'R$ #,##0.00';
      }
      // Coluna H (índice 7): Valor Empenho
      const cellH = XLSX.utils.encode_cell({ r: row, c: 7 });
      if (wsRelatorio[cellH] && typeof wsRelatorio[cellH].v === 'number') {
        wsRelatorio[cellH].z = 'R$ #,##0.00';
      }
      // Coluna I (índice 8): Diferença Valor
      const cellI = XLSX.utils.encode_cell({ r: row, c: 8 });
      if (wsRelatorio[cellI] && typeof wsRelatorio[cellI].v === 'number') {
        wsRelatorio[cellI].z = 'R$ #,##0.00';
      }
    }
    
    XLSX.utils.book_append_sheet(workbook, wsRelatorio, 'Relatório Comparativo');
    
    // Aba 2: Totais
    const wsTotais = XLSX.utils.json_to_sheet(totais);
    wsTotais['!cols'] = [
      { wch: 20 }, { wch: 20 }, { wch: 25 }, { wch: 20 }, { wch: 22 }
    ];
    
    // Aplicar formato de moeda nas colunas de valores (Totais)
    const rangeTotais = XLSX.utils.decode_range(wsTotais['!ref'] || 'A1');
    for (let row = 1; row <= rangeTotais.e.r; row++) {
      // Colunas C, E (valores)
      for (let col of [2, 4]) {
        const cell = XLSX.utils.encode_cell({ r: row, c: col });
        if (wsTotais[cell] && typeof wsTotais[cell].v === 'number') {
          wsTotais[cell].z = 'R$ #,##0.00';
        }
      }
    }
    
    XLSX.utils.book_append_sheet(workbook, wsTotais, 'Totais');
    
    // Aba 3: Lista de Funcionários
    const wsFuncionarios = XLSX.utils.json_to_sheet(funcionarios);
    wsFuncionarios['!cols'] = [
      { wch: 35 }, // Comunidade
      { wch: 15 }, // Equipe
      { wch: 35 }, // Nome
      { wch: 25 }, // Posto
      { wch: 18 }, // Valor
      { wch: 15 }, // Gerente
      { wch: 15 }  // Preposto
    ];
    
    // Aplicar formato de moeda na coluna Valor Proporcional (Funcionários)
    const rangeFuncionarios = XLSX.utils.decode_range(wsFuncionarios['!ref'] || 'A1');
    for (let row = 1; row <= rangeFuncionarios.e.r; row++) {
      // Coluna E (índice 4): Valor Proporcional
      const cell = XLSX.utils.encode_cell({ r: row, c: 4 });
      if (wsFuncionarios[cell] && typeof wsFuncionarios[cell].v === 'number') {
        wsFuncionarios[cell].z = 'R$ #,##0.00';
      }
    }
    
    XLSX.utils.book_append_sheet(workbook, wsFuncionarios, 'Funcionários');
    
    // Aba 4: Membros do Empenho
    const wsMembros = XLSX.utils.json_to_sheet(membros);
    wsMembros['!cols'] = [
      { wch: 35 }, // Comunidade
      { wch: 15 }, // Equipe
      { wch: 35 }, // Nome
      { wch: 12 }  // Matrícula
    ];
    XLSX.utils.book_append_sheet(workbook, wsMembros, 'Membros Empenho');
    
    // Aba 5: Funcionários Faltando no Empenho
    const wsFaltando = XLSX.utils.json_to_sheet(faltandoEmpenho);
    wsFaltando['!cols'] = [
      { wch: 35 }, // Comunidade
      { wch: 15 }, // Equipe
      { wch: 35 }, // Nome
      { wch: 25 }, // Posto
      { wch: 18 }, // Valor
      { wch: 15 }, // Gerente
      { wch: 15 }  // Preposto
    ];
    
    // Aplicar formato de moeda
    const rangeFaltando = XLSX.utils.decode_range(wsFaltando['!ref'] || 'A1');
    for (let row = 1; row <= rangeFaltando.e.r; row++) {
      const cell = XLSX.utils.encode_cell({ r: row, c: 4 });
      if (wsFaltando[cell] && typeof wsFaltando[cell].v === 'number') {
        wsFaltando[cell].z = 'R$ #,##0.00';
      }
    }
    
    XLSX.utils.book_append_sheet(workbook, wsFaltando, 'Faltando no Empenho');
    
    // Gerar buffer
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    
    // Definir nome do arquivo
    const nomeArquivo = mesReferencia 
      ? `relatorio_comparativo_${mesReferencia}.xlsx`
      : `relatorio_comparativo_completo.xlsx`;
    
    console.log(`✅ Planilha gerada: ${nomeArquivo}`);
    console.log(`📊 ${relatorio.length} equipes no relatório`);
    console.log(`👥 ${funcionarios.length} funcionários listados`);
    console.log(`👤 ${membros.length} membros do empenho listados`);
    console.log(`⚠️ ${faltandoEmpenho.length} funcionários faltando no empenho`);
    
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
