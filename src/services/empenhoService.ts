import * as XLSX from 'xlsx';
import path from 'path';
import { pool } from '../config/database';
import { extrairMesReferencia } from '../utils/mesReferencia';

interface EmpenhoData {
  comunidade?: string;
  equipe?: string;
  quantidade_membros?: number;
  valor_liquido?: number;
}

export const processarPlanilhaEmpenho = async (filePath: string, mesReferencia: string | null, nomeOriginal?: string) => {
  console.log('🔄 Iniciando processamento da planilha de empenho...');
  console.log('📁 Arquivo:', filePath);
  
  const nomeArquivo = nomeOriginal || path.basename(filePath);
  const mesRef = mesReferencia || extrairMesReferencia(nomeArquivo);
  console.log(`📅 Mês detectado de "${nomeArquivo}": ${mesRef}`);
  const mesReferenciaFinal = mesRef;
  
  try {
    const workbook = XLSX.readFile(filePath);
    console.log('📚 Planilha carregada com sucesso');
    
    // Procurar pela aba "Empenho"
    let sheetName = workbook.SheetNames.find(name => 
      name.toLowerCase() === 'empenho'
    );
    
    // Se não encontrar, tentar variações
    if (!sheetName) {
      sheetName = workbook.SheetNames.find(name => 
        name.toLowerCase().includes('empenho')
      );
    }
    
    // Se ainda não encontrar, usar a primeira aba e avisar
    if (!sheetName) {
      sheetName = workbook.SheetNames[0];
      console.log(`⚠️ Aba "Empenho" não encontrada. Usando: ${sheetName}`);
      console.log(`📋 Abas disponíveis: ${workbook.SheetNames.join(', ')}`);
    } else {
      console.log(`✅ Usando aba: ${sheetName}`);
    }
    
    const worksheet = workbook.Sheets[sheetName];
    
    // Ler todos os dados da planilha incluindo células vazias
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    console.log(`📊 Range da planilha: ${worksheet['!ref']}`);
    
    let registrosImportados = 0;
    const erros: string[] = [];

    // Limpar dados anteriores do mesmo mês
    const deleteEmpenhos = pool.prepare('DELETE FROM empenhos WHERE mes_referencia = ?');
    const deleteMembros = pool.prepare('DELETE FROM membros_empenho WHERE mes_referencia = ?');
    await deleteEmpenhos.run(mesReferenciaFinal);
    await deleteMembros.run(mesReferenciaFinal);
    console.log(`🗑️ Dados anteriores do mês ${mesReferenciaFinal} removidos`);
    
    let comunidadeAtual = '';
    let dentroDeTabelaEquipe = false;
    
    // Map para agrupar dados por comunidade + equipe
    const equipesMap = new Map<string, EmpenhoData>();
    
    // Percorrer todas as linhas da planilha
    for (let row = range.s.r; row <= range.e.r; row++) {
      try {
        // Ler células da linha atual
        const cellA = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })]; // Coluna A
        const cellB = worksheet[XLSX.utils.encode_cell({ r: row, c: 1 })]; // Coluna B
        
        // Verificar se há a palavra "Comunidade" em qualquer coluna dessa linha
        for (let col = 0; col <= range.e.c; col++) {
          const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
          if (cell && String(cell.v).toLowerCase().includes('comunidade')) {
            // Procurar o nome da comunidade na próxima célula com conteúdo
            for (let nextCol = col + 1; nextCol <= range.e.c; nextCol++) {
              const nextCell = worksheet[XLSX.utils.encode_cell({ r: row, c: nextCol })];
              if (nextCell && nextCell.v && String(nextCell.v).trim() !== '') {
                comunidadeAtual = String(nextCell.v).trim();
                console.log(`📍 Comunidade detectada na linha ${row + 1}, coluna ${String.fromCharCode(65 + nextCol)}: ${comunidadeAtual}`);
                dentroDeTabelaEquipe = false;
                break;
              }
            }
            break;
          }
        }
        
        // Verificar se é o cabeçalho "Termos de Equipe"
        if (cellA && String(cellA.v).toLowerCase().includes('termos de equipe')) {
          dentroDeTabelaEquipe = true;
          console.log(`📋 Iniciando seção de equipes na linha ${row + 1}`);
          continue;
        }
        
        // Verificar se é o cabeçalho da tabela (linha com "Equipe")
        if (cellA && String(cellA.v).toLowerCase().trim() === 'equipe') {
          console.log(`📊 Cabeçalho de tabela encontrado na linha ${row + 1}`);
          continue;
        }
        
        // Verificar se é uma linha de dados de equipe (começa com "BRE")
        if (dentroDeTabelaEquipe && cellA && String(cellA.v).startsWith('BRE')) {
          const equipe = String(cellA.v).trim();
          
          // Ler Quantidade Efetiva de Membros da coluna B
          let quantidadeMembros = 0;
          if (cellB && cellB.v) {
            const qtdStr = String(cellB.v);
            quantidadeMembros = parseFloat(qtdStr.replace(',', '.')) || 0;
          }
          
          // Ler Valor Líquido da coluna J (índice 9)
          const cellJ = worksheet[XLSX.utils.encode_cell({ r: row, c: 9 })];
          let valorLiquido = 0;
          if (cellJ && cellJ.v) {
            // Se for número direto, usar o valor
            if (typeof cellJ.v === 'number') {
              valorLiquido = cellJ.v;
            } else {
              // Se for string, tentar converter
              const valor = String(cellJ.v);
              const valorLimpo = valor.replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
              valorLiquido = parseFloat(valorLimpo) || 0;
            }
          }
          
          // Filtrar valores absurdos (maiores que 1 bilhão ou muito negativos)
          if (Math.abs(valorLiquido) > 1000000000) {
            console.log(`⚠️ Linha ${row + 1}: Valor muito grande, pulando: ${valorLiquido}`);
            continue;
          }
          
          // Só processar se quantidade for razoável (0 a 50 membros)
          if (quantidadeMembros < 0 || quantidadeMembros > 50) {
            console.log(`⚠️ Linha ${row + 1}: Quantidade de membros inválida (${quantidadeMembros}), pulando`);
            continue;
          }
          
          console.log(`🔍 Linha ${row + 1}: Equipe=${equipe}, Qtd=${quantidadeMembros}, Valor=${valorLiquido}`);
          
          // Agrupar por comunidade + equipe (aceita valores negativos também - são ajustes/descontos)
          if (equipe && !equipe.toLowerCase().includes('total') && valorLiquido !== 0) {
            const chave = `${comunidadeAtual}|${equipe}`;
            
            if (equipesMap.has(chave)) {
              // Se já existe, somar os valores e pegar a maior quantidade de membros
              const existente = equipesMap.get(chave)!;
              existente.valor_liquido = (existente.valor_liquido || 0) + valorLiquido;
              existente.quantidade_membros = Math.max(existente.quantidade_membros || 0, quantidadeMembros);
              console.log(`➕ Somando à equipe existente: ${equipe} - Valor ${valorLiquido >= 0 ? '+' : ''}${valorLiquido.toFixed(2)} - Novo total: R$ ${existente.valor_liquido.toFixed(2)}`);
            } else {
              // Se não existe, criar novo registro
              equipesMap.set(chave, {
                comunidade: comunidadeAtual || 'Não especificado',
                equipe: equipe,
                quantidade_membros: quantidadeMembros,
                valor_liquido: valorLiquido
              });
              console.log(`📝 Nova equipe registrada: ${equipe}`);
            }
          }
        }
        
        // Se encontrar linha "TOTAL", sair da tabela de equipe
        if (cellA && String(cellA.v).toUpperCase().trim() === 'TOTAL') {
          dentroDeTabelaEquipe = false;
        }
      } catch (error: any) {
        console.error(`❌ Erro ao processar linha ${row + 1}:`, error.message);
        erros.push(`Linha ${row + 1}: ${error.message}`);
      }
    }

    // Inserir todos os registros agrupados no banco
    console.log(`\n📊 Inserindo ${equipesMap.size} equipes no banco de dados...`);
    
    for (const [chave, empenho] of equipesMap.entries()) {
      try {
        console.log(`💾 Inserindo: ${empenho.comunidade} | ${empenho.equipe} | ${empenho.quantidade_membros} membros | R$ ${empenho.valor_liquido?.toFixed(2)}`);

        const stmt = pool.prepare(`
          INSERT INTO empenhos (mes_referencia, comunidade, equipe, quantidade_membros, valor_liquido)
          VALUES (?, ?, ?, ?, ?)
        `);
        await stmt.run(mesReferenciaFinal, empenho.comunidade, empenho.equipe, empenho.quantidade_membros, empenho.valor_liquido);

        registrosImportados++;
      } catch (insertError: any) {
        console.error(`❌ Erro ao inserir ${empenho.equipe}:`, insertError.message);
        erros.push(`Equipe ${empenho.equipe}: ${insertError.message}`);
      }
    }
    
    console.log(`✅ Processamento concluído: ${registrosImportados} registros importados`);
    
    // Processar aba "Membro" se existir
    await processarAbaMembros(workbook, mesReferenciaFinal);
    
    return {
      registros_importados: registrosImportados,
      total_linhas: range.e.r - range.s.r + 1,
      erros: erros.length > 0 ? erros : undefined
    };
  } catch (error: any) {
    console.error('❌ Erro geral:', error);
    throw new Error(`Erro ao processar planilha: ${error.message}`);
  }
};

// Processar aba "Membros Equipes" com lista detalhada de membros por equipe
async function processarAbaMembros(workbook: XLSX.WorkBook, mesReferencia: string) {
  console.log('\n📋 Processando aba de Membros...');
  
  // Procurar aba "Membros Equipes" ou variações
  let sheetNameMembro = workbook.SheetNames.find(name => 
    name.toLowerCase().includes('membro') || 
    name.toLowerCase().includes('equipe')
  );
  
  if (!sheetNameMembro) {
    console.log('⚠️ Aba de Membros não encontrada. Pulando...');
    return;
  }
  
  console.log(`✅ Aba encontrada: ${sheetNameMembro}`);
  
  const worksheet = workbook.Sheets[sheetNameMembro];
  const data = XLSX.utils.sheet_to_json(worksheet);
  
  if (data.length === 0) {
    console.log('⚠️ Aba de membros está vazia');
    return;
  }
  
  console.log(`📊 Total de linhas na aba: ${data.length}`);
  
  let membrosSalvos = 0;
  
  // Log das colunas encontradas
  if (data.length > 0) {
    console.log('📋 Colunas encontradas:', Object.keys(data[0] as any));
  }
  
  for (const row of data as any[]) {
    try {
      // Ler colunas da planilha "Membros Equipes"
      // A planilha não tem cabeçalho, então usa __EMPTY, __EMPTY_1, etc
      const termo = row['Detalhamento dos Membros das Equipes do Termos Empenhados'] || '';
      const equipe = row['__EMPTY'] || row['Equipe'] || row['BRE'] || '';
      const matricula = row['__EMPTY_1'] || row['Matricula'] || '';
      const nome = row['__EMPTY_2'] || row['Nome'] || '';
      
      // Log das primeiras linhas para debug
      if (membrosSalvos < 3) {
        console.log(`🔍 Linha ${membrosSalvos + 1}:`, { termo, equipe, matricula, nome });
      }
      
      // Pular linhas sem equipe ou nome (ou com nome "-")
      if (!equipe || !nome || nome === '-') {
        continue;
      }
      
      // Normalizar equipe (remover espaços)
      const equipeLimpa = String(equipe).trim();
      const nomeLimpo = String(nome).trim();
      const matriculaLimpa = String(matricula).trim();
      
      console.log(`📝 Salvando: ${equipeLimpa} - ${nomeLimpo} (${matriculaLimpa})`);
      
      // Inserir membro no banco
      const stmt = pool.prepare(`
        INSERT INTO membros_empenho (mes_referencia, comunidade, equipe, nome, matricula)
        VALUES (?, ?, ?, ?, ?)
      `);
      
      await stmt.run(mesReferencia, termo || '', equipeLimpa, nomeLimpo, matriculaLimpa);
      membrosSalvos++;
      
    } catch (error: any) {
      console.error(`❌ Erro ao processar membro: ${error.message}`);
    }
  }
  
  console.log(`✅ ${membrosSalvos} membros salvos no banco de dados`);
}

export const listarEmpenhos = async (mesReferencia?: string) => {
  try {
    if (mesReferencia) {
      const result = await pool.query('SELECT * FROM empenhos WHERE mes_referencia = ? ORDER BY created_at DESC', [mesReferencia]);
      return result.rows;
    }
    const result = await pool.query(`
      SELECT * FROM empenhos 
      WHERE mes_referencia = (SELECT MAX(mes_referencia) FROM empenhos)
      ORDER BY created_at DESC
    `);
    return result.rows;
  } catch (error: any) {
    throw new Error(`Erro ao listar empenhos: ${error.message}`);
  }
};

export const obterTotalEmpenho = async (mesReferencia?: string) => {
  try {
    if (mesReferencia) {
      const result = await pool.query('SELECT COALESCE(SUM(valor_liquido), 0) as total FROM empenhos WHERE mes_referencia = ?', [mesReferencia]);
      return parseFloat(result.rows[0].total);
    }
    const result = await pool.query(`
      SELECT COALESCE(SUM(valor_liquido), 0) as total FROM empenhos
      WHERE mes_referencia = (SELECT MAX(mes_referencia) FROM empenhos)
    `);
    return parseFloat(result.rows[0].total);
  } catch (error: any) {
    throw new Error(`Erro ao calcular total de empenho: ${error.message}`);
  }
};

export const listarMesesDisponiveisEmpenho = async () => {
  try {
    const result = await pool.query(`
      SELECT DISTINCT mes_referencia 
      FROM empenhos 
      ORDER BY mes_referencia DESC
    `);
    return result.rows.map(row => row.mes_referencia);
  } catch (error: any) {
    throw new Error(`Erro ao listar meses disponíveis: ${error.message}`);
  }
};

export const listarMembrosEmpenho = async (equipe: string) => {
  try {
    const result = await pool.query(
      'SELECT * FROM membros_empenho WHERE equipe = ? ORDER BY nome',
      [equipe]
    );
    return result.rows;
  } catch (error: any) {
    throw new Error(`Erro ao listar membros do empenho: ${error.message}`);
  }
};

export const analisarFuncionariosForaEmpenho = async (mesReferencia?: string) => {
  try {
    console.log('🔍 Analisando funcionários fora do empenho para o mês:', mesReferencia || 'mais recente');
    
    const params: string[] = [];
    
    // Primeiro, contar totais para debug
    let countFuncQuery = 'SELECT COUNT(*) as total FROM funcionarios';
    let countMembrosQuery = 'SELECT COUNT(*) as total FROM membros_empenho';
    
    if (mesReferencia) {
      countFuncQuery += ' WHERE mes_referencia = ?';
      countMembrosQuery += ' WHERE mes_referencia = ?';
      const totalFunc = await pool.query(countFuncQuery, [mesReferencia]);
      const totalMembros = await pool.query(countMembrosQuery, [mesReferencia]);
      const totalFuncValue = (totalFunc.rows && totalFunc.rows[0]) ? totalFunc.rows[0].total : (totalFunc as any)[0]?.total;
      const totalMembrosValue = (totalMembros.rows && totalMembros.rows[0]) ? totalMembros.rows[0].total : (totalMembros as any)[0]?.total;
      console.log(`📊 Total de funcionários no mês ${mesReferencia}: ${totalFuncValue}`);
      console.log(`📊 Total de membros empenho no mês ${mesReferencia}: ${totalMembrosValue}`);
    }
    
    // Query - comparar por nome dentro da mesma equipe
    let query = `
      SELECT DISTINCT f.* 
      FROM funcionarios f
      WHERE f.time_bre IS NOT NULL
        AND NOT EXISTS (
          SELECT 1 FROM membros_empenho m 
          WHERE m.equipe = f.time_bre
            AND UPPER(TRIM(m.nome)) = UPPER(TRIM(f.nome))
    `;
    
    if (mesReferencia) {
      query += ` AND m.mes_referencia = ?`;
      params.push(mesReferencia);
      query += `
        )
      AND f.mes_referencia = ?`;
      params.push(mesReferencia);
    } else {
      // Se não especificado, usar o mês mais recente
      query += ` AND m.mes_referencia = (SELECT MAX(mes_referencia) FROM membros_empenho)`;
      query += `
        )
      AND f.mes_referencia = (SELECT MAX(mes_referencia) FROM funcionarios)`;
    }
    
    query += `
      ORDER BY f.time_bre, f.nome
    `;
    
    console.log('🔍 Query SQL:', query);
    console.log('🔍 Parâmetros:', params);
    
    const result = await pool.query(query, params);
    const funcionarios = result.rows || result;
    
    console.log(`✅ Funcionários fora do empenho encontrados: ${funcionarios.length}`);
    
    // Log de alguns exemplos
    if (funcionarios.length > 0) {
      console.log('📝 Primeiros 5 funcionários fora do empenho:');
      funcionarios.slice(0, 5).forEach((f: any) => {
        console.log(`  - ${f.nome} (${f.matricula}) - BRE: ${f.time_bre}`);
      });
    }
    
    return funcionarios;
  } catch (error: any) {
    console.error('❌ Erro ao analisar funcionários fora do empenho:', error);
    throw new Error(`Erro ao analisar funcionários fora do empenho: ${error.message}`);
  }
};
