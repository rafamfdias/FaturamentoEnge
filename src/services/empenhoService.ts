import * as XLSX from 'xlsx';
import { pool } from '../config/database';

interface EmpenhoData {
  comunidade?: string;
  equipe?: string;
  quantidade_membros?: number;
  valor_liquido?: number;
}

export const processarPlanilhaEmpenho = async (filePath: string) => {
  console.log('🔄 Iniciando processamento da planilha de empenho...');
  console.log('📁 Arquivo:', filePath);
  
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

    // Limpar tabela antes de importar novos dados
    await pool.query('DELETE FROM empenhos');
    
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
          
          // Agrupar por comunidade + equipe
          if (equipe && !equipe.toLowerCase().includes('total') && valorLiquido > 0) {
            const chave = `${comunidadeAtual}|${equipe}`;
            
            if (equipesMap.has(chave)) {
              // Se já existe, somar os valores e pegar a maior quantidade de membros
              const existente = equipesMap.get(chave)!;
              existente.valor_liquido = (existente.valor_liquido || 0) + valorLiquido;
              existente.quantidade_membros = Math.max(existente.quantidade_membros || 0, quantidadeMembros);
              console.log(`➕ Somando à equipe existente: ${equipe} - Novo total: R$ ${existente.valor_liquido}`);
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

        await pool.query(
          `INSERT INTO empenhos (comunidade, equipe, quantidade_membros, valor_liquido) 
           VALUES (?, ?, ?, ?)`,
          [empenho.comunidade, empenho.equipe, empenho.quantidade_membros, empenho.valor_liquido]
        );

        registrosImportados++;
      } catch (insertError: any) {
        console.error(`❌ Erro ao inserir ${empenho.equipe}:`, insertError.message);
        erros.push(`Equipe ${empenho.equipe}: ${insertError.message}`);
      }
    }
    
    console.log(`✅ Processamento concluído: ${registrosImportados} registros importados`);
    
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

export const listarEmpenhos = async () => {
  try {
    const result = await pool.query('SELECT * FROM empenhos ORDER BY created_at DESC');
    return result.rows;
  } catch (error: any) {
    throw new Error(`Erro ao listar empenhos: ${error.message}`);
  }
};

export const obterTotalEmpenho = async () => {
  try {
    const result = await pool.query('SELECT COALESCE(SUM(valor_liquido), 0) as total FROM empenhos');
    return parseFloat(result.rows[0].total);
  } catch (error: any) {
    throw new Error(`Erro ao calcular total de empenho: ${error.message}`);
  }
};
