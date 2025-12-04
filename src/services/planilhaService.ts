import xlsx from 'xlsx';
import path from 'path';
import { pool } from '../config/database';
import { Funcionario } from '../types/funcionario';
import { extrairMesReferencia } from '../utils/mesReferencia';

export class PlanilhaService {
  async processarPlanilha(filePath: string, mesReferencia: string | null, nomeOriginal?: string): Promise<{ sucesso: number; erros: string[] }> {
    const nomeArquivo = nomeOriginal || path.basename(filePath);
    const mesRef = mesReferencia || extrairMesReferencia(nomeArquivo);
    console.log(`📅 Mês detectado de "${nomeArquivo}": ${mesRef}`);
    const mesReferenciaFinal = mesRef;
    try {
      // Ler arquivo Excel
      const workbook = xlsx.readFile(filePath);
      
      // Procurar pela aba "faturamento Novembro_2025"
      let sheetName = workbook.SheetNames.find(name => 
        name.toLowerCase().includes('faturamento') && name.toLowerCase().includes('novembro')
      );
      
      // Se não encontrar, usar a primeira aba
      if (!sheetName) {
        sheetName = workbook.SheetNames[0];
        console.log(`⚠️ Aba "faturamento Novembro_2025" não encontrada. Usando: ${sheetName}`);
        console.log(`📋 Abas disponíveis: ${workbook.SheetNames.join(', ')}`);
      } else {
        console.log(`✅ Usando aba: ${sheetName}`);
      }
      
      const worksheet = workbook.Sheets[sheetName];
      
      // Tentar diferentes formas de ler a planilha
      // Primeiro: ler normalmente
      let data: any[] = xlsx.utils.sheet_to_json(worksheet);
      
      console.log(`📊 Tentativa 1 - Total de linhas: ${data.length}`);
      if (data.length > 0) {
        console.log('📋 Colunas detectadas:', Object.keys(data[0]).join(', '));
      }
      
      // Se as colunas não forem as esperadas, tentar pular linhas
      if (data.length > 0 && !data[0]['CONTRATO'] && !data[0]['NOME']) {
        console.log('⚠️ Colunas não encontradas, tentando com range...');
        // Tentar começar da linha 3 (índice 2)
        data = xlsx.utils.sheet_to_json(worksheet, { range: 2 });
        console.log(`📊 Tentativa 2 - Total de linhas: ${data.length}`);
        if (data.length > 0) {
          console.log('📋 Colunas detectadas:', Object.keys(data[0]).join(', '));
        }
      }
      
      // Limpar dados anteriores SOMENTE do mês atual
      const deleteStmt = pool.prepare('DELETE FROM funcionarios WHERE mes_referencia = ?');
      await deleteStmt.run(mesReferenciaFinal);
      console.log(`🗑️ Dados anteriores do mês ${mesReferenciaFinal} removidos`);
      
      let sucesso = 0;
      const erros: string[] = [];

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        
        try {
          // Mapear colunas da planilha para os campos corretos
          const funcionario: Funcionario = {
            contrato: row['OS'] || row['CONTRATO'] || row['contrato'] || '',
            comunidade: row['SIGLA'] || row['COMUNIDADE'] || row['comunidade'] || '',
            time_bre: row['Time'] || row['TIME(BRE)'] || row['TIME'] || row['time'] || '',
            gerente: row['Gerente'] || row['GERENTE'] || row['gerente'] || '',
            preposto: row['Preposto'] || row['PREPOSTO'] || row['preposto'] || '',
            nome: row['NOME'] || row['nome'] || '',
            matricula: row['MATRICULA'] || row['matricula'] || '',
            posto: row['POSTO'] || row['posto'] || '',
            grupo: row['GRUPO'] || row['grupo'] || '',
            valor_proporcional: parseFloat(row['Valor do Posto proporcional'] || row['VALOR PROPORCIONAL'] || row['valor_proporcional'] || '0')
          };

          // Validação básica - pular linhas vazias ou sem dados importantes
          if (!funcionario.nome && !funcionario.matricula && !funcionario.contrato) {
            continue; // Pular linha completamente vazia
          }
          
          // Pular linhas onde não tem nome E não tem matrícula
          if (!funcionario.nome && !funcionario.matricula) {
            continue;
          }
          
          // Pular linhas com valor proporcional zerado E sem contrato
          if (funcionario.valor_proporcional === 0 && !funcionario.contrato) {
            continue;
          }

          // Inserir no banco de dados
          await this.salvarFuncionario(funcionario, mesReferenciaFinal);
          sucesso++;
        } catch (error: any) {
          erros.push(`Linha ${i + 2}: ${error.message}`);
        }
      }

      return { sucesso, erros };
    } catch (error: any) {
      throw new Error(`Erro ao processar planilha: ${error.message}`);
    }
  }

  private async salvarFuncionario(funcionario: Funcionario, mesReferencia: string): Promise<void> {
    // Verificar se já existe registro com mesma matrícula no mesmo mês
    if (funcionario.matricula) {
      const checkQuery = `
        SELECT id FROM funcionarios 
        WHERE matricula = ? AND nome = ? AND mes_referencia = ?
        LIMIT 1
      `;
      const existing = await pool.query(checkQuery, [funcionario.matricula, funcionario.nome, mesReferencia]);
      
      if (existing.rows.length > 0) {
        // Já existe, não inserir duplicado
        return;
      }
    }
    
    const query = `
      INSERT INTO funcionarios 
      (mes_referencia, contrato, comunidade, time_bre, gerente, preposto, nome, matricula, posto, grupo, valor_proporcional)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `;

    const values = [
      mesReferencia,
      funcionario.contrato,
      funcionario.comunidade,
      funcionario.time_bre,
      funcionario.gerente,
      funcionario.preposto,
      funcionario.nome,
      funcionario.matricula,
      funcionario.posto,
      funcionario.grupo,
      funcionario.valor_proporcional
    ];

    await pool.query(query, values);
  }

  async listarFuncionarios(mesReferencia?: string): Promise<Funcionario[]> {
    if (mesReferencia) {
      const result = await pool.query('SELECT * FROM funcionarios WHERE mes_referencia = ? ORDER BY nome ASC', [mesReferencia]);
      return result.rows;
    }
    // Se não especificado, pegar o mês mais recente
    const result = await pool.query(`
      SELECT * FROM funcionarios 
      WHERE mes_referencia = (SELECT MAX(mes_referencia) FROM funcionarios)
      ORDER BY nome ASC
    `);
    return result.rows;
  }

  async obterTotalValorProporcional(mesReferencia?: string): Promise<number> {
    let query = 'SELECT COALESCE(SUM(valor_proporcional), 0) as total FROM funcionarios';
    const params: string[] = [];
    
    if (mesReferencia) {
      query += ' WHERE mes_referencia = ?';
      params.push(mesReferencia);
    } else {
      // Se não especificado, pegar o mês mais recente
      query += ' WHERE mes_referencia = (SELECT MAX(mes_referencia) FROM funcionarios)';
    }
    
    const result = await pool.query(query, params);
    const total = parseFloat(result.rows[0].total);
    console.log(`💰 Total calculado: R$ ${total.toFixed(2)}`);
    return total;
  }

  async listarMesesDisponiveis(): Promise<string[]> {
    const result = await pool.all(`
      SELECT DISTINCT mes_referencia 
      FROM funcionarios 
      UNION
      SELECT DISTINCT mes_referencia 
      FROM empenhos
      ORDER BY mes_referencia DESC
    `);
    return result.map((row: any) => row.mes_referencia);
  }

  async limparDados(): Promise<void> {
    await pool.query('DELETE FROM funcionarios');
  }
}

// Exportar função auxiliar para uso em outras rotas
export const obterTotalValorProporcional = async (): Promise<number> => {
  const result = await pool.query('SELECT COALESCE(SUM(valor_proporcional), 0) as total FROM funcionarios');
  return parseFloat(result.rows[0].total);
};
