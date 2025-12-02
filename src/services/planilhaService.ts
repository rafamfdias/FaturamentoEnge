import xlsx from 'xlsx';
import { pool } from '../config/database';
import { Funcionario } from '../types/funcionario';

export class PlanilhaService {
  async processarPlanilha(filePath: string): Promise<{ sucesso: number; erros: string[] }> {
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
          await this.salvarFuncionario(funcionario);
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

  private async salvarFuncionario(funcionario: Funcionario): Promise<void> {
    // Verificar se já existe registro com mesma matrícula
    if (funcionario.matricula) {
      const checkQuery = `
        SELECT id FROM funcionarios 
        WHERE matricula = $1 AND nome = $2
        LIMIT 1
      `;
      const existing = await pool.query(checkQuery, [funcionario.matricula, funcionario.nome]);
      
      if (existing.rows.length > 0) {
        // Já existe, não inserir duplicado
        return;
      }
    }
    
    const query = `
      INSERT INTO funcionarios 
      (contrato, comunidade, time_bre, gerente, preposto, nome, matricula, posto, grupo, valor_proporcional)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)
    `;

    const values = [
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

  async listarFuncionarios(): Promise<Funcionario[]> {
    const result = await pool.query('SELECT * FROM funcionarios ORDER BY nome ASC');
    return result.rows;
  }

  async obterTotalValorProporcional(): Promise<number> {
    const result = await pool.query('SELECT COALESCE(SUM(valor_proporcional), 0) as total FROM funcionarios');
    const total = parseFloat(result.rows[0].total);
    console.log(`💰 Total calculado: R$ ${total.toFixed(2)}`);
    return total;
  }

  async limparDados(): Promise<void> {
    await pool.query('TRUNCATE TABLE funcionarios RESTART IDENTITY');
  }
}

// Exportar função auxiliar para uso em outras rotas
export const obterTotalValorProporcional = async (): Promise<number> => {
  const result = await pool.query('SELECT COALESCE(SUM(valor_proporcional), 0) as total FROM funcionarios');
  return parseFloat(result.rows[0].total);
};
