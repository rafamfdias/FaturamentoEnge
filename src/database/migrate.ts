import { pool } from '../config/database';

export const recreateTables = async () => {
  try {
    console.log('🔄 Recriando tabelas...');
    
    // Dropar tabelas existentes
    await pool.query('DROP TABLE IF EXISTS empenhos');
    await pool.query('DROP TABLE IF EXISTS funcionarios');
    
    // Criar tabelas com nova estrutura
    await pool.query(`
      CREATE TABLE IF NOT EXISTS funcionarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contrato TEXT,
        comunidade TEXT,
        time_bre TEXT,
        gerente TEXT,
        preposto TEXT,
        nome TEXT,
        matricula TEXT,
        posto TEXT,
        grupo TEXT,
        valor_proporcional REAL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
      )
    `);
    
    await pool.query(`
      CREATE TABLE IF NOT EXISTS empenhos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        comunidade TEXT,
        equipe TEXT,
        quantidade_membros REAL,
        valor_liquido REAL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
      )
    `);
    
    console.log('✅ Tabelas recriadas com sucesso!');
  } catch (error) {
    console.error('❌ Erro ao recriar tabelas:', error);
    throw error;
  }
};

// Se executado diretamente
if (require.main === module) {
  recreateTables()
    .then(() => {
      console.log('✅ Migração concluída');
      process.exit(0);
    })
    .catch((error) => {
      console.error('❌ Erro na migração:', error);
      process.exit(1);
    });
}
