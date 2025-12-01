import { pool } from '../config/database';

export const createTables = async () => {
  const query = `
    CREATE TABLE IF NOT EXISTS funcionarios (
      id SERIAL PRIMARY KEY,
      contrato VARCHAR(255),
      comunidade VARCHAR(255),
      time_bre VARCHAR(255),
      gerente VARCHAR(255),
      preposto VARCHAR(255),
      nome VARCHAR(255),
      matricula VARCHAR(100),
      posto VARCHAR(255),
      grupo VARCHAR(255),
      valor_proporcional DECIMAL(10, 2),
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
  `;

  try {
    await pool.query(query);
    console.log('✅ Tabela "funcionarios" criada ou já existe');
  } catch (error) {
    console.error('❌ Erro ao criar tabela:', error);
    throw error;
  }
};
