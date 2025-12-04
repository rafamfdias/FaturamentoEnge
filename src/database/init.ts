import { pool } from '../config/database';

export const createTables = async () => {
  const queries = [
    `CREATE TABLE IF NOT EXISTS funcionarios (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      mes_referencia TEXT NOT NULL,
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
    )`,
    `CREATE TABLE IF NOT EXISTS empenhos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      mes_referencia TEXT NOT NULL,
      comunidade TEXT,
      equipe TEXT,
      quantidade_membros REAL,
      valor_liquido REAL,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`,
    `CREATE TABLE IF NOT EXISTS membros_empenho (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      mes_referencia TEXT NOT NULL,
      comunidade TEXT,
      equipe TEXT,
      nome TEXT,
      matricula TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`
  ];

  try {
    for (const query of queries) {
      await pool.exec(query);
    }
    console.log('✅ Tabelas "funcionarios", "empenhos" e "membros_empenho" criadas ou já existem');
  } catch (error) {
    console.error('❌ Erro ao criar tabelas:', error);
    throw error;
  }
};
