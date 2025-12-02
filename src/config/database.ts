import sqlite3 from 'sqlite3';
import { open, Database } from 'sqlite';
import path from 'path';

let db: Database | null = null;

export const getDatabase = async () => {
  if (db) {
    return db;
  }

  db = await open({
    filename: path.join(__dirname, '../../faturamento.db'),
    driver: sqlite3.Database
  });

  console.log('✅ Conectado ao banco de dados SQLite');
  return db;
};

// Manter compatibilidade com código que usa pool
export const pool = {
  query: async (text: string, params?: any[]) => {
    const database = await getDatabase();
    
    // Converter sintaxe PostgreSQL ($1, $2) para SQLite (?, ?)
    const sqliteQuery = text.replace(/\$(\d+)/g, '?');
    
    if (text.trim().toUpperCase().startsWith('SELECT') || text.trim().toUpperCase().startsWith('DELETE') || text.trim().toUpperCase().startsWith('TRUNCATE')) {
      const rows = await database.all(sqliteQuery, params);
      return { rows };
    } else {
      await database.run(sqliteQuery, params);
      return { rows: [] };
    }
  }
};
