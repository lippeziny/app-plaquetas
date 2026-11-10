const Database = require('better-sqlite3');
const path = require('node:path');
const fs = require('node:fs');

const dbPath = process.env.DATABASE_PATH
  ? path.resolve(process.env.DATABASE_PATH)
  : path.join(__dirname, 'data', 'plaquetas.sqlite');
fs.mkdirSync(path.dirname(dbPath), { recursive: true });

const db = new Database(dbPath);
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');
db.exec(`
  CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL,
    name TEXT NOT NULL,
    quantity TEXT DEFAULT '',
    sheet_name TEXT DEFAULT '',
    source_file TEXT DEFAULT '',
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
  );
  CREATE TABLE IF NOT EXISTS items_stage (
    code TEXT NOT NULL, name TEXT NOT NULL, quantity TEXT DEFAULT '',
    sheet_name TEXT DEFAULT '', source_file TEXT DEFAULT ''
  );
  CREATE INDEX IF NOT EXISTS idx_items_code ON items(code);
  CREATE INDEX IF NOT EXISTS idx_items_name ON items(name COLLATE NOCASE);
  CREATE TABLE IF NOT EXISTS counts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL,
    name TEXT DEFAULT '',
    quantity INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(code)
  );
`);

module.exports = db;
