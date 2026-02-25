import express from 'express';
import cors from 'cors';
import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import libre from 'libreoffice-convert';
import fetch from 'node-fetch';
import FormData from 'form-data';
import pg from 'pg';
import { spawnSync } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = process.env.PORT || 4000;
const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change';
const DB_PATH = process.env.DB_PATH ? path.resolve(process.env.DB_PATH) : path.join(__dirname, 'users.db');
const ROOT_DIR = path.resolve(__dirname, '..');
const PG_URL = process.env.DATABASE_URL || '';
const ADMIN_EMAIL = process.env.ADMIN_EMAIL || '';

// Try to ensure LibreOffice (soffice) is discoverable on macOS by preprending common path
(() => {
  const MAC_LO_PATH = '/Applications/LibreOffice.app/Contents/MacOS';
  try {
    if (fs.existsSync(MAC_LO_PATH) && !String(process.env.PATH || '').includes(MAC_LO_PATH)) {
      process.env.PATH = `${MAC_LO_PATH}:${process.env.PATH}`;
    }
  } catch {}
})();

function hasLibreOffice() {
  try {
    const r = spawnSync('soffice', ['--headless', '--version'], { stdio: 'ignore' });
    return r.status === 0;
  } catch {
    return false;
  }
}

async function hasGotenberg() {
  const base = process.env.GOTENBERG_URL;
  if (!base) return false;
  try {
    const url = `${base.replace(/\/+$/, '')}/health`;
    const ac = new AbortController();
    const to = setTimeout(() => ac.abort(), 2000);
    const r = await fetch(url, { signal: ac.signal });
    clearTimeout(to);
    return r.ok;
  } catch {
    return false;
  }
}

let dbMode = 'sqlite';
let db = null;
let pool = null;
let sqlite3 = null; // will be loaded dynamically only if needed

async function initDb() {
  if (PG_URL) {
    pool = new pg.Pool({ connectionString: PG_URL, ssl: { rejectUnauthorized: false } });
    await pool.query(`
      CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        created_at TEXT NOT NULL
      )
    `);
    await pool.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS first_name TEXT`);
    await pool.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS last_name TEXT`);
    await pool.query(`
      CREATE TABLE IF NOT EXISTS login_entries (
        id SERIAL PRIMARY KEY,
        user_id INTEGER,
        email TEXT,
        first_name TEXT,
        last_name TEXT,
        created_at TEXT NOT NULL
      )
    `);
    dbMode = 'pg';
    return;
  }
  // Lazy-load sqlite3 only when Postgres is not configured (e.g., local dev)
  const sqlite = await import('sqlite3');
  sqlite3 = sqlite.default;
  sqlite3.verbose();
  db = new sqlite3.Database(DB_PATH);
  await new Promise((resolve, reject) => {
    db.run(
      `CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        created_at TEXT NOT NULL
      )`,
      (err) => (err ? reject(err) : resolve())
    );
  });
  await new Promise((resolve) => {
    db.run(`ALTER TABLE users ADD COLUMN first_name TEXT`, () => resolve());
  });
  await new Promise((resolve) => {
    db.run(`ALTER TABLE users ADD COLUMN last_name TEXT`, () => resolve());
  });
  await new Promise((resolve, reject) => {
    db.run(
      `CREATE TABLE IF NOT EXISTS login_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        email TEXT,
        first_name TEXT,
        last_name TEXT,
        created_at TEXT NOT NULL
      )`,
      (err) => (err ? reject(err) : resolve())
    );
  });
  dbMode = 'sqlite';
}

async function userGetByEmail(email) {
  if (dbMode === 'pg') {
    const r = await pool.query('SELECT * FROM users WHERE email = $1', [email]);
    return r.rows[0] || null;
  }
  return await new Promise((resolve, reject) => {
    db.get('SELECT * FROM users WHERE email = ?', [email], (err, row) => {
      if (err) return reject(err);
      resolve(row || null);
    });
  });
}

async function userInsert(email, password_hash, created_at, first_name, last_name) {
  if (dbMode === 'pg') {
    await pool.query(
      'INSERT INTO users (email, password_hash, created_at, first_name, last_name) VALUES ($1, $2, $3, $4, $5)',
      [email, password_hash, created_at, first_name || null, last_name || null]
    );
    return;
  }
  await new Promise((resolve, reject) => {
    const stmt = db.prepare('INSERT INTO users (email, password_hash, created_at, first_name, last_name) VALUES (?, ?, ?, ?, ?)');
    stmt.run(email, password_hash, created_at, first_name || null, last_name || null, (err) => {
      if (err) return reject(err);
      resolve();
    });
  });
}

async function userUpdatePassword(email, password_hash) {
  if (dbMode === 'pg') {
    await pool.query('UPDATE users SET password_hash = $1 WHERE email = $2', [password_hash, email]);
    return;
  }
  await new Promise((resolve, reject) => {
    const stmt = db.prepare('UPDATE users SET password_hash = ? WHERE email = ?');
    stmt.run(password_hash, email, (err) => {
      if (err) return reject(err);
      resolve();
    });
  });
}
const app = express();
app.use(cors({
  origin: (origin, cb) => {
    // Allow requests from localhost and 127.0.0.1 during development
    if (!origin) return cb(null, true);
    const allow = [/^http:\/\/localhost:\d+$/, /^http:\/\/127\.0\.0\.1:\d+$/];
    if (allow.some((re) => re.test(origin))) return cb(null, true);
    return cb(null, false);
  },
  credentials: true,
}));
app.use(express.json());

app.get('/api/health', async (_req, res) => {
  res.json({
    ok: true,
    ts: Date.now(),
    libreoffice: hasLibreOffice(),
    gotenberg: await hasGotenberg()
  });
});

function requireAdmin(req, res, next) {
  try {
    const h = req.headers.authorization || '';
    const m = h.match(/^Bearer\s+(.+)$/i);
    if (!m) return res.status(401).json({ error: 'Unauthorized' });
    const payload = jwt.verify(m[1], JWT_SECRET);
    if (!ADMIN_EMAIL || payload.email !== ADMIN_EMAIL) return res.status(403).json({ error: 'Forbidden' });
    req.user = payload;
    return next();
  } catch {
    return res.status(401).json({ error: 'Unauthorized' });
  }
}

async function logLogin(user) {
  const createdAt = new Date().toISOString();
  if (dbMode === 'pg') {
    await pool.query(
      'INSERT INTO login_entries (user_id, email, first_name, last_name, created_at) VALUES ($1, $2, $3, $4, $5)',
      [user.id, user.email, user.first_name || null, user.last_name || null, createdAt]
    );
    await pool.query(`DELETE FROM login_entries WHERE created_at < $1`, [new Date(Date.now() - 60 * 24 * 3600 * 1000).toISOString()]);
    return;
  }
  await new Promise((resolve, reject) => {
    const stmt = db.prepare('INSERT INTO login_entries (user_id, email, first_name, last_name, created_at) VALUES (?, ?, ?, ?, ?)');
    stmt.run(user.id, user.email, user.first_name || null, user.last_name || null, createdAt, (err) => {
      if (err) return reject(err);
      resolve();
    });
  });
  await new Promise((resolve) => {
    const cutoff = new Date(Date.now() - 60 * 24 * 3600 * 1000).toISOString();
    db.run('DELETE FROM login_entries WHERE created_at < ?', [cutoff], () => resolve());
  });
}

function detectDelimiters(xml) {
  const hasCurly = /(\{ITEM\})|(\{VALOR\})/i.test(xml);
  const hasMustache = /(\{\{\s*ITEM\s*\}\})|(\{\{\s*VALOR\s*\}\})/i.test(xml);
  const hasBrackets = /(\[\[\s*ITEM\s*\]\])|(\[\[\s*VALOR\s*\]\])/i.test(xml);
  const hasGuillemets = /(«\s*ITEM\s*»)|(«\s*VALOR\s*»)/i.test(xml);
  const baseOpts = { paragraphLoop: true, linebreaks: true };
  if (hasMustache) return { ...baseOpts, delimiters: { start: '{{', end: '}}' } };
  if (hasBrackets) return { ...baseOpts, delimiters: { start: '[[', end: ']]' } };
  if (hasGuillemets) return { ...baseOpts, delimiters: { start: '«', end: '»' } };
  return baseOpts;
}

function templatePathForModel(model) {
  return model === 'wellington'
    ? path.join(ROOT_DIR, 'modelo_placas', 'wellington.docx')
    : path.join(ROOT_DIR, 'modelo_placas', 'patricia.docx');
}

function generateDocxBuffer({ model, item, valor }) {
  const p = templatePathForModel(model);
  if (!fs.existsSync(p)) throw new Error('Template DOCX não encontrado');
  const content = fs.readFileSync(p);
  const zip = new PizZip(content);
  let xml = '';
  try { xml = zip.file('word/document.xml').asText(); } catch {}
  const opts = detectDelimiters(xml);
  const doc = new Docxtemplater(zip, opts);
  doc.setData({ ITEM: String(item), VALOR: String(valor) });
  doc.render();
  const buf = doc.getZip().generate({ type: 'nodebuffer' });
  return buf;
}

async function convertWithGotenberg(buf) {
  const base = process.env.GOTENBERG_URL;
  if (!base) return null;
  const endpoint = `${base.replace(/\/+$/, '')}/forms/libreoffice/convert`;
  const form = new FormData();
  form.append('files', buf, {
    filename: 'document.docx',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const res = await fetch(endpoint, {
    method: 'POST',
    body: form,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Gotenberg falhou (${res.status}): ${text.slice(0,200)}`);
  }
  const pdf = await res.arrayBuffer();
  return Buffer.from(pdf);
}

app.post('/api/generate/docx', (req, res) => {
  try {
    const { model, item, valor } = req.body || {};
    if (!model || !item || !valor) return res.status(400).json({ error: 'Parâmetros inválidos' });
    const buf = generateDocxBuffer({ model, item, valor });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${model}-item-${item}.docx"`);
    return res.send(buf);
  } catch (e) {
    return res.status(500).json({ error: e.message || 'Falha ao gerar DOCX' });
  }
});

app.post('/api/generate/pdf', async (req, res) => {
  try {
    const { model, item, valor } = req.body || {};
    if (!model || !item || !valor) return res.status(400).json({ error: 'Parâmetros inválidos' });
    const buf = generateDocxBuffer({ model, item, valor });
    // Tenta via Gotenberg se configurado
    if (process.env.GOTENBERG_URL) {
      try {
        const pdf = await convertWithGotenberg(buf);
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="${model}-item-${item}.pdf"`);
        return res.send(pdf);
      } catch (err) {
        // Continua para fallback local
        console.warn('Gotenberg erro:', err.message);
      }
    }
    // Fallback: LibreOffice local
    libre.convert(buf, '.pdf', undefined, (err, done) => {
      if (err) {
        const hint = process.env.GOTENBERG_URL
          ? 'Falha via Gotenberg e sem LibreOffice. Verifique GOTENBERG_URL e o /health do serviço Gotenberg.'
          : 'Falha ao converter para PDF. Configure GOTENBERG_URL ou instale o LibreOffice.';
        return res.status(500).json({ error: hint });
      }
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${model}-item-${item}.pdf"`);
      return res.send(done);
    });
  } catch (e) {
    return res.status(500).json({ error: e.message || 'Falha ao gerar PDF' });
  }
});

app.post('/api/register', async (req, res) => {
  try {
    const { firstName, lastName, email, password } = req.body || {};
    if (!firstName || !lastName || !email || !password) return res.status(400).json({ error: 'Dados inválidos' });
    const normalized = String(email).trim().toLowerCase();
    const exists = await userGetByEmail(normalized);
    if (exists) return res.status(409).json({ error: 'E-mail já cadastrado' });
    const hash = bcrypt.hashSync(password, 10);
    const createdAt = new Date().toISOString();
    await userInsert(normalized, hash, createdAt, String(firstName).trim(), String(lastName).trim());
    return res.json({ ok: true });
  } catch (err) {
    return res.status(500).json({ error: 'Erro ao cadastrar' });
  }
});

app.post('/api/login', async (req, res) => {
  try {
    const { email, password } = req.body || {};
    if (!email || !password) return res.status(400).json({ error: 'Dados inválidos' });
    const normalized = String(email).trim().toLowerCase();
    const row = await userGetByEmail(normalized);
    if (!row) return res.status(404).json({ error: 'Usuário não encontrado' });
    const ok = bcrypt.compareSync(password, row.password_hash);
    if (!ok) return res.status(401).json({ error: 'Senha incorreta' });
    const token = jwt.sign({ sub: row.id, email: row.email }, JWT_SECRET, { expiresIn: '7d' });
    await logLogin({ id: row.id, email: row.email, first_name: row.first_name, last_name: row.last_name });
    return res.json({ ok: true, token, email: row.email });
  } catch (err) {
    return res.status(500).json({ error: 'Erro no servidor' });
  }
});

app.get('/api/admin/logins', requireAdmin, async (req, res) => {
  try {
    const cutoff = new Date(Date.now() - 60 * 24 * 3600 * 1000).toISOString();
    if (dbMode === 'pg') {
      const r = await pool.query('SELECT * FROM login_entries WHERE created_at >= $1 ORDER BY created_at DESC', [cutoff]);
      return res.json({ ok: true, entries: r.rows });
    }
    const entries = await new Promise((resolve, reject) => {
      db.all('SELECT * FROM login_entries WHERE created_at >= ? ORDER BY created_at DESC', [cutoff], (err, rows) => {
        if (err) return reject(err);
        resolve(rows || []);
      });
    });
    return res.json({ ok: true, entries });
  } catch {
    return res.status(500).json({ error: 'Erro ao listar' });
  }
});
app.post('/api/reset-password', async (req, res) => {
  try {
    const { email, newPassword } = req.body || {};
    if (!email || !newPassword) return res.status(400).json({ error: 'Dados inválidos' });
    const normalized = String(email).trim().toLowerCase();
    const user = await userGetByEmail(normalized);
    if (!user) return res.status(404).json({ error: 'Usuário não encontrado' });
    const hash = bcrypt.hashSync(newPassword, 10);
    await userUpdatePassword(normalized, hash);
    return res.json({ ok: true });
  } catch (e) {
    return res.status(500).json({ error: 'Erro ao resetar senha' });
  }
});
// Serve frontend estático do diretório raiz do projeto
app.use(express.static(ROOT_DIR, { extensions: ['html'] }));
app.get('*', (req, res) => {
  res.sendFile(path.join(ROOT_DIR, 'index.html'));
});

initDb()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`API listening on http://localhost:${PORT}`);
    });
    (async () => {
      try {
        if (dbMode === 'pg' && pool) {
          await pool.query('SELECT 1');
        }
        await hasGotenberg();
      } catch {}
    })();
  })
  .catch((e) => {
    console.error('DB init failed', e);
    process.exit(1);
  });
