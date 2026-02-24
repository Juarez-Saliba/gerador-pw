import express from 'express';
import cors from 'cors';
import sqlite3 from 'sqlite3';
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
import { spawnSync } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = process.env.PORT || 4000;
const JWT_SECRET = process.env.JWT_SECRET || 'dev-secret-change';
const DB_PATH = process.env.DB_PATH ? path.resolve(process.env.DB_PATH) : path.join(__dirname, 'users.db');
const ROOT_DIR = path.resolve(__dirname, '..');

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

sqlite3.verbose();
const db = new sqlite3.Database(DB_PATH);

db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      email TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      created_at TEXT NOT NULL
    )
  `);
});

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

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, ts: Date.now(), libreoffice: hasLibreOffice() });
});

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
        return res.status(500).json({ error: 'Falha ao converter para PDF. Configure GOTENBERG_URL ou instale o LibreOffice.' });
      }
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${model}-item-${item}.pdf"`);
      return res.send(done);
    });
  } catch (e) {
    return res.status(500).json({ error: e.message || 'Falha ao gerar PDF' });
  }
});

app.post('/api/register', (req, res) => {
  const { email, password } = req.body || {};
  if (!email || !password) return res.status(400).json({ error: 'Dados inválidos' });
  const normalized = String(email).trim().toLowerCase();
  const hash = bcrypt.hashSync(password, 10);
  const createdAt = new Date().toISOString();
  const stmt = db.prepare('INSERT INTO users (email, password_hash, created_at) VALUES (?, ?, ?)');
  stmt.run(normalized, hash, createdAt, (err) => {
    if (err) {
      if (err.message && err.message.includes('UNIQUE')) {
        return res.status(409).json({ error: 'E-mail já cadastrado' });
      }
      return res.status(500).json({ error: 'Erro ao cadastrar' });
    }
    return res.json({ ok: true });
  });
});

app.post('/api/login', (req, res) => {
  const { email, password } = req.body || {};
  if (!email || !password) return res.status(400).json({ error: 'Dados inválidos' });
  const normalized = String(email).trim().toLowerCase();
  db.get('SELECT * FROM users WHERE email = ?', [normalized], (err, row) => {
    if (err) return res.status(500).json({ error: 'Erro no servidor' });
    if (!row) return res.status(404).json({ error: 'Usuário não encontrado' });
    const ok = bcrypt.compareSync(password, row.password_hash);
    if (!ok) return res.status(401).json({ error: 'Senha incorreta' });
    const token = jwt.sign({ sub: row.id, email: row.email }, JWT_SECRET, { expiresIn: '7d' });
    return res.json({ ok: true, token, email: row.email });
  });
});

// Serve frontend estático do diretório raiz do projeto
app.use(express.static(ROOT_DIR, { extensions: ['html'] }));
app.get('*', (req, res) => {
  res.sendFile(path.join(ROOT_DIR, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`API listening on http://localhost:${PORT}`);
});
