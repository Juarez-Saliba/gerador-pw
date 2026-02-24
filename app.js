const $ = (q) => document.querySelector(q);
const $$ = (q) => [...document.querySelectorAll(q)];

const state = {
  user: null,
  items: [],
  cardsRendered: false,
  generationTimer: null,
};

function show(el) { el.classList.remove('hidden'); }
function hide(el) { el.classList.add('hidden'); }

function formatCurrencyBRL(str) {
  try {
    const n = toNumberFromBRL(str);
    return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  } catch { return str; }
}

function toNumberFromBRL(str) {
  const s = String(str).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  if (Number.isFinite(n)) return n;
  throw new Error('valor inválido');
}

const API_BASE = '';

function setGenerationStatus(text, isError) {
  const banner = $('#generationStatus');
  const bannerText = $('#generationStatusText');
  if (!banner || !bannerText) return;
  bannerText.textContent = text || '';
  banner.classList.toggle('error', !!isError);
  show(banner);
  try { banner.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); } catch {}
  const timerBar = banner.querySelector('.status-timer');
  if (timerBar) {
    timerBar.style.animation = 'none';
    // force reflow to restart animation
    // eslint-disable-next-line no-unused-expressions
    timerBar.offsetHeight;
    timerBar.style.animation = 'shrinkTimer 8s linear forwards';
  }
  if (state.generationTimer) {
    clearTimeout(state.generationTimer);
    state.generationTimer = null;
  }
  state.generationTimer = setTimeout(() => {
    hide(banner);
    bannerText.textContent = '';
    if (timerBar) timerBar.style.animation = 'none';
  }, 8000);
}

function setSession(email) {
  state.user = email;
  localStorage.setItem('session-user', email);
  $('#logoutBtn').classList.remove('hidden');
  $('#authView').classList.add('hidden');
  $('#workspaceView').classList.remove('hidden');
}
function clearSession() {
  state.user = null;
  localStorage.removeItem('session-user');
  localStorage.removeItem('session-token');
  $('#logoutBtn').classList.add('hidden');
  $('#workspaceView').classList.add('hidden');
  $('#authView').classList.remove('hidden');
}

function bootSession() {
  const u = localStorage.getItem('session-user');
  if (u) setSession(u);
}

function bindAuth() {
  const tabs = $$('.tab');
  const forms = $$('.form');
  tabs.forEach(t => {
    t.addEventListener('click', () => {
      tabs.forEach(x => x.classList.remove('active'));
      t.classList.add('active');
      forms.forEach(f => f.classList.remove('visible'));
      const id = t.dataset.tab === 'login' ? '#loginForm' : '#registerForm';
      $(id).classList.add('visible');
    });
  });

  $('#registerForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = e.target.email.value.trim().toLowerCase();
    const password = e.target.password.value;
    try {
      const res = await fetch(`${API_BASE}/api/register`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Falha no cadastro');
      alert('Cadastro realizado. Faça login.');
      tabs[0].click();
    } catch (err) {
      alert(err.message || 'Erro no cadastro');
    }
  });

  $('#loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = e.target.email.value.trim().toLowerCase();
    const password = e.target.password.value;
    try {
      const res = await fetch(`${API_BASE}/api/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Falha no login');
      if (data.token) localStorage.setItem('session-token', data.token);
      setSession(email);
    } catch (err) {
      alert(err.message || 'Erro no login');
    }
  });

  $('#logoutBtn').addEventListener('click', () => {
    clearSession();
  });
}

function parseTextTable(raw) {
  const lines = raw.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const items = [];
  for (const line of lines) {
    const mItem = line.match(/^\s*(\d{1,4})\b/);
    if (!mItem) continue;
    const itemNum = parseInt(mItem[1], 10);

    const moneyMatches = [...line.matchAll(/R\$\s*([\d\.\s]*\d,\d{2})/g)];
    if (moneyMatches.length === 0) continue;
    const lastMoney = moneyMatches[moneyMatches.length - 1][1];
    const valor = formatCurrencyBRL(lastMoney);

    items.push({ item: itemNum, valor });
  }
  return items;
}

function renderTablePreview(items) {
  const container = $('#tablePreview');
  container.innerHTML = '';
  const head = document.createElement('div');
  head.className = 'table-row table-head';
  head.innerHTML = `
    <div class="table-cell">Item</div>
    <div class="table-cell">Descrição</div>
    <div class="table-cell">Avaliação</div>
  `;
  container.appendChild(head);
  items.forEach(it => {
    const row = document.createElement('div');
    row.className = 'table-row';
    row.innerHTML = `
      <div class="table-cell">${it.item}</div>
      <div class="table-cell">—</div>
      <div class="table-cell">${it.valor}</div>
    `;
    container.appendChild(row);
  });
}

function cardClassForModel(model) {
  if (model === 'wellington') return 'model-wellington';
  if (model === 'patricia') return 'model-patricia';
  return '';
}

function renderCards(items) {
  const model = $('#modeloSelect').value;
  const grid = $('#cardsGrid');
  grid.innerHTML = '';
  items.forEach(it => {
    const card = document.createElement('div');
    card.className = `card ${cardClassForModel(model)} fallback`;
    card.dataset.item = String(it.item);
    card.dataset.valor = it.valor;
    card.dataset.model = model;
    card.innerHTML = `
      <div class="background"></div>
      <div class="model tag">${model === 'wellington' ? 'Wellington Silva' : 'Patrícia G. de Andrade'}</div>
      <div class="content">
        <div class="badge">ITEM ${it.item}</div>
        <div class="price">${it.valor}</div>
      </div>
      <div class="actions">
        <button class="btn mini download-docx">DOCX</button>
        <button class="btn mini download-pdf">PDF</button>
      </div>
    `;
    grid.appendChild(card);
  });
  state.cardsRendered = true;
  const disabled = items.length === 0;
  $('#downloadAllDocxBtn').disabled = disabled;
  $('#downloadAllPdfBtn').disabled = disabled;
  show($('#downloadAllDocxBtn'));
  show($('#downloadAllPdfBtn'));
}

async function canvasInfoFromCard(cardEl) {
  const canvas = await html2canvas(cardEl, { backgroundColor: null, scale: 2 });
  const blob = await new Promise(res => canvas.toBlob(res, 'image/png'));
  const buf = await blob.arrayBuffer();
  return { blob, buf, width: canvas.width, height: canvas.height };
}

async function generateDocxFromCard(cardEl) {
  if (!window.docx) throw new Error('biblioteca DOCX não carregada');
  const { buf, width, height } = await canvasInfoFromCard(cardEl);
  const doc = new window.docx.Document({
    sections: [
      {
        properties: { page: { size: { orientation: window.docx.PageOrientation.LANDSCAPE } } },
        children: [
          new window.docx.Paragraph({
            spacing: { before: 0, after: 0, line: 240 },
            children: [
              new window.docx.ImageRun({
                data: buf,
                transformation: { width, height },
              }),
            ],
          }),
        ],
      },
    ],
  });
  const out = await window.docx.Packer.toBlob(doc);
  return out;
}

async function generatePdfFromCard(cardEl) {
  const { buf, width, height } = await canvasInfoFromCard(cardEl);
  const pdfDoc = await PDFLib.PDFDocument.create();
  const page = pdfDoc.addPage([width, height]);
  const img = await pdfDoc.embedPng(buf);
  page.drawImage(img, { x: 0, y: 0, width, height });
  const pdfBytes = await pdfDoc.save();
  return new Blob([pdfBytes], { type: 'application/pdf' });
}

async function fetchArrayBuffer(url) {
  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error('http error');
    return await res.arrayBuffer();
  } catch {
    return null;
  }
}

async function generateDocxUsingTemplate(model, item, valor) {
  const path = model === 'wellington' ? 'modelo_placas/wellington.docx' : 'modelo_placas/patricia.docx';
  const ab = await fetchArrayBuffer(path);
  if (!ab) throw new Error('template docx não encontrado');
  const zip = new window.PizZip(ab);
  let xml = '';
  try { xml = zip.file('word/document.xml').asText(); } catch {}
  const hasCurly = /(\{ITEM\})|(\{VALOR\})/i.test(xml);
  const hasMustache = /(\{\{\s*ITEM\s*\}\})|(\{\{\s*VALOR\s*\}\})/i.test(xml);
  const hasBrackets = /(\[\[\s*ITEM\s*\]\])|(\[\[\s*VALOR\s*\]\])/i.test(xml);
  const hasGuillemets = /(«\s*ITEM\s*»)|(«\s*VALOR\s*»)/i.test(xml);
  const placeholdersPresent = hasCurly || hasMustache || hasBrackets || hasGuillemets;
  if (!placeholdersPresent) {
    throw new Error('template sem placeholders');
  }

  const baseOpts = { paragraphLoop: true, linebreaks: true };
  let opts = baseOpts;
  if (hasMustache) opts = { ...baseOpts, delimiters: { start: '{{', end: '}}' } };
  else if (hasBrackets) opts = { ...baseOpts, delimiters: { start: '[[', end: ']]' } };
  else if (hasGuillemets) opts = { ...baseOpts, delimiters: { start: '«', end: '»' } };

  const doc = new window.docxtemplater(zip, opts);
  doc.setData({ ITEM: String(item), VALOR: String(valor) });
  try {
    doc.render();
  } catch (e) {
    throw new Error('placeholders ausentes: use {{ITEM}} e {{VALOR}} no .docx');
  }
  const out = doc.getZip().generate({ type: 'blob' });
  return out;
}

async function generateDocxForCardOrTemplate(cardEl) {
  const item = cardEl.dataset.item;
  const model = cardEl.dataset.model;
  const valor = cardEl.dataset.valor;
  try {
    return await generateDocxUsingTemplate(model, item, valor);
  } catch {
    return await generateDocxFromCard(cardEl);
  }
}

async function fetchApiFile(apiPath, payload) {
  const res = await fetch(`${API_BASE}${apiPath}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    let msg = 'Falha na solicitação';
    try {
      const data = await res.json();
      if (data && data.error) msg = data.error;
    } catch {}
    throw new Error(msg);
  }
  return await res.blob();
}

async function onDownloadDocx(card) {
  try {
    const item = card.dataset.item;
    const model = card.dataset.model;
    const valor = card.dataset.valor;
    const blob = await fetchApiFile('/api/generate/docx', { model, item, valor });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${model}-item-${item}.docx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (e) {
    alert('Erro ao gerar DOCX: ' + (e && e.message ? e.message : e));
  }
}

async function onDownloadPdf(card) {
  try {
    const item = card.dataset.item;
    const model = card.dataset.model;
    const valor = card.dataset.valor;
    const blob = await fetchApiFile('/api/generate/pdf', { model, item, valor });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${model}-item-${item}.pdf`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (e) {
    alert('Erro ao gerar PDF: ' + (e && e.message ? e.message : e));
  }
}

async function downloadAllDocxZip() {
  const zip = new JSZip();
  const cards = $$('.cards-grid .card');
  for (const card of cards) {
    const item = card.dataset.item;
    const model = card.dataset.model;
    const valor = card.dataset.valor;
    const blob = await fetchApiFile('/api/generate/docx', { model, item, valor });
    const buf = await blob.arrayBuffer();
    zip.file(`${model}-item-${item}.docx`, buf);
  }
  const content = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(content);
  const a = document.createElement('a');
  a.href = url; a.download = 'plaquinhas-docx.zip';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function downloadAllPdfZip() {
  const zip = new JSZip();
  const cards = $$('.cards-grid .card');
  for (const card of cards) {
    const item = card.dataset.item;
    const model = card.dataset.model;
    const valor = card.dataset.valor;
    const blob = await fetchApiFile('/api/generate/pdf', { model, item, valor });
    const buf = await blob.arrayBuffer();
    zip.file(`${model}-item-${item}.pdf`, buf);
  }
  const content = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(content);
  const a = document.createElement('a');
  a.href = url; a.download = 'plaquinhas-pdf.zip';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function bindWorkspace() {
  $('#parseBtn').addEventListener('click', () => {
    const raw = $('#textoInput').value;
    const items = parseTextTable(raw);
    state.items = items;
    const hint = $('#parseInfo');
    hint.classList.remove('error');
    hint.textContent = items.length
      ? `Foram identificados ${items.length} item(ns).`
      : 'Nenhum item identificado. Confira o texto.';
    renderTablePreview(items);
    $('#generateBtn').disabled = items.length === 0;
    $('#downloadAllDocxBtn').disabled = true;
    $('#downloadAllPdfBtn').disabled = true;
    hide($('#downloadAllDocxBtn'));
    hide($('#downloadAllPdfBtn'));
    $('#cardsGrid').innerHTML = '';
    const resultArea = $('#resultArea');
    if (items.length > 0) show(resultArea); else hide(resultArea);
    const banner = $('#generationStatus');
    const bannerText = $('#generationStatusText');
    bannerText.textContent = '';
    hide(banner);
    if (state.generationTimer) { clearTimeout(state.generationTimer); state.generationTimer = null; }
  });

  $('#clearBtn').addEventListener('click', () => {
    $('#textoInput').value = '';
    $('#tablePreview').innerHTML = '';
    $('#cardsGrid').innerHTML = '';
    const hint = $('#parseInfo');
    hint.textContent = '';
    hint.classList.remove('error');
    $('#generateBtn').disabled = true;
    $('#downloadAllDocxBtn').disabled = true;
    $('#downloadAllPdfBtn').disabled = true;
    hide($('#downloadAllDocxBtn'));
    hide($('#downloadAllPdfBtn'));
    state.items = [];
    hide($('#resultArea'));
    const banner = $('#generationStatus');
    const bannerText = $('#generationStatusText');
    bannerText.textContent = '';
    hide(banner);
    if (state.generationTimer) { clearTimeout(state.generationTimer); state.generationTimer = null; }
  });

  $('#generateBtn').addEventListener('click', () => {
    const model = $('#modeloSelect').value;
    const resultArea = $('#resultArea');
    if (resultArea) show(resultArea);
    if (!model) {
      const hint = $('#parseInfo');
      hint.textContent = 'Selecione um modelo da plaquinha e clique novamente em "Gerar plaquinhas".';
      hint.classList.add('error');
      try { hint.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); } catch {}
      setGenerationStatus('Falha ao gerar: selecione um modelo da plaquinha.', true);
      return;
    }
    $('#parseInfo').classList.remove('error');
    try {
      renderCards(state.items);
      if ((state.items || []).length > 0) {
        setGenerationStatus('Plaquinhas geradas com sucesso.', false);
      } else {
        setGenerationStatus('Falha ao gerar: nenhum item encontrado.', true);
      }
    } catch {
      setGenerationStatus('Falha ao gerar as plaquinhas.', true);
    }
  });

  $('#modeloSelect').addEventListener('change', () => {
    if (state.cardsRendered) {
      renderCards(state.items);
    } else {
      const items = state.items || [];
      $('#generateBtn').disabled = items.length === 0;
    }
  });

  document.body.addEventListener('click', (e) => {
    const btnDocx = e.target.closest('.download-docx');
    const btnPdf = e.target.closest('.download-pdf');
    if (btnDocx) {
      const card = e.target.closest('.card');
      onDownloadDocx(card);
      return;
    }
    if (btnPdf) {
      const card = e.target.closest('.card');
      onDownloadPdf(card);
      return;
    }
  });

  $('#downloadAllDocxBtn').addEventListener('click', downloadAllDocxZip);
  $('#downloadAllPdfBtn').addEventListener('click', downloadAllPdfZip);
}

function boot() {
  bindAuth();
  bindWorkspace();
  bootSession();
  const sel = $('#modeloSelect');
  if (sel) sel.value = '';
  const items = state.items || [];
  $('#generateBtn').disabled = items.length === 0;
  hide($('#downloadAllDocxBtn'));
  hide($('#downloadAllPdfBtn'));
  const banner = $('#generationStatus');
  const bannerText = $('#generationStatusText');
  if (bannerText) bannerText.textContent = '';
  if (banner) hide(banner);
  (async () => {
    const el = $('#pwLogo');
    if (!el) return;
    const candidates = [
      'logo/logo_two.png',
      'LOGO/logo_two.png',
      'logo/pw.png','logo/pw.jpg','logo/pw.svg',
      'logo/logo.png','logo/logo.jpg','logo/logo.svg',
      'LOGO/pw.png','LOGO/pw.jpg','LOGO/pw.svg',
      'LOGO/logo.png','LOGO/logo.jpg','LOGO/logo.svg'
    ];
    for (const u of candidates) {
      try {
        const r = await fetch(u, { cache: 'no-store' });
        if (r.ok) {
          el.src = u;
          el.classList.remove('hidden');
          const fb = document.querySelector('.logo');
          if (fb) fb.classList.add('hidden');
          break;
        }
      } catch {}
    }
    if (el.classList.contains('hidden')) {
      const folders = ['logo/','LOGO/'];
      const exts = ['png','jpg','jpeg','svg','webp'];
      for (const dir of folders) {
        try {
          const res = await fetch(dir, { cache: 'no-store' });
          if (!res.ok) continue;
          const html = await res.text();
          const matches = [...html.matchAll(/href="([^"]+\.(?:png|jpg|jpeg|svg|webp))"/gi)];
          const found = matches.map(m => m[1]).find(h => exts.some(e => h.toLowerCase().endsWith('.'+e)));
          if (found) {
            const url = dir + found.replace(/^\.?\//,'');
            el.src = url;
            el.classList.remove('hidden');
            const fb = document.querySelector('.logo');
            if (fb) fb.classList.add('hidden');
            break;
          }
        } catch {}
      }
    }
  })();
}

boot();
