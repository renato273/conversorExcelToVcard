const express = require('express');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

const UPLOADS_DIR = path.join(__dirname, 'uploads');
const EXPORTS_DIR = path.join(__dirname, 'exports');
const PUBLIC_DIR = path.join(__dirname, 'public');

for (const dir of [UPLOADS_DIR, EXPORTS_DIR, PUBLIC_DIR]) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(PUBLIC_DIR));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    const safe = file.originalname.replace(/[^\w\-. ()\[\]]/g, '_');
    cb(null, `${Date.now()}_${safe}`);
  }
});
const upload = multer({ storage });

function normalizePhone(raw, dialCode) {
  if (raw === null || raw === undefined) return '';
  let s = String(raw).trim();
  if (!s) return '';
  s = s.split(/[;,/]| y | o /i)[0].trim();
  const hadPlus = s.startsWith('+');
  const digits = s.replace(/[^\d]/g, '');
  if (!digits) return '';
  if (hadPlus) return `+${digits}`;
  if (dialCode) {
    const dcDigits = String(dialCode).replace(/[^\d]/g, '');
    const dcPref = `+${dcDigits}`;
    if (digits.startsWith(dcDigits)) return `${dcPref}${digits.slice(dcDigits.length)}`;
    return `${dcPref}${digits}`;
  }
  return digits;
}

function makeVCard(name, phone) {
  const safeName = String(name || '').trim();
  const safePhone = String(phone || '').trim();
  if (!safeName && !safePhone) return '';
  const lines = [
    'BEGIN:VCARD',
    'VERSION:3.0',
    `N:;${safeName};;;`,
    `FN:${safeName}`,
    safePhone ? `TEL;TYPE=CELL:${safePhone}` : null,
    'END:VCARD'
  ].filter(Boolean);
  return lines.join('\r\n') + '\r\n';
}

app.get('/api/files', (req, res) => {
  const entries = fs.readdirSync(UPLOADS_DIR)
    .filter(f => /\.(xlsx|xls|csv)$/i.test(f))
    .map(f => ({ name: f, size: fs.statSync(path.join(UPLOADS_DIR, f)).size }));
  res.json(entries);
});

app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  res.json({ ok: true, file: req.file.filename });
});

app.delete('/api/files/:name', (req, res) => {
  const name = req.params.name;
  const filePath = path.join(UPLOADS_DIR, name);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'No existe' });
  fs.unlinkSync(filePath);
  res.json({ ok: true });
});

app.post('/api/generate', (req, res) => {
  const { file, sheetIndex = 0, phoneCol, nameCol, hasHeader, dialCode } = req.body || {};
  if (!file) return res.status(400).json({ error: 'Parametro file requerido' });
  const filePath = path.join(UPLOADS_DIR, file);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Archivo no encontrado' });

  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames[Number(sheetIndex) || 0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (!rows.length) return res.status(400).json({ error: 'Hoja vacia' });

  let start = 0;
  let pCol = phoneCol !== undefined ? Number(phoneCol) : undefined;
  let nCol = nameCol !== undefined ? Number(nameCol) : undefined;

  function looksLikePhone(s) {
    if (!s) return false;
    const t = String(s).trim();
    return /^[+]?[\d\s().-]+$/.test(t) && /\d/.test(t);
  }
  function looksLikeName(s) {
    if (!s) return false;
    return /[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]/.test(String(s));
  }
  function isHeaderRow(row, phoneIdx, nameIdx) {
    const phone = String(row[phoneIdx ?? 1] ?? '').trim();
    const name = String(row[nameIdx ?? 2] ?? '').trim();
    return !looksLikePhone(phone) && looksLikeName(name);
  }
  if (hasHeader === true || hasHeader === 'true') {
    start = 1;
  } else if (hasHeader === false || hasHeader === 'false') {
    start = 0;
  } else {
    start = isHeaderRow(rows[0], pCol, nCol) ? 1 : 0;
  }
  if (pCol === undefined) pCol = 1;
  if (nCol === undefined) nCol = 2;

  const outName = `${path.parse(file).name}.vcf`;
  const outPath = path.join(EXPORTS_DIR, outName);
  const out = fs.createWriteStream(outPath, { encoding: 'utf8' });
  let count = 0;
  for (let i = start; i < rows.length; i++) {
    const r = rows[i] || [];
    const phoneRaw = r[pCol];
    const name = r[nCol];
    const phone = normalizePhone(phoneRaw, dialCode);
    const card = makeVCard(name, phone);
    if (card) {
      out.write(card);
      count++;
    }
  }
  out.end(() => {
    if (!count) return res.status(400).json({ error: 'Sin contactos generados' });
    res.json({ ok: true, file: `/exports/${outName}`, count });
  });
});

app.use('/exports', express.static(EXPORTS_DIR));

// Página principal y fallback para rutas no-API (compatible con Express 5)
app.get('/', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'index.html'));
});
app.use((req, res, next) => {
  if (req.method === 'GET' && !req.path.startsWith('/api') && !req.path.startsWith('/exports')) {
    return res.sendFile(path.join(PUBLIC_DIR, 'index.html'));
  }
  return next();
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en http://localhost:${PORT}`);
});


