const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// --------- Helpers de CLI ----------
function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith('--')) {
      const [k, v] = a.split('=');
      args[k.replace(/^--/, '')] = v === undefined ? true : v;
    }
  }
  return args;
}

const args = parseArgs(process.argv);

// --------- Config ----------
const EXCEL_FILE = args.file || 'SIGEWHA (13).xlsx';
const OUT_FILE = args.out || 'contactos.vcf';

// Prioridad de hoja: --sheet-name, luego --sheet-index (0-based), sino primera hoja
const SHEET_NAME = args['sheet-name'];
const SHEET_INDEX = args['sheet-index'] !== undefined ? Number(args['sheet-index']) : 0;

// Columnas: si no se pasan, intentamos autodetectar por encabezados; si tampoco, por defecto tel=col 2 (idx 1), nombre=col 3 (idx 2)
let PHONE_COL = args['phone-col'] !== undefined ? Number(args['phone-col']) : undefined; // 0-based
let NAME_COL = args['name-col'] !== undefined ? Number(args['name-col']) : undefined;

// --has-header=true/false para forzar; si no viene, se intenta detectar
const FORCE_HAS_HEADER = args['has-header'] !== undefined ? String(args['has-header']).toLowerCase() === 'true' : undefined;

// Prefijo país opcional, p.ej. --dial-code=+34
const DIAL_CODE = args['dial-code'] || '';

// --------- Utilidades ----------
function normalizePhone(raw) {
  if (raw === null || raw === undefined) return '';
  let s = String(raw).trim();
  if (!s) return '';
  // Si contiene múltiples, tomamos el primero separando por comas / ; / /
  s = s.split(/[;,/]| y | o /i)[0].trim();

  const hadPlus = s.startsWith('+');
  const digits = s.replace(/[^\d]/g, '');
  if (!digits) return '';

  if (hadPlus) return `+${digits}`;
  if (DIAL_CODE) {
    const dc = String(DIAL_CODE).trim();
    const dcDigits = dc.replace(/[^\d]/g, '');
    const dcPref = dc.startsWith('+') ? `+${dcDigits}` : `+${dcDigits}`;
    // Evitar doble prefijo si el número ya lo incluye al inicio
    if (digits.startsWith(dcDigits)) return `${dcPref}${digits.slice(dcDigits.length)}`;
    return `${dcPref}${digits}`;
  }
  return digits;
}

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

function detectColumnsFromHeader(headerRow) {
  const phoneCandidates = ['tel', 'telefono', 'teléfono', 'phone', 'mobile', 'cel', 'celular', 'whatsapp', 'móvil'];
  const nameCandidates = ['nombre', 'name', 'contacto', 'fullname', 'full name', 'fn'];

  const norm = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  let phoneIdx, nameIdx;
  headerRow.forEach((h, idx) => {
    const n = norm(h);
    if (phoneIdx === undefined && phoneCandidates.some(k => n.includes(k))) phoneIdx = idx;
    if (nameIdx === undefined && nameCandidates.some(k => n.includes(k))) nameIdx = idx;
  });
  return { phoneIdx, nameIdx };
}

function makeVCard(name, phone) {
  const safeName = String(name || '').trim();
  const safePhone = normalizePhone(phone);
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

// --------- Main ----------
(function main() {
  const filePath = path.resolve(EXCEL_FILE);
  if (!fs.existsSync(filePath)) {
    console.error(`No se encuentra el archivo: ${filePath}`);
    process.exit(1);
  }

  const wb = XLSX.readFile(filePath);
  let ws;
  if (SHEET_NAME && wb.Sheets[SHEET_NAME]) {
    ws = wb.Sheets[SHEET_NAME];
  } else {
    const sheetNames = wb.SheetNames || [];
    const idx = SHEET_NAME ? sheetNames.indexOf(SHEET_NAME) : SHEET_INDEX;
    const useName = sheetNames[idx] || sheetNames[0];
    ws = wb.Sheets[useName];
  }
  if (!ws) {
    console.error('No se pudo abrir la hoja deseada.');
    process.exit(1);
  }

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (!rows.length) {
    console.error('La hoja está vacía.');
    process.exit(1);
  }

  // Detectar header y columnas si no se especificaron
  let startRow = 0;
  const firstRow = rows[0];

  let hasHeader;
  if (FORCE_HAS_HEADER !== undefined) {
    hasHeader = FORCE_HAS_HEADER;
  } else {
    // Si no se forzó: heurística
    hasHeader = isHeaderRow(firstRow, PHONE_COL, NAME_COL);
  }

  if (hasHeader) {
    if (PHONE_COL === undefined || NAME_COL === undefined) {
      const det = detectColumnsFromHeader(firstRow);
      if (PHONE_COL === undefined) PHONE_COL = det.phoneIdx;
      if (NAME_COL === undefined) NAME_COL = det.nameIdx;
    }
    startRow = 1;
  }

  // Si aún no tenemos columnas, usar por defecto: tel=col 2 (idx 1), nombre=col 3 (idx 2)
  if (PHONE_COL === undefined) PHONE_COL = 1;
  if (NAME_COL === undefined) NAME_COL = 2;

  let created = 0;
  const out = fs.createWriteStream(OUT_FILE, { encoding: 'utf8' });

  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i] || [];
    const phone = row[PHONE_COL];
    const name = row[NAME_COL];
    const card = makeVCard(name, phone);
    if (card) {
      out.write(card);
      created++;
    }
  }
  out.end();

  if (!created) {
    console.error('No se generaron contactos. Revisa columnas de teléfono y nombre.');
    process.exit(1);
  }

  console.log(`Generado: ${OUT_FILE} con ${created} contactos.`);
})();