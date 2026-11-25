// Lõi xử lý dùng chung cho trình duyệt và Node (test):
// - đọc cấu trúc mẫu .docx thành bản mô tả có ID cho AI
// - đọc file dữ liệu Excel/CSV thành văn bản
// - áp kế hoạch (plan) AI trả về lên mẫu để sinh từng file

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS = 'http://www.w3.org/XML/1998/namespace';

// ---------- Số học an toàn + định dạng số ----------

export function evalExpr(src) {
  const s = String(src).replace(/\s+/g, '');
  let i = 0;
  const fail = () => { throw new Error(`Biểu thức không hợp lệ: "${src}"`); };
  function factor() {
    if (s[i] === '-') { i++; return -factor(); }
    if (s[i] === '+') { i++; return factor(); }
    if (s[i] === '(') {
      i++;
      const v = expr();
      if (s[i] !== ')') fail();
      i++;
      return v;
    }
    const m = /^\d+(\.\d+)?/.exec(s.slice(i));
    if (!m) fail();
    i += m[0].length;
    return Number(m[0]);
  }
  function term() {
    let v = factor();
    while (s[i] === '*' || s[i] === '/') {
      const op = s[i++];
      const r = factor();
      v = op === '*' ? v * r : v / r;
    }
    return v;
  }
  function expr() {
    let v = term();
    while (s[i] === '+' || s[i] === '-') {
      const op = s[i++];
      const r = term();
      v = op === '+' ? v + r : v - r;
    }
    return v;
  }
  if (!s) fail();
  const v = expr();
  if (i !== s.length || !Number.isFinite(v)) fail();
  return v;
}

export function formatNumber(n) {
  const rounded = Math.round(n * 100) / 100;
  const neg = rounded < 0;
  const [int, dec] = Math.abs(rounded).toString().split('.');
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (neg ? '-' : '') + grouped + (dec ? ',' + dec : '');
}

const DIGITS = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];

function readTriple(n, full) {
  const tr = Math.floor(n / 100);
  const ch = Math.floor(n / 10) % 10;
  const dv = n % 10;
  const out = [];
  if (tr || full) out.push(DIGITS[tr], 'trăm');
  if (ch === 0) {
    if (dv && (tr || full)) out.push('lẻ');
  } else if (ch === 1) out.push('mười');
  else out.push(DIGITS[ch], 'mươi');
  if (dv) {
    if (dv === 1 && ch > 1) out.push('mốt');
    else if (dv === 5 && ch >= 1) out.push('lăm');
    else if (dv === 4 && ch > 1) out.push('tư');
    else out.push(DIGITS[dv]);
  }
  return out.join(' ');
}

// Đọc số nguyên thành chữ tiếng Việt (chữ thường, không kèm đơn vị).
export function readVietnamese(value) {
  let n = Math.round(value);
  if (n === 0) return 'không';
  const neg = n < 0;
  n = Math.abs(n);
  const units = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ', 'triệu tỷ'];
  const groups = [];
  while (n > 0) { groups.push(n % 1000); n = Math.floor(n / 1000); }
  if (groups.length > units.length) throw new Error('Số quá lớn để đọc thành chữ');
  const parts = [];
  for (let g = groups.length - 1; g >= 0; g--) {
    if (!groups[g]) continue;
    parts.push((readTriple(groups[g], g !== groups.length - 1) + ' ' + units[g]).trim());
  }
  return (neg ? 'âm ' : '') + parts.join(' ');
}

// Thay {{num:...}}, {{words:...}}, {{Words:...}} trong giá trị AI trả về.
export function renderValue(text, warnings = []) {
  return String(text ?? '').replace(/\{\{\s*(num|words|Words)\s*:\s*([^{}]+?)\s*\}\}/g, (all, fn, ex) => {
    try {
      const v = evalExpr(ex);
      if (fn === 'num') return formatNumber(v);
      const w = readVietnamese(v);
      return fn === 'Words' ? w[0].toUpperCase() + w.slice(1) : w;
    } catch (e) {
      warnings.push(`${e.message}`);
      return all;
    }
  });
}

// ---------- Đọc file dữ liệu ----------

// XLSX: thư viện SheetJS (truyền vào để dùng được cả ở trình duyệt lẫn Node).
export function workbookToText(XLSX, wb) {
  const out = [];
  let cellCount = 0;
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) continue;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const merges = (ws['!merges'] || []).map((m) => XLSX.utils.encode_range(m));
    out.push(`## Sheet "${name}" (vùng ${ws['!ref']}${merges.length ? `; ô gộp: ${merges.join(', ')}` : ''})`);
    for (let r = range.s.r; r <= range.e.r; r++) {
      const cells = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell) continue;
        let v;
        if (cell.v === undefined || cell.v === null || cell.v === '') {
          if (!cell.f) continue;
          v = `=${cell.f} (chưa có giá trị tính sẵn)`;
        } else if (cell.t === 'n') {
          const isDate = cell.z && XLSX.SSF && XLSX.SSF.is_date && XLSX.SSF.is_date(cell.z);
          v = isDate && cell.w ? cell.w : String(Math.round(cell.v * 1e6) / 1e6);
        } else if (cell.t === 'b') v = cell.v ? 'TRUE' : 'FALSE';
        else if (cell.t === 'e') v = cell.w || '#ERROR';
        else v = String(cell.v);
        v = v.replace(/\s+/g, ' ').trim();
        if (!v) continue;
        cells.push(`${XLSX.utils.encode_col(c)}=${v}`);
        cellCount++;
      }
      if (cells.length) out.push(`R${r + 1}: ${cells.join(' | ')}`);
    }
    out.push('');
  }
  return { text: out.join('\n').trim(), cellCount, sheets: wb.SheetNames.length };
}

// Đoán dòng tiêu đề của sheet đầu tiên để gợi ý các trường.
export function guessHeaders(XLSX, wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws || !ws['!ref']) return [];
  const range = XLSX.utils.decode_range(ws['!ref']);
  let best = [];
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 15); r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.t === 's' && String(cell.v).trim()) row.push(String(cell.v).replace(/\s+/g, ' ').trim());
    }
    if (row.length > best.length) best = row;
  }
  return best.length >= 2 ? [...new Set(best)] : [];
}

// ---------- Mẫu .docx ----------

function kids(node, name) {
  const out = [];
  for (let c = node.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && (!name || c.nodeName === name)) out.push(c);
  }
  return out;
}

function allText(node) {
  const ts = node.getElementsByTagName('w:t');
  let s = '';
  for (let i = 0; i < ts.length; i++) s += ts[i].textContent || '';
  return s;
}

function indexRow(map, tr, rid) {
  map.set(rid, tr);
  kids(tr, 'w:tc').forEach((tc, ci) => {
    map.set(`${rid}.c${ci}`, tc);
    kids(tc, 'w:p').forEach((p, pi) => map.set(`${rid}.c${ci}.p${pi}`, p));
  });
}

function indexDocument(doc) {
  const body = doc.getElementsByTagName('w:body')[0];
  if (!body) throw new Error('File mẫu không phải tài liệu Word hợp lệ (thiếu w:body).');
  const map = new Map();
  let pi = 0;
  let ti = 0;
  for (const el of kids(body)) {
    if (el.nodeName === 'w:p') map.set(`p${pi++}`, el);
    else if (el.nodeName === 'w:tbl') {
      const tid = `t${ti++}`;
      map.set(tid, el);
      kids(el, 'w:tr').forEach((tr, ri) => indexRow(map, tr, `${tid}.r${ri}`));
    }
  }
  return map;
}

function paragraphFlags(p) {
  const pPr = kids(p, 'w:pPr')[0];
  const jc = pPr && kids(pPr, 'w:jc')[0];
  const val = jc && jc.getAttribute('w:val');
  return val && val !== 'left' && val !== 'both' ? ` [${val}]` : '';
}

// Bản mô tả mẫu gửi cho AI: mỗi đoạn văn / dòng bảng / ô có một ID.
export function describeTemplate(xml, dom) {
  const doc = new dom.DOMParser().parseFromString(xml, 'application/xml');
  const map = indexDocument(doc);
  const lines = [];
  let paragraphs = 0;
  let tables = 0;
  for (const [id, el] of map) {
    const depth = id.split('.').length;
    const pad = '  '.repeat(depth - 1);
    if (el.nodeName === 'w:p') {
      paragraphs++;
      lines.push(`${pad}${id}: ${JSON.stringify(allText(el))}${paragraphFlags(el)}`);
    } else if (el.nodeName === 'w:tbl') {
      tables++;
      lines.push(`${id} (bảng, ${kids(el, 'w:tr').length} dòng)`);
    } else if (el.nodeName === 'w:tr') {
      lines.push(`${pad}${id} (dòng, ${kids(el, 'w:tc').length} ô)`);
    } else if (el.nodeName === 'w:tc') {
      const tcPr = kids(el, 'w:tcPr')[0];
      const span = tcPr && kids(tcPr, 'w:gridSpan')[0];
      const n = span && span.getAttribute('w:val');
      if (n && n !== '1') lines.push(`${pad}${id} (ô gộp ${n} cột)`);
    }
  }
  return { text: lines.join('\n'), paragraphs, tables, ids: [...map.keys()] };
}

function runText(r) {
  return kids(r, 'w:t').map((t) => t.textContent || '').join('');
}

function setRunText(doc, r, text) {
  for (const t of kids(r, 'w:t')) r.removeChild(t);
  if (!text) return;
  text.split('\n').forEach((part, i) => {
    if (i > 0) r.appendChild(doc.createElementNS(W, 'w:br'));
    const t = doc.createElementNS(W, 'w:t');
    t.setAttributeNS(XMLNS, 'xml:space', 'preserve');
    t.appendChild(doc.createTextNode(part));
    r.appendChild(t);
  });
}

const PPR_AFTER_JC = ['w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'];

function setAlign(doc, p, align) {
  let pPr = kids(p, 'w:pPr')[0];
  if (!pPr) {
    pPr = doc.createElementNS(W, 'w:pPr');
    p.insertBefore(pPr, p.firstChild);
  }
  for (const jc of kids(pPr, 'w:jc')) pPr.removeChild(jc);
  const jc = doc.createElementNS(W, 'w:jc');
  jc.setAttributeNS(W, 'w:val', align);
  const before = kids(pPr).find((k) => PPR_AFTER_JC.includes(k.nodeName));
  pPr.insertBefore(jc, before || null);
}

function setBold(doc, r) {
  let rPr = kids(r, 'w:rPr')[0];
  if (!rPr) {
    rPr = doc.createElementNS(W, 'w:rPr');
    r.insertBefore(rPr, r.firstChild);
  }
  if (kids(rPr, 'w:b').length) return;
  const b = doc.createElementNS(W, 'w:b');
  const lead = kids(rPr).filter((k) => k.nodeName === 'w:rStyle' || k.nodeName === 'w:rFonts');
  const after = lead.length ? lead[lead.length - 1].nextSibling : rPr.firstChild;
  rPr.insertBefore(b, after);
}

const norm = (s) => s.replace(/ /g, ' ');

// Đặt lại toàn bộ nội dung chữ của đoạn văn nhưng giữ định dạng:
// phần đầu trùng với nội dung cũ (nhãn) được giữ nguyên run,
// phần khác biệt được ghi vào run tại điểm bắt đầu khác.
export function setParagraphText(doc, p, newText, opts = {}) {
  let runs = kids(p, 'w:r');
  if (!runs.length) {
    const r = doc.createElementNS(W, 'w:r');
    const pPr = kids(p, 'w:pPr')[0];
    const rPr = pPr && kids(pPr, 'w:rPr')[0];
    if (rPr) r.appendChild(rPr.cloneNode(true));
    p.appendChild(r);
    runs = [r];
  }
  const texts = runs.map(runText);
  const oldText = texts.join('');
  const a = norm(oldText);
  const b = norm(newText);
  let k = 0;
  while (k < a.length && k < b.length && a[k] === b[k]) k++;

  let target = -1;
  let offset = 0;
  if (k >= oldText.length) {
    for (let i = runs.length - 1; i >= 0; i--) if (texts[i]) { target = i; break; }
    if (target === -1) target = 0;
    offset = texts[target].length;
  } else {
    let start = 0;
    for (let i = 0; i < runs.length; i++) {
      if (texts[i] && k < start + texts[i].length) { target = i; offset = k - start; break; }
      start += texts[i].length;
    }
  }
  setRunText(doc, runs[target], texts[target].slice(0, offset) + newText.slice(k));
  for (let i = target + 1; i < runs.length; i++) if (texts[i]) setRunText(doc, runs[i], '');
  if (opts.bold) setBold(doc, runs[target]);
  if (opts.align && opts.align !== 'keep') setAlign(doc, p, opts.align);
}

function stripIds(node) {
  if (node.nodeType !== 1) return;
  if (node.hasAttribute('w14:paraId')) node.removeAttribute('w14:paraId');
  if (node.hasAttribute('w14:textId')) node.removeAttribute('w14:textId');
  for (let c = node.firstChild; c; c = c.nextSibling) stripIds(c);
}

function applySetup(doc, map, ops, warnings) {
  for (const op of ops || []) {
    const target = map.get(op.target);
    if (!target) { warnings.push(`Bước chuẩn bị: không tìm thấy "${op.target}" trong mẫu.`); continue; }
    if (!op.as || map.has(op.as)) { warnings.push(`Bước chuẩn bị: ID mới "${op.as}" trống hoặc bị trùng.`); continue; }
    if (op.op === 'clone_row' && target.nodeName === 'w:tr') {
      const copy = target.cloneNode(true);
      stripIds(copy);
      target.parentNode.insertBefore(copy, target.nextSibling);
      indexRow(map, copy, op.as);
    } else if (op.op === 'insert_paragraph' && target.nodeName === 'w:p') {
      const src = map.get(op.copy_from) || target;
      if (src.nodeName !== 'w:p') { warnings.push(`Bước chuẩn bị: "${op.copy_from}" không phải đoạn văn.`); continue; }
      const copy = src.cloneNode(true);
      stripIds(copy);
      for (const r of kids(copy, 'w:r')) setRunText(doc, r, '');
      target.parentNode.insertBefore(copy, target.nextSibling);
      map.set(op.as, copy);
    } else {
      warnings.push(`Bước chuẩn bị "${op.op}" không áp dụng được cho "${op.target}".`);
    }
  }
}

// Sinh document.xml cho một file đầu ra.
export function buildDocumentXml(xml, plan, docSpec, dom) {
  const warnings = [];
  const doc = new dom.DOMParser().parseFromString(xml, 'application/xml');
  const map = indexDocument(doc);
  applySetup(doc, map, plan.setup, warnings);

  const values = new Map((docSpec.values || []).map((v) => [v.key, v.value]));
  const known = new Set((plan.slots || []).map((s) => s.key));
  for (const key of values.keys()) if (!known.has(key)) warnings.push(`Giá trị "${key}" không khớp ô điền nào.`);
  const removed = new Set(docSpec.remove || []);
  const isRemoved = (id) => [...removed].some((r) => id === r || id.startsWith(r + '.'));

  for (const slot of plan.slots || []) {
    if (!values.has(slot.key) || isRemoved(slot.target)) continue;
    const p = map.get(slot.target);
    if (!p || p.nodeName !== 'w:p') { warnings.push(`Ô điền "${slot.key}": "${slot.target}" không phải đoạn văn trong mẫu.`); continue; }
    const value = renderValue(values.get(slot.key), warnings);
    const tpl = slot.template && slot.template.includes('{value}') ? slot.template : '{value}';
    setParagraphText(doc, p, tpl.split('{value}').join(value), { align: slot.align, bold: slot.bold });
  }

  for (const id of removed) {
    const el = map.get(id);
    if (!el) { warnings.push(`Không tìm thấy "${id}" để xoá.`); continue; }
    if (el.nodeName === 'w:tc') { warnings.push(`Không xoá riêng một ô ("${id}"), chỉ xoá dòng hoặc đoạn văn.`); continue; }
    const parent = el.parentNode;
    if (el.nodeName === 'w:p' && parent.nodeName === 'w:tc' && kids(parent, 'w:p').length === 1) {
      for (const r of kids(el, 'w:r')) setRunText(doc, r, '');
      continue;
    }
    parent.removeChild(el);
  }

  let out = new dom.XMLSerializer().serializeToString(doc);
  out = out.replace(/^\s*<\?xml[^>]*\?>\s*/, '');
  return { xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out, warnings };
}

export function safeFilename(name, used) {
  let base = String(name || 'file').replace(/\.docx$/i, '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 150) || 'file';
  let candidate = base;
  let n = 2;
  while (used.has(candidate.toLowerCase())) candidate = `${base} (${n++})`;
  used.add(candidate.toLowerCase());
  return candidate + '.docx';
}

// Sinh toàn bộ file .docx từ mẫu (JSZip truyền vào).
export async function generateDocuments({ JSZip, templateBytes, plan, dom, onProgress }) {
  const zip = await JSZip.loadAsync(templateBytes);
  const main = zip.file('word/document.xml');
  if (!main) throw new Error('File mẫu không phải .docx hợp lệ (thiếu word/document.xml).');
  const xml = await main.async('string');
  const others = [];
  for (const [path, entry] of Object.entries(zip.files)) {
    if (entry.dir || path === 'word/document.xml') continue;
    others.push([path, await entry.async('uint8array')]);
  }
  const used = new Set();
  const files = [];
  const docs = plan.documents || [];
  for (let i = 0; i < docs.length; i++) {
    const built = buildDocumentXml(xml, plan, docs[i], dom);
    const out = new JSZip();
    for (const [path, bytes] of others) out.file(path, bytes, { createFolders: false });
    out.file('word/document.xml', built.xml, { createFolders: false });
    const bytes = await out.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    files.push({ name: safeFilename(docs[i].filename, used), bytes, warnings: built.warnings });
    if (onProgress) onProgress(i + 1, docs.length);
  }
  return files;
}
