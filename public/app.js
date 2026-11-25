import { describeTemplate, workbookToText, guessHeaders, generateDocuments, renderValue } from './core.js';

const $ = (id) => document.getElementById(id);
const dom = { DOMParser, XMLSerializer };
const state = { templateBytes: null, templateText: '', templateName: '', dataText: '', dataName: '', files: [], urls: [] };

function setInfo(el, text, cls) {
  el.textContent = text;
  el.className = 'hint' + (cls ? ' ' + cls : '');
}

function setStatus(text, isError) {
  $('status').textContent = text;
  $('status').className = 'status' + (isError ? ' error' : '');
}

function remember(key, value) {
  try {
    if (value === undefined) return localStorage.getItem(key) || '';
    localStorage.setItem(key, value);
  } catch { /* trình duyệt chặn lưu trữ: bỏ qua */ }
  return '';
}

async function init() {
  try {
    const cfg = await (await fetch('/api/config')).json();
    if (cfg.passwordRequired) {
      $('access').hidden = false;
      $('password').value = remember('accessCode');
    }
    if (!cfg.configured) setStatus('Máy chủ chưa cấu hình API key Claude.', true);
  } catch { /* vẫn cho dùng, lỗi sẽ hiện khi gửi */ }
}

$('template').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  state.templateBytes = null;
  if (!file) return;
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const zip = await JSZip.loadAsync(bytes);
    const main = zip.file('word/document.xml');
    if (!main) throw new Error('Đây không phải file .docx hợp lệ.');
    const info = describeTemplate(await main.async('string'), dom);
    state.templateBytes = bytes;
    state.templateText = info.text;
    state.templateName = file.name;
    setInfo($('templateInfo'), `Đã đọc: ${info.paragraphs} đoạn văn, ${info.tables} bảng.`, 'ok');
  } catch (err) {
    setInfo($('templateInfo'), `Không đọc được mẫu: ${err.message}`, 'bad');
  }
});

$('data').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  state.dataText = '';
  $('chips').replaceChildren();
  if (!file) return;
  try {
    const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array', cellNF: true });
    const info = workbookToText(XLSX, wb);
    if (!info.cellCount) throw new Error('File không có dữ liệu.');
    state.dataText = info.text;
    state.dataName = file.name;
    setInfo($('dataInfo'), `Đã đọc: ${info.sheets} sheet, ${info.cellCount} ô có dữ liệu.`, 'ok');
    for (const header of guessHeaders(XLSX, wb).slice(0, 24)) {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'chip';
      chip.textContent = header;
      chip.setAttribute('aria-pressed', 'false');
      chip.addEventListener('click', () => toggleField(chip, header));
      $('chips').appendChild(chip);
    }
  } catch (err) {
    setInfo($('dataInfo'), `Không đọc được dữ liệu: ${err.message}`, 'bad');
  }
});

function toggleField(chip, header) {
  const list = $('fields').value.split(',').map((s) => s.trim()).filter(Boolean);
  const at = list.indexOf(header);
  if (at >= 0) list.splice(at, 1);
  else list.push(header);
  $('fields').value = list.join(', ');
  chip.setAttribute('aria-pressed', String(at < 0));
}

async function requestPlan(payload, password) {
  const res = await fetch('/api/generate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-app-password': password },
    body: JSON.stringify(payload),
  });
  if (!res.ok || !res.body) {
    let message = `Máy chủ trả lỗi ${res.status}.`;
    try { message = (await res.json()).error || message; } catch { /* giữ thông báo mặc định */ }
    throw new Error(message);
  }
  const started = Date.now();
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let json = '';
  let finished = false;
  const tick = () => {
    const secs = Math.round((Date.now() - started) / 1000);
    setStatus(json ? `AI đang viết kế hoạch điền mẫu... ${secs}s` : `AI đang đọc mẫu và dữ liệu... ${secs}s`);
  };
  const timer = setInterval(tick, 1000);
  tick();
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      let nl;
      while ((nl = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, nl).trim();
        buffer = buffer.slice(nl + 1);
        if (!line) continue;
        const event = JSON.parse(line);
        if (event.type === 'delta') json += event.text;
        else if (event.type === 'error') throw new Error(event.message);
        else if (event.type === 'done') finished = true;
      }
    }
  } finally {
    clearInterval(timer);
  }
  if (!finished) throw new Error('Kết nối bị ngắt trước khi AI trả xong kết quả (có thể do quá thời gian xử lý). Hãy thử lại hoặc tách nhỏ file dữ liệu.');
  try {
    return JSON.parse(json);
  } catch {
    throw new Error('AI trả về kết quả không đọc được. Hãy thử lại.');
  }
}

function showResult(plan, files, payload) {
  for (const url of state.urls) URL.revokeObjectURL(url);
  state.urls = [];
  state.files = files;
  const docs = plan.documents || [];

  const summary = $('summary');
  summary.replaceChildren();
  const line = document.createElement('p');
  const count = document.createElement('strong');
  count.textContent = `${files.length} file`;
  line.append('Đã tạo ', count, '.');
  if (payload.expectedFiles && Number(payload.expectedFiles) !== files.length) {
    line.append(` Anh/chị dự kiến ${payload.expectedFiles} file - xem ghi chú bên dưới để biết lý do lệch.`);
  }
  summary.appendChild(line);

  const notes = $('notes');
  notes.hidden = !plan.notes;
  notes.replaceChildren();
  if (plan.notes) {
    const h = document.createElement('h3');
    h.textContent = 'Ghi chú của AI - nên đọc trước khi dùng file';
    notes.append(h, plan.notes);
  }

  const keys = (plan.slots || []).map((s) => s.key);
  const hasWarn = files.some((f) => f.warnings.length);
  const table = $('preview');
  table.replaceChildren();
  const head = table.insertRow();
  for (const title of ['#', 'File', ...keys, ...(hasWarn ? ['Cảnh báo'] : [])]) {
    const th = document.createElement('th');
    th.textContent = title;
    head.appendChild(th);
  }
  files.forEach((file, i) => {
    const row = table.insertRow();
    row.insertCell().textContent = String(i + 1);
    const url = URL.createObjectURL(new Blob([file.bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    state.urls.push(url);
    const link = document.createElement('a');
    link.href = url;
    link.download = file.name;
    link.textContent = file.name;
    row.insertCell().appendChild(link);
    const values = new Map((docs[i].values || []).map((v) => [v.key, v.value]));
    for (const key of keys) {
      const cell = row.insertCell();
      const text = values.has(key) ? renderValue(values.get(key)) : '';
      cell.textContent = text;
      cell.title = text;
    }
    if (hasWarn) {
      const cell = row.insertCell();
      cell.className = 'warn';
      cell.textContent = file.warnings.join(' ');
    }
  });
  $('result').hidden = false;
  $('result').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

$('form').addEventListener('submit', async (e) => {
  e.preventDefault();
  if (!state.templateBytes) return setStatus('Chưa chọn mẫu Word hợp lệ.', true);
  if (!state.dataText) return setStatus('Chưa chọn file dữ liệu hợp lệ.', true);
  const prompt = $('prompt').value.trim();
  if (!prompt) return setStatus('Hãy mô tả việc cần làm.', true);

  const payload = {
    template: state.templateText,
    templateName: state.templateName,
    data: state.dataText,
    dataName: state.dataName,
    prompt,
    fields: $('fields').value.trim(),
    customerCount: $('customerCount').value.trim(),
    expectedFiles: $('expectedFiles').value.trim(),
  };
  const password = $('password').value;
  remember('accessCode', password);
  $('submit').disabled = true;
  $('result').hidden = true;
  try {
    const plan = await requestPlan(payload, password);
    if (!plan.documents || !plan.documents.length) {
      throw new Error(plan.notes ? `AI không tạo được file nào. ${plan.notes}` : 'AI không tạo được file nào từ dữ liệu này.');
    }
    setStatus('Đang tạo file Word...');
    const files = await generateDocuments({
      JSZip, templateBytes: state.templateBytes, plan, dom,
      onProgress: (done, total) => setStatus(`Đang tạo file Word... ${done}/${total}`),
    });
    showResult(plan, files, payload);
    setStatus('Xong. Hãy mở thử vài file để kiểm tra trước khi gửi đi.');
  } catch (err) {
    setStatus(err.message, true);
  } finally {
    $('submit').disabled = false;
  }
});

$('downloadAll').addEventListener('click', async () => {
  if (!state.files.length) return;
  const zip = new JSZip();
  for (const file of state.files) zip.file(file.name, file.bytes, { createFolders: false });
  const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${state.templateName.replace(/\.docx$/i, '') || 'ket-qua'} - ${state.files.length} file.zip`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
});

init();
