import test from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { evalExpr, formatNumber, readVietnamese, renderValue, describeTemplate, buildDocumentXml, safeFilename } from '../public/core.js';

const dom = { DOMParser, XMLSerializer };
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const run = (text, bold) => `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const cell = (inner) => `<w:tc><w:p>${inner}</w:p></w:tc>`;
const XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>` +
  `<w:p>${run('Người', true)}${run(' nhận: ', true)}${run(' ')}${run(' ')}<w:r/></w:p>` +
  `<w:tbl>` +
  `<w:tr>${cell(run('1'))}${cell(run('Voucher'))}${cell('<w:r/>')}</w:tr>` +
  `<w:tr>${cell(run('2'))}${cell(run('Voucher'))}${cell('<w:r/>')}</w:tr>` +
  `<w:tr>${cell(run('Bằng chữ: ') + run('./.'))}</w:tr>` +
  `</w:tbl>` +
  `<w:p>${run('BÊN NHẬN', true)}</w:p>` +
  `</w:body></w:document>`;

const textOf = (xml, id) => {
  const line = describeTemplate(xml, dom).text.split('\n').find((l) => l.trim().startsWith(id + ':'));
  return line ? JSON.parse(line.trim().slice(id.length + 2).replace(/ \[\w+\]$/, '')) : undefined;
};

test('số học và định dạng số', () => {
  assert.equal(evalExpr('35*200000 + 50*500000'), 32000000);
  assert.equal(evalExpr('(64000000-4000000)/500000'), 120);
  assert.throws(() => evalExpr('alert(1)'));
  assert.throws(() => evalExpr('1.000.000'));
  assert.equal(formatNumber(1700200000), '1.700.200.000');
  assert.equal(formatNumber(12.5), '12,5');
});

test('đọc số thành chữ tiếng Việt', () => {
  assert.equal(readVietnamese(63000000), 'sáu mươi ba triệu');
  assert.equal(readVietnamese(144400000), 'một trăm bốn mươi tư triệu bốn trăm nghìn');
  assert.equal(readVietnamese(106000000), 'một trăm lẻ sáu triệu');
  assert.equal(readVietnamese(55600000), 'năm mươi lăm triệu sáu trăm nghìn');
  assert.equal(readVietnamese(1005), 'một nghìn không trăm lẻ năm');
  assert.equal(readVietnamese(21), 'hai mươi mốt');
  assert.equal(readVietnamese(15), 'mười lăm');
  assert.equal(readVietnamese(0), 'không');
});

test('renderValue thay biểu thức', () => {
  assert.equal(renderValue('{{num: 35*200000}}'), '7.000.000');
  assert.equal(renderValue('{{Words: 7000000}} đồng'), 'Bảy triệu đồng');
  const warnings = [];
  assert.equal(renderValue('{{num: abc}}', warnings), '{{num: abc}}');
  assert.equal(warnings.length, 1);
});

test('mô tả mẫu có ID', () => {
  const info = describeTemplate(XML, dom);
  assert.equal(info.paragraphs, 9);
  assert.equal(info.tables, 1);
  assert.ok(info.ids.includes('t0.r1.c2.p0'));
  assert.equal(textOf(XML, 'p0'), 'Người nhận:   ');
});

test('điền mẫu: giữ nhãn, xóa dòng, chèn đoạn', () => {
  const plan = {
    setup: [{ op: 'insert_paragraph', target: 'p1', copy_from: '', as: 'sig' }, { op: 'clone_row', target: 't0.r1', copy_from: '', as: 'extra' }],
    slots: [
      { key: 'ten', target: 'p0', template: 'Người nhận: {value}', align: 'keep', bold: false },
      { key: 'sl1', target: 't0.r0.c2.p0', template: '{value}', align: 'center', bold: false },
      { key: 'chu', target: 't0.r2.c0.p0', template: 'Bằng chữ: {value}./.', align: 'keep', bold: false },
      { key: 'ky', target: 'sig', template: '{value}', align: 'center', bold: true },
      { key: 'them', target: 'extra.c2.p0', template: '{value}', align: 'keep', bold: false },
    ],
  };
  const doc = {
    filename: 'a',
    values: [
      { key: 'ten', value: 'Trần Văn A' }, { key: 'sl1', value: '{{num: 35*200000}}' },
      { key: 'chu', value: '{{Words: 7000000}} đồng' }, { key: 'ky', value: 'Trần Văn A' }, { key: 'them', value: 'x' },
    ],
    remove: ['t0.r1'],
  };
  const { xml, warnings } = buildDocumentXml(XML, plan, doc, dom);
  assert.deepEqual(warnings, []);
  assert.ok(xml.startsWith('<?xml'));
  assert.equal(textOf(xml, 'p0'), 'Người nhận: Trần Văn A');
  assert.equal(textOf(xml, 't0.r0.c2.p0'), '7.000.000');
  // dòng r1 gốc bị xóa, dòng nhân bản lên thay vị trí r1
  assert.equal(textOf(xml, 't0.r1.c2.p0'), 'x');
  assert.equal(textOf(xml, 't0.r2.c0.p0'), 'Bằng chữ: Bảy triệu đồng./.');
  assert.equal(textOf(xml, 'p2'), 'Trần Văn A');
  assert.match(xml, /<w:jc w:val="center"\/>/);
  // nhãn "Người nhận" vẫn nằm trong run in đậm ban đầu
  assert.match(xml, /<w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Người<\/w:t>/);
});

test('cảnh báo khi ID không tồn tại', () => {
  const plan = { setup: [], slots: [{ key: 'k', target: 'p99', template: '{value}', align: 'keep', bold: false }] };
  const { warnings } = buildDocumentXml(XML, plan, { values: [{ key: 'k', value: 'v' }], remove: ['t9.r9'] }, dom);
  assert.equal(warnings.length, 2);
});

test('tên file an toàn và không trùng', () => {
  const used = new Set();
  assert.equal(safeFilename('BBGH: A/B', used), 'BBGH A B.docx');
  assert.equal(safeFilename('BBGH: A/B', used), 'BBGH A B (2).docx');
});
