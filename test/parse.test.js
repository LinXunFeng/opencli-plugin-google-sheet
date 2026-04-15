import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildColumnNames,
  extractSheetListFromHtml,
  normalizeCellValue,
  parseCsvRows,
  parseGvizResponse,
  parseGvizTableRows,
  resolveSheet,
} from '../src/parse.js';

test('extractSheetListFromHtml parses gid and title', () => {
  const html = `
    <script nonce="x">
      var bootstrap = {"sheets":[
        {"properties":{"sheetId":0,"title":"Sheet1"}},
        {"properties":{"sheetId":12345,"title":"Finance Q1"}}
      ]};
    </script>
  `;

  const sheets = extractSheetListFromHtml(html);
  assert.deepEqual(sheets, [
    { gid: '0', title: 'Sheet1', index: 0 },
    { gid: '12345', title: 'Finance Q1', index: 1 },
  ]);
});

test('extractSheetListFromHtml decodes escaped unicode titles', () => {
  const html = '{"sheetId":7,"title":"\\u4e2d\\u6587\\u8868"}';
  const sheets = extractSheetListFromHtml(html);
  assert.equal(sheets[0]?.title, '中文表');
});

test('extractSheetListFromHtml supports non-adjacent title fields', () => {
  const html = `
    {"sheetId":101,"gridProperties":{"rowCount":1000},"title":"Backlog"}
    {"sheetId":202,"gridProperties":{"rowCount":2000},"name":"Roadmap"}
  `;
  const sheets = extractSheetListFromHtml(html);
  assert.deepEqual(sheets, [
    { gid: '101', title: 'Backlog', index: 0 },
    { gid: '202', title: 'Roadmap', index: 1 },
  ]);
});

test('extractSheetListFromHtml parses htmlview gid anchors', () => {
  const html = `
    <div id="sheet-menu">
      <a href="/spreadsheets/d/doc123/htmlview?gid=2118959825">计划表</a>
      <a href="/spreadsheets/d/doc123/htmlview?gid=638061341"><span>Roadmap</span></a>
    </div>
  `;
  const sheets = extractSheetListFromHtml(html);
  assert.deepEqual(sheets, [
    { gid: '2118959825', title: '计划表', index: 0 },
    { gid: '638061341', title: 'Roadmap', index: 1 },
  ]);
});

test('extractSheetListFromHtml falls back to gid token scan', () => {
  const html = `
    <script>
      const x = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=2118959825";
      const y = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=638061341";
      const z = "/spreadsheets/d/doc123/htmlview?rm=minimal&amp;gid=508915592";
    </script>
  `;
  const sheets = extractSheetListFromHtml(html);
  assert.deepEqual(sheets, [
    { gid: '2118959825', title: 'Sheet_2118959825', index: 0 },
    { gid: '638061341', title: 'Sheet_638061341', index: 1 },
    { gid: '508915592', title: 'Sheet_508915592', index: 2 },
  ]);
});

test('extractSheetListFromHtml parses unquoted href anchors', () => {
  const html = `
    <a href=/spreadsheets/d/doc123/htmlview?gid=2118959825>计划</a>
    <a href=/spreadsheets/d/doc123/htmlview?gid=638061341>Roadmap</a>
  `;
  const sheets = extractSheetListFromHtml(html);
  assert.deepEqual(sheets, [
    { gid: '2118959825', title: '计划', index: 0 },
    { gid: '638061341', title: 'Roadmap', index: 1 },
  ]);
});

test('extractSheetListFromHtml upgrades fallback title when better title appears later', () => {
  const html = `
    <script>const gid = "2118959825";</script>
    <a href="/spreadsheets/d/doc123/htmlview?gid=2118959825">版本计划</a>
  `;
  const sheets = extractSheetListFromHtml(html);
  assert.deepEqual(sheets, [
    { gid: '2118959825', title: '版本计划', index: 0 },
  ]);
});

test('parseGvizResponse unwraps google visualization callback body', () => {
  const raw = 'google.visualization.Query.setResponse({"version":"0.6","status":"ok"});';
  const parsed = parseGvizResponse(raw);
  assert.equal(parsed.status, 'ok');
});

test('parseGvizTableRows uses formatted value first and fills blanks', () => {
  const payload = {
    table: {
      cols: [{ id: 'A' }, { id: 'B' }, { id: 'C' }],
      rows: [
        { c: [{ v: 'name' }, { v: 'amount' }, { v: 'note' }] },
        { c: [{ v: 'alice' }, { f: '$12.00', v: 12 }, null] },
        { c: [{ v: 'bob' }, { v: 8 }, { v: 'ok' }] },
      ],
    },
  };

  const rows = parseGvizTableRows(payload);
  assert.deepEqual(rows, [
    ['name', 'amount', 'note'],
    ['alice', '$12.00', ''],
    ['bob', '8', 'ok'],
  ]);
});

test('buildColumnNames prefers header row but falls back to A/B/C', () => {
  const first = buildColumnNames([
    ['Name', '', 'Score'],
    ['Alice', 'x', '10'],
  ]);
  assert.deepEqual(first, ['Name', 'B', 'Score']);

  const second = buildColumnNames([
    ['', '', ''],
    ['1', '2', '3'],
  ]);
  assert.deepEqual(second, ['A', 'B', 'C']);
});

test('resolveSheet matches gid first then case-insensitive title', () => {
  const sheets = [
    { gid: '0', title: 'Sheet1', index: 0 },
    { gid: '123', title: 'Finance Q1', index: 1 },
  ];

  assert.deepEqual(resolveSheet(sheets, '123'), sheets[1]);
  assert.deepEqual(resolveSheet(sheets, 'finance q1'), sheets[1]);
  assert.equal(resolveSheet(sheets, 'missing'), null);
});

test('normalizeCellValue stringifies primitives and null', () => {
  assert.equal(normalizeCellValue({ f: '1,234', v: 1234 }), '1,234');
  assert.equal(normalizeCellValue({ v: true }), 'true');
  assert.equal(normalizeCellValue(null), '');
});

test('parseCsvRows handles quoted commas and escaped quotes', () => {
  const csv = 'Name,Note\n"Alice","x,y"\n"Bob","He said ""Hi"""';
  const rows = parseCsvRows(csv);
  assert.deepEqual(rows, [
    ['Name', 'Note'],
    ['Alice', 'x,y'],
    ['Bob', 'He said "Hi"'],
  ]);
});
