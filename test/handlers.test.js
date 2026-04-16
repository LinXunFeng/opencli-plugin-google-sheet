import test from 'node:test';
import assert from 'node:assert/strict';

import { createReadHandler, getRequestedFormat, toTableRows } from '../src/read-handler.js';
import { createSheetsHandler } from '../src/sheets-handler.js';

test('sheets handler returns worksheet list', async () => {
  const handler = createSheetsHandler({
    fetchWorkbookMeta: async () => [
      { gid: '0', title: 'Sheet1', index: 0 },
      { gid: '1', title: 'Sheet2', index: 1 },
    ],
  });

  const result = await handler({}, { 'docId': 'doc123' });
  assert.deepEqual(result, [
    { gid: '0', title: 'Sheet1', index: 0 },
    { gid: '1', title: 'Sheet2', index: 1 },
  ]);
});

test('read handler without --sheet returns sheet list in json mode', async () => {
  const handler = createReadHandler({
    fetchWorkbookMeta: async () => [{ gid: '0', title: 'Overview', index: 0 }],
    fetchSheetRows: async () => {
      throw new Error('should not call');
    },
    getRequestedFormat: () => 'json',
  });

  const result = await handler({}, { docId: 'doc123' });
  assert.equal(result.docId, 'doc123');
  assert.match(result.message, /please specify/i);
  assert.deepEqual(result.sheets, [{ gid: '0', title: 'Overview', index: 0 }]);
});

test('read handler without --sheet returns list rows in table mode', async () => {
  const handler = createReadHandler({
    fetchWorkbookMeta: async () => [{ gid: '0', title: 'Overview', index: 0 }],
    fetchSheetRows: async () => [],
    getRequestedFormat: () => 'table',
  });

  const result = await handler({}, { docId: 'doc123' });
  assert.deepEqual(result, [
    {
      gid: '0',
      title: 'Overview',
      index: 0,
      message: 'Use --sheet <name|gid> to read worksheet rows',
    },
  ]);
});

test('read handler resolves sheet and returns normalized payload', async () => {
  const handler = createReadHandler({
    fetchWorkbookMeta: async () => [{ gid: '0', title: 'Overview', index: 0 }],
    fetchSheetRows: async () => [
      ['Name', 'Score'],
      ['Alice', '90'],
      ['Bob', '80'],
    ],
    getRequestedFormat: () => 'json',
    now: () => '2026-04-14T00:00:00.000Z',
  });

  const result = await handler({}, { docId: 'doc123', sheet: 'Overview' });
  assert.deepEqual(result.sheet, { gid: '0', title: 'Overview' });
  assert.deepEqual(result.columns, ['Name', 'Score']);
  assert.deepEqual(result.rows, [['Alice', '90'], ['Bob', '80']]);
  assert.equal(result.meta.rowCount, 2);
  assert.equal(result.meta.colCount, 2);
  assert.equal(result.meta.fetchedAt, '2026-04-14T00:00:00.000Z');
});

test('read handler raises SHEET_NOT_FOUND when selector misses', async () => {
  const handler = createReadHandler({
    fetchWorkbookMeta: async () => [{ gid: '0', title: 'Overview', index: 0 }],
    fetchSheetRows: async () => [],
    getRequestedFormat: () => 'json',
  });

  await assert.rejects(() => handler({}, { docId: 'doc123', sheet: 'missing' }), (error) => {
    assert.equal(error.code, 'SHEET_NOT_FOUND');
    return true;
  });
});

test('read handler accepts numeric gid even when sheet list is incomplete', async () => {
  const handler = createReadHandler({
    fetchWorkbookMeta: async () => [{ gid: '0', title: 'Overview', index: 0 }],
    fetchSheetRows: async (_page, _docId, gid) => [
      ['Name', 'Score'],
      ['Alice', gid],
    ],
    getRequestedFormat: () => 'json',
    now: () => '2026-04-14T00:00:00.000Z',
  });

  const result = await handler({}, { docId: 'doc123', sheet: '638061341' });
  assert.deepEqual(result.sheet, { gid: '638061341', title: 'gid:638061341' });
  assert.deepEqual(result.rows, [['Alice', '638061341']]);
});

test('read handler skips workbook metadata lookup for numeric gid selector', async () => {
  const handler = createReadHandler({
    fetchWorkbookMeta: async () => {
      throw new Error('should not fetch workbook metadata for direct gid read');
    },
    fetchSheetRows: async (_page, _docId, gid) => [
      ['Name', 'Score'],
      ['Alice', gid],
    ],
    getRequestedFormat: () => 'json',
    now: () => '2026-04-14T00:00:00.000Z',
  });

  const result = await handler({}, { docId: 'doc123', sheet: '638061341' });
  assert.deepEqual(result.sheet, { gid: '638061341', title: 'gid:638061341' });
  assert.deepEqual(result.rows, [['Alice', '638061341']]);
});

test('toTableRows maps 2D rows into keyed objects', () => {
  const result = toTableRows(['Name', 'Score'], [['Alice', '90'], ['Bob', '80']]);
  assert.deepEqual(result, [
    { row: 1, Name: 'Alice', Score: '90' },
    { row: 2, Name: 'Bob', Score: '80' },
  ]);
});

test('getRequestedFormat prefers kwargs over argv flags', () => {
  assert.equal(
    getRequestedFormat(['node', 'read.js', '--format', 'table'], { format: 'json' }),
    'json',
  );
  assert.equal(
    getRequestedFormat(['node', 'read.js', '-f', 'table'], { f: 'md' }),
    'md',
  );
});

test('getRequestedFormat falls back to argv format flags', () => {
  assert.equal(getRequestedFormat(['node', 'read.js', '--format', 'table'], {}), 'table');
  assert.equal(getRequestedFormat(['node', 'read.js', '-f', 'csv'], {}), 'csv');
  assert.equal(getRequestedFormat(['node', 'read.js'], {}), 'json');
});
