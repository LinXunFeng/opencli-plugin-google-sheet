import { createGoogleSheetError } from './error.js';
import { buildColumnNames, resolveSheet } from './parse.js';
import { fetchSheetRows, fetchWorkbookMeta } from './sheets-api.js';

/**
 * Read command handler helpers.
 *
 * read 命令处理器辅助函数集合。
 */

const TABULAR_FORMATS = new Set(['table', 'csv', 'md', 'markdown', 'plain']);

/**
 * Return true when a value is non-empty after trimming.
 *
 * 判断值在去除空白后是否非空。
 */
function hasValue(value) {
  return String(value ?? '').trim() !== '';
}

/**
 * Detect whether a row is likely a header row.
 *
 * 判断一行是否更像表头行。
 */
function isHeaderRow(row) {
  return Array.isArray(row) && row.some((cell) => hasValue(cell));
}

/**
 * Normalize a column key name for object output.
 *
 * 规范化列名，使其可作为对象 key 使用。
 */
function normalizeKeyName(name) {
  return String(name || '')
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^\w]/g, '_') || 'column';
}

/**
 * Ensure unique object keys for duplicated column names.
 *
 * 为重复列名生成唯一对象 key。
 */
function makeUniqueKeys(columns) {
  const seen = new Map();
  return columns.map((name, index) => {
    const base = normalizeKeyName(name || `col_${index + 1}`);
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

/**
 * Convert 2D rows into list of row objects for table-like output.
 *
 * 将二维行数据转换为适合表格展示的对象数组。
 */
export function toTableRows(columns, rows) {
  const keys = makeUniqueKeys(columns);
  return rows.map((row, idx) => {
    const record = { row: idx + 1 };
    keys.forEach((key, colIndex) => {
      record[key] = String(row[colIndex] ?? '');
    });
    return record;
  });
}

/**
 * Read requested output format from process argv.
 *
 * 从命令行参数中读取输出格式。
 */
export function getRequestedFormat(argv = process.argv, kwargs = {}) {
  const fromKwargs = kwargs?.format ?? kwargs?.f;
  if (fromKwargs !== undefined && fromKwargs !== null && String(fromKwargs).trim() !== '') {
    return String(fromKwargs).toLowerCase();
  }

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if ((token === '-f' || token === '--format') && argv[i + 1]) {
      return String(argv[i + 1]).toLowerCase();
    }
  }
  return 'json';
}

/**
 * Best-effort gid selector check.
 *
 * 粗略判断 sheet 选择器是否为 gid。
 */
function isLikelyGid(selector) {
  return /^\d+$/.test(String(selector || '').trim());
}

/**
 * Create read command handler with dependency injection for tests.
 *
 * 创建 read 命令处理器，并支持依赖注入以便测试。
 */
export function createReadHandler(deps = {}) {
  const getWorkbookMeta = deps.fetchWorkbookMeta ?? fetchWorkbookMeta;
  const getSheetRows = deps.fetchSheetRows ?? fetchSheetRows;
  const requestedFormat = deps.getRequestedFormat ?? getRequestedFormat;
  const now = deps.now ?? (() => new Date().toISOString());

  return async function readHandler(page, kwargs) {
    const docId = kwargs.docId;
    const sheetSelector = kwargs.sheet;
    const sheets = await getWorkbookMeta(page, docId);
    const format = requestedFormat(process.argv, kwargs);
    const tabular = TABULAR_FORMATS.has(format);

    if (!sheetSelector) {
      if (tabular) {
        return sheets.map((sheet) => ({
          gid: sheet.gid,
          title: sheet.title,
          index: sheet.index,
          message: 'Use --sheet <name|gid> to read worksheet rows',
        }));
      }
      return {
        docId,
        message: 'Please specify --sheet <name|gid> to read worksheet rows',
        sheets,
      };
    }

    // Allow direct gid reads even when sheet list extraction is partial.
    //
    // 即使工作表列表提取不完整，也允许通过 gid 直接读取。
    let target = resolveSheet(sheets, sheetSelector);
    if (!target && isLikelyGid(sheetSelector)) {
      target = {
        gid: String(sheetSelector).trim(),
        title: `gid:${String(sheetSelector).trim()}`,
      };
    }

    if (!target) {
      throw createGoogleSheetError(
        'SHEET_NOT_FOUND',
        `Sheet not found: ${sheetSelector}`,
        `Available sheets: ${sheets.map((sheet) => `${sheet.title}(${sheet.gid})`).join(', ')}`,
      );
    }

    const parsedRows = await getSheetRows(page, docId, target.gid);
    const columns = buildColumnNames(parsedRows);
    const bodyRows = isHeaderRow(parsedRows[0]) ? parsedRows.slice(1) : parsedRows;
    const colCount = columns.length || bodyRows.reduce((max, row) => Math.max(max, row.length), 0);

    if (tabular) {
      return toTableRows(columns.length ? columns : Array.from({ length: colCount }, (_, i) => `col_${i + 1}`), bodyRows);
    }

    return {
      docId,
      sheet: {
        gid: target.gid,
        title: target.title,
      },
      columns: columns.length ? columns : Array.from({ length: colCount }, (_, i) => `col_${i + 1}`),
      rows: bodyRows,
      meta: {
        rowCount: bodyRows.length,
        colCount,
        fetchedAt: now(),
        source: 'opencli-plugin-google-sheet@0.0.1',
      },
    };
  };
}
