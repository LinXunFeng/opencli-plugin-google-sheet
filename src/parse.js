/**
 * Parse helpers for Google Sheets metadata and table payloads.
 *
 * 处理 Google Sheets 元数据与表格载荷的解析辅助函数。
 */

/**
 * Decode escaped string fragments from Google payloads.
 *
 * 解码 Google 返回中被转义的字符串片段。
 */
export function decodeEscapedString(value) {
  if (!value) {
    return '';
  }

  try {
    return JSON.parse(`"${value.replace(/"/g, '\\"')}"`);
  } catch {
    return value;
  }
}

/**
 * Decode single-quoted escaped string fragments.
 *
 * 解码使用单引号包裹的转义字符串片段。
 */
function decodeMaybeSingleQuoted(value) {
  return decodeEscapedString(String(value || '').replace(/\\'/g, "'"));
}

/**
 * htmlview often embeds HTML entities in tab labels.
 *
 * htmlview 里的工作表标题常包含 HTML 实体，需要先还原。
 */
function decodeHtmlEntities(value) {
  return String(value || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

/**
 * Strip HTML tags from a string.
 *
 * 从字符串中移除 HTML 标签。
 */
function stripHtmlTags(value) {
  return String(value || '').replace(/<[^>]+>/g, ' ');
}

/**
 * Normalize title text from HTML snippets.
 *
 * 规范化 HTML 片段中的标题文本。
 */
function normalizeTitleText(value) {
  return decodeHtmlEntities(stripHtmlTags(String(value || ''))).replace(/\s+/g, ' ').trim();
}

/**
 * Generic fallback names are replaced when a better title is discovered later.
 *
 * 如果后续解析到更真实的标题，需要覆盖掉通用兜底名。
 */
function isGenericSheetTitle(title, gid) {
  const normalized = String(title || '').trim();
  return !normalized || normalized === `Sheet_${gid}` || normalized === `gid:${gid}`;
}

/**
 * Read a likely title field from a local JSON/JS chunk.
 *
 * 从局部 JSON/JS 片段中提取可能的标题字段。
 */
function pickTitleFromChunk(chunk) {
  const text = String(chunk || '');
  const doubleMatch = /(?:["']?(?:title|name)["']?)\s*:\s*"((?:\\.|[^"\\])*)"/.exec(text);
  const singleMatch = /(?:["']?(?:title|name)["']?)\s*:\s*'((?:\\.|[^'\\])*)'/.exec(text);

  if (doubleMatch && singleMatch) {
    if (doubleMatch.index <= singleMatch.index) {
      return decodeEscapedString(doubleMatch[1]);
    }
    return decodeMaybeSingleQuoted(singleMatch[1]);
  }

  if (doubleMatch) {
    return decodeEscapedString(doubleMatch[1]);
  }

  if (singleMatch) {
    return decodeMaybeSingleQuoted(singleMatch[1]);
  }

  return '';
}

/**
 * Extract worksheet list from HTML using layered heuristics.
 *
 * 使用分层启发式策略，从 HTML 中提取工作表列表。
 */
export function extractSheetListFromHtml(html) {
  const input = String(html || '');
  const list = [];
  const indexByGid = new Map();

  /**
   * Add or upgrade a worksheet record while preserving first-seen order.
   *
   * 新增或升级工作表记录，同时保持首次出现顺序不变。
   */
  const add = (gidValue, titleValue) => {
    const gid = String(gidValue || '').trim();
    if (!/^\d+$/.test(gid)) {
      return;
    }

    const title = normalizeTitleText(titleValue);
    // Keep first order for stable index, but allow title upgrade later.
    //
    // 保留首次出现顺序以稳定 index，同时允许后续补全更好的标题。
    const existingIndex = indexByGid.get(gid);
    if (existingIndex !== undefined) {
      if (title && isGenericSheetTitle(list[existingIndex]?.title, gid)) {
        list[existingIndex].title = title;
      }
      return;
    }

    indexByGid.set(gid, list.length);
    list.push({
      gid,
      title: title || `Sheet_${gid}`,
      index: list.length,
    });
  };

  // Metadata shape is unstable across edit/htmlview pages.
  // - key order can vary; keys may be quoted or unquoted.
  // - title key may be "title" or "name".
  //
  // Google 在 edit/htmlview 页面里的元数据结构不稳定。
  // - 字段顺序会变，key 可能有引号也可能没有。
  // - 标题字段可能是 "title" 或 "name"。
  const chunkPattern = /(?:["']?sheetId["']?)\s*:\s*(\d+)([\s\S]*?)(?=(?:["']?sheetId["']?\s*:)|$)/g;

  let match;
  while ((match = chunkPattern.exec(input)) !== null) {
    const gid = String(match[1] || '').trim();
    const tail = String(match[2] || '');
    const title = pickTitleFromChunk(tail);
    if (title) {
      add(gid, title);
    }
    if (list.length >= 1000) {
      break;
    }
  }

  // Fallback for payloads where title appears before sheetId.
  //
  // 兼容 title 在 sheetId 前面的对象结构。
  if (list.length === 0) {
    const reversePattern = /(?:["']?(?:title|name)["']?)\s*:\s*"((?:\\.|[^"\\])*)"[\s\S]{0,800}?(?:["']?sheetId["']?)\s*:\s*(\d+)/g;
    while ((match = reversePattern.exec(input)) !== null) {
      add(match[2], decodeEscapedString(match[1]));
      if (list.length >= 1000) {
        break;
      }
    }
  }

  // JS payload fallback: match gid/title pairs in embedded scripts.
  //
  // JS 片段兜底：从内联脚本中匹配 gid/title 成对信息。
  if (list.length <= 1) {
    const gidThenTitleDouble = /(?:["']?gid["']?)\s*[:=]\s*["']?(\d+)["']?[\s\S]{0,260}?(?:["']?(?:title|name|sheetName|label)["']?)\s*:\s*"((?:\\.|[^"\\])*)"/g;
    while ((match = gidThenTitleDouble.exec(input)) !== null) {
      add(match[1], decodeEscapedString(match[2]));
      if (list.length >= 1000) {
        break;
      }
    }
  }

  if (list.length <= 1) {
    const gidThenTitleSingle = /(?:["']?gid["']?)\s*[:=]\s*["']?(\d+)["']?[\s\S]{0,260}?(?:["']?(?:title|name|sheetName|label)["']?)\s*:\s*'((?:\\.|[^'\\])*)'/g;
    while ((match = gidThenTitleSingle.exec(input)) !== null) {
      add(match[1], decodeMaybeSingleQuoted(match[2]));
      if (list.length >= 1000) {
        break;
      }
    }
  }

  if (list.length <= 1) {
    const titleThenGidDouble = /(?:["']?(?:title|name|sheetName|label)["']?)\s*:\s*"((?:\\.|[^"\\])*)"[\s\S]{0,260}?(?:["']?gid["']?)\s*[:=]\s*["']?(\d+)["']?/g;
    while ((match = titleThenGidDouble.exec(input)) !== null) {
      add(match[2], decodeEscapedString(match[1]));
      if (list.length >= 1000) {
        break;
      }
    }
  }

  // htmlview fallback: tabs are often plain anchors containing gid.
  //
  // htmlview 兜底：很多页面把 tab 渲染成带 gid 的链接。
  if (list.length <= 1) {
    const linkPattern = /<a\b[^>]*href\s*=\s*(['"])([\s\S]*?)\1[^>]*>([\s\S]*?)<\/a>/gi;
    while ((match = linkPattern.exec(input)) !== null) {
      const href = String(match[2] || '');
      const gidMatch = /[?#&]gid=(\d+)/.exec(href);
      if (!gidMatch) {
        continue;
      }
      add(gidMatch[1], match[3]);
      if (list.length >= 1000) {
        break;
      }
    }
  }

  if (list.length <= 1) {
    const unquotedLinkPattern = /<a\b[^>]*href\s*=\s*([^'"\s>]+)[^>]*>([\s\S]*?)<\/a>/gi;
    while ((match = unquotedLinkPattern.exec(input)) !== null) {
      const href = String(match[1] || '');
      const gidMatch = /[?#&]gid=(\d+)/.exec(href);
      if (!gidMatch) {
        continue;
      }
      add(gidMatch[1], match[2]);
      if (list.length >= 1000) {
        break;
      }
    }
  }

  if (list.length <= 1) {
    const optionPattern = /<option\b[^>]*value\s*=\s*(?:(['"])([\s\S]*?)\1|([^'"\s>]+))[^>]*>([\s\S]*?)<\/option>/gi;
    while ((match = optionPattern.exec(input)) !== null) {
      const value = String(match[2] || match[3] || '');
      const gidMatch = /[?#&]gid=(\d+)/.exec(value);
      if (!gidMatch) {
        continue;
      }
      add(gidMatch[1], match[4]);
      if (list.length >= 1000) {
        break;
      }
    }
  }

  // Last resort: scan gid tokens only.
  // Some htmlview variants do not expose stable title nodes.
  //
  // 最后兜底：只扫描 gid，即使拿不到标题也能继续 read。
  // 某些 htmlview 变体不会暴露稳定标题节点，但 gid 通常仍可用。
  if (list.length <= 1) {
    const gidPattern = /(?:[?#&]|\\u0026|&amp;)gid=(\d+)/gi;
    while ((match = gidPattern.exec(input)) !== null) {
      add(match[1], '');
      if (list.length >= 1000) {
        break;
      }
    }
  }

  return list;
}

/**
 * Parse Google Visualization callback wrapper into JSON payload.
 *
 * 将 Google Visualization 的回调包裹格式解析为 JSON 载荷。
 */
export function parseGvizResponse(raw) {
  const match = /setResponse\(([\s\S]+)\);?\s*$/.exec(raw.trim());
  if (!match) {
    throw new Error('invalid gviz response');
  }

  return JSON.parse(match[1]);
}

/**
 * Normalize one gviz cell value to a display string.
 *
 * 将单个 gviz 单元格值标准化为可展示字符串。
 */
export function normalizeCellValue(cell) {
  if (!cell) {
    return '';
  }

  const value = cell.f ?? cell.v ?? '';
  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

/**
 * Convert gviz table payload into a 2D string matrix.
 *
 * 将 gviz 表格载荷转换为二维字符串矩阵。
 */
export function parseGvizTableRows(payload) {
  const cols = Array.isArray(payload?.table?.cols) ? payload.table.cols : [];
  const colCount = cols.length;
  const rows = Array.isArray(payload?.table?.rows) ? payload.table.rows : [];

  return rows.map((row) => {
    const cells = Array.isArray(row?.c) ? row.c : [];
    const currentColCount = Math.max(colCount, cells.length);
    return Array.from({ length: currentColCount }, (_, index) => normalizeCellValue(cells[index]));
  });
}

/**
 * Parse CSV text with quoted-field support.
 *
 * 解析 CSV 文本，支持带引号字段与转义双引号。
 */
export function parseCsvRows(text) {
  const input = String(text ?? '');
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;

  for (let i = 0; i < input.length; i += 1) {
    const ch = input[i];

    if (inQuotes) {
      if (ch === '"') {
        const next = input[i + 1];
        if (next === '"') {
          cell += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        cell += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }

    if (ch === ',') {
      row.push(cell);
      cell = '';
      continue;
    }

    if (ch === '\n') {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
      continue;
    }

    if (ch === '\r') {
      continue;
    }

    cell += ch;
  }

  if (cell.length > 0 || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }

  return rows;
}

/**
 * Convert a 1-based index to spreadsheet column letters.
 *
 * 将 1 开始的列索引转换为表格列字母。
 */
function columnLabel(index) {
  let n = index;
  let out = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out || 'A';
}

/**
 * Build output column names from rows.
 *
 * 从行数据构建输出列名。
 */
export function buildColumnNames(rows) {
  const header = Array.isArray(rows[0]) ? rows[0] : [];
  const colCount = rows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
  const hasHeader = header.some((cell) => String(cell ?? '').trim() !== '');

  if (!colCount) {
    return [];
  }

  if (!hasHeader) {
    return Array.from({ length: colCount }, (_, idx) => columnLabel(idx + 1));
  }

  return Array.from({ length: colCount }, (_, idx) => {
    const name = String(header[idx] ?? '').trim();
    return name || columnLabel(idx + 1);
  });
}

/**
 * Resolve sheet selector by gid first, then case-insensitive title.
 *
 * 先按 gid 匹配，再按不区分大小写的标题匹配工作表选择器。
 */
export function resolveSheet(sheets, selector) {
  if (!selector) {
    return null;
  }

  const normalized = String(selector).trim();
  if (!normalized) {
    return null;
  }

  const byGid = sheets.find((sheet) => sheet.gid === normalized);
  if (byGid) {
    return byGid;
  }

  const lower = normalized.toLowerCase();
  return sheets.find((sheet) => sheet.title.toLowerCase() === lower) ?? null;
}
