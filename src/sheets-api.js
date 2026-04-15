import {
  extractSheetListFromHtml,
  parseCsvRows,
  parseGvizResponse,
  parseGvizTableRows,
} from './parse.js';
import { createGoogleSheetError } from './error.js';

/**
 * Google Sheets network and page extraction utilities.
 *
 * Google Sheets 的网络访问与页面提取工具集合。
 */

const GOOGLE_DOCS_DOMAIN = 'docs.google.com';
const GOOGLE_PARENT_DOMAIN = 'google.com';
const GOOGLE_DOCS_ORIGIN = `https://${GOOGLE_DOCS_DOMAIN}`;
const DEFAULT_HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
  Origin: GOOGLE_DOCS_ORIGIN,
  Referer: `${GOOGLE_DOCS_ORIGIN}/`,
};

/**
 * Build a Google Sheets document URL by path.
 *
 * 按路径构造 Google Sheets 文档 URL。
 */
function buildDocUrl(docId, path = 'edit') {
  return `${GOOGLE_DOCS_ORIGIN}/spreadsheets/d/${encodeURIComponent(docId)}/${path}`;
}

/**
 * Build htmlview URL used for metadata fallback.
 *
 * 构造 htmlview URL，用于工作表元数据补偿提取。
 */
function buildHtmlViewUrl(docId) {
  return `${buildDocUrl(docId, 'htmlview')}?rm=minimal`;
}

/**
 * Build CSV export URL for a specific worksheet gid.
 *
 * 构造指定 gid 的 CSV 导出 URL。
 */
function buildExportCsvUrl(docId, gid) {
  return `${GOOGLE_DOCS_ORIGIN}/spreadsheets/d/${encodeURIComponent(docId)}/export?format=csv&gid=${encodeURIComponent(gid)}`;
}

/**
 * Convert cookie array to `Cookie` header value.
 *
 * 将 cookie 数组转换为请求头 `Cookie` 字符串。
 */
export function buildCookieHeader(cookies) {
  return cookies.map((cookie) => `${cookie.name}=${cookie.value}`).join('; ');
}

/**
 * Merge cookies from URL scope, docs domain, and parent google domain.
 * URL cookies take precedence to match live session behavior.
 *
 * 合并 URL 作用域、docs 域、google 父域下的 cookie。
 * URL 级 cookie 优先级更高，以贴近用户当前会话。
 */
async function getMergedCookies(page, url) {
  const [urlCookies, docsDomainCookies, parentDomainCookies] = await Promise.all([
    page.getCookies({ url }),
    page.getCookies({ domain: GOOGLE_DOCS_DOMAIN }),
    page.getCookies({ domain: GOOGLE_PARENT_DOMAIN }),
  ]);

  const merged = new Map();
  for (const cookie of docsDomainCookies) {
    merged.set(cookie.name, cookie);
  }
  for (const cookie of parentDomainCookies) {
    merged.set(cookie.name, cookie);
  }
  for (const cookie of urlCookies) {
    merged.set(cookie.name, cookie);
  }
  return Array.from(merged.values());
}

/**
 * Execute fetch inside browser context so credentials and redirects
 * follow the same policy as interactive browsing.
 *
 * 在浏览器上下文中执行 fetch，让凭证与重定向策略与人工访问一致。
 */
async function fetchWithinBrowserContext(page, url) {
  if (!page || typeof page.evaluate !== 'function') {
    return null;
  }

  try {
    const result = await page.evaluate(`
      (async () => {
        try {
          const response = await fetch(${JSON.stringify(url)}, {
            credentials: 'include',
            redirect: 'follow',
          });
          const text = await response.text();
          return {
            ok: response.ok,
            status: response.status,
            url: response.url,
            body: text,
          };
        } catch (error) {
          return {
            ok: false,
            status: 0,
            url: ${JSON.stringify(url)},
            body: '',
            error: String(error && error.message ? error.message : error),
          };
        }
      })()
    `);

    const hasValidShape = (
      result
      && typeof result === 'object'
      && typeof result.ok === 'boolean'
      && typeof result.status === 'number'
      && Number.isFinite(result.status)
      && typeof result.url === 'string'
    );

    if (!hasValidShape || result.status === 0) {
      return null;
    }

    return result;
  } catch {
    return null;
  }
}

async function fetchWithSession(page, url) {
  // Prefer browser-context fetch so auth/cookies match the user's live session.
  //
  // 优先在浏览器上下文内 fetch，确保鉴权与用户当前登录态一致。
  const browserResult = await fetchWithinBrowserContext(page, url);
  if (browserResult) {
    return browserResult;
  }

  const cookies = await getMergedCookies(page, url);
  const cookieHeader = buildCookieHeader(cookies);

  let response;
  try {
    response = await fetch(url, {
      headers: {
        ...DEFAULT_HEADERS,
        ...(cookieHeader ? { Cookie: cookieHeader } : {}),
      },
    });
  } catch (error) {
    throw createGoogleSheetError(
      'FETCH_FAILED',
      `Failed to fetch Google Sheets endpoint: ${error instanceof Error ? error.message : String(error)}`,
      'Check network connectivity and try again',
    );
  }

  const body = await response.text();
  return {
    ok: response.ok,
    status: response.status,
    url: response.url,
    body,
  };
}

/**
 * Fetch with Node + merged browser cookies.
 * Used when browser-context fetch is unavailable.
 *
 * 使用 Node + 合并后的浏览器 cookie 请求。
 * 主要用于浏览器上下文 fetch 不可用时。
 */
async function fetchWithNodeCookies(page, url, referer = `${GOOGLE_DOCS_ORIGIN}/`) {
  const cookies = await getMergedCookies(page, url);
  const cookieHeader = buildCookieHeader(cookies);

  let response;
  try {
    response = await fetch(url, {
      headers: {
        ...DEFAULT_HEADERS,
        Referer: referer,
        ...(cookieHeader ? { Cookie: cookieHeader } : {}),
      },
    });
  } catch (error) {
    return {
      ok: false,
      status: 0,
      url,
      body: '',
      error: `node-fetch-failed:${error instanceof Error ? error.message : String(error)}`,
    };
  }

  const body = await response.text();
  return {
    ok: response.ok,
    status: response.status,
    url: response.url,
    body,
  };
}

/**
 * Detect whether a response URL points to Google login flow.
 *
 * 判断响应 URL 是否进入 Google 登录流程。
 */
function isAuthRedirect(url) {
  try {
    const parsed = new URL(url);
    if (!/accounts\.google\.com$/i.test(parsed.hostname)) {
      return false;
    }
    return /servicelogin|signin|identifier/i.test(parsed.pathname + parsed.search);
  } catch {
    return false;
  }
}

/**
 * Heuristic: identify HTML-like content quickly.
 *
 * 启发式判断文本是否像 HTML。
 */
function looksLikeHtml(text) {
  return /^\s*<(?:!doctype|html|head|body)\b/i.test(String(text || ''));
}

/**
 * Build edit URL pinned to a specific worksheet gid.
 *
 * 构造定位到指定 gid 的 edit URL。
 */
function buildSheetEditUrl(docId, gid) {
  return `${buildDocUrl(docId, 'edit')}?gid=${encodeURIComponent(gid)}#gid=${encodeURIComponent(gid)}`;
}

/**
 * Compress long response bodies into short debug previews.
 *
 * 将长响应体压缩为短调试预览文本。
 */
function compactPreview(text, max = 80) {
  return String(text || '').replace(/\s+/g, ' ').trim().slice(0, max);
}

/**
 * Detect whether a parsed worksheet title is still a generic fallback value.
 *
 * 判断解析出的工作表标题是否仍为兜底占位值。
 */
function isGenericWorksheetTitle(title, gid) {
  const normalizedTitle = String(title || '').trim();
  const normalizedGid = String(gid || '').trim();
  return !normalizedTitle || normalizedTitle === `Sheet_${normalizedGid}` || normalizedTitle === `gid:${normalizedGid}`;
}

/**
 * Score sheet list quality by count first, then non-generic title coverage.
 *
 * 先按数量、再按“非占位标题”的覆盖率评估工作表列表质量。
 */
function scoreSheetListQuality(list) {
  const sheets = Array.isArray(list) ? list : [];
  let namedCount = 0;
  for (const sheet of sheets) {
    if (!isGenericWorksheetTitle(sheet?.title, sheet?.gid)) {
      namedCount += 1;
    }
  }
  return {
    count: sheets.length,
    namedCount,
  };
}

/**
 * Decide whether candidate metadata is better than current best.
 * Better means: more sheets, or same count with more named titles.
 *
 * 判断候选元数据是否优于当前最佳结果。
 * 规则是：数量更多，或数量相同但真实标题更多。
 */
function isBetterSheetListCandidate(candidate, best) {
  const candidateScore = scoreSheetListQuality(candidate);
  const bestScore = scoreSheetListQuality(best);
  if (candidateScore.count !== bestScore.count) {
    return candidateScore.count > bestScore.count;
  }
  return candidateScore.namedCount > bestScore.namedCount;
}

/**
 * Return true when a sheet list still contains generic placeholder titles.
 *
 * 判断工作表列表中是否仍存在占位标题。
 */
function hasGenericWorksheetTitles(sheets) {
  return (Array.isArray(sheets) ? sheets : []).some((sheet) => isGenericWorksheetTitle(sheet?.title, sheet?.gid));
}

/**
 * Merge known titles into a base sheet list without changing base order.
 * Only upgrades generic/empty titles to better candidate titles.
 *
 * 在不改变 base 顺序的前提下，把已知标题合并进结果。
 * 仅把占位/空标题升级为更可靠的候选标题。
 */
function mergeSheetTitles(baseSheets, ...sourceLists) {
  const merged = (Array.isArray(baseSheets) ? baseSheets : []).map((sheet, index) => ({
    gid: String(sheet?.gid || ''),
    title: String(sheet?.title || '').trim() || `Sheet_${String(sheet?.gid || '').trim()}`,
    index: Number.isFinite(sheet?.index) ? sheet.index : index,
  }));

  const indexByGid = new Map();
  for (let i = 0; i < merged.length; i += 1) {
    const gid = String(merged[i].gid || '').trim();
    if (gid) {
      indexByGid.set(gid, i);
    }
  }

  for (const list of sourceLists) {
    for (const sourceSheet of Array.isArray(list) ? list : []) {
      const gid = String(sourceSheet?.gid || '').trim();
      const title = String(sourceSheet?.title || '').trim();
      if (!gid || !title) {
        continue;
      }

      const targetIndex = indexByGid.get(gid);
      if (targetIndex === undefined) {
        continue;
      }

      if (isGenericWorksheetTitle(merged[targetIndex].title, gid) && !isGenericWorksheetTitle(title, gid)) {
        merged[targetIndex].title = title;
      }
    }
  }

  return merged;
}

/**
 * Normalize worksheet title strings collected from live page probes.
 *
 * 规范化从页面探测得到的工作表标题文本。
 */
function normalizeWorksheetTitle(value) {
  const raw = String(value || '').replace(/\s+/g, ' ').trim();
  if (!raw) {
    return '';
  }
  return raw.replace(/\s*[-|]\s*Google(?:\s+Sheets|\s+\S+表格)?\s*$/i, '').trim();
}

/**
 * Verify gid-title mapping by visiting each gid and reading active tab label.
 * This is slower but provides a high-confidence correction path.
 *
 * 通过逐个 gid 跳转并读取当前激活 tab 标题来校验映射。
 * 该路径更慢，但可作为高置信度纠偏方案。
 */
async function tryResolveTitlesByGidNavigation(page, docId, sheets) {
  if (!page || typeof page.goto !== 'function' || typeof page.evaluate !== 'function') {
    return null;
  }

  const targets = (Array.isArray(sheets) ? sheets : []).slice(0, 12);
  if (targets.length < 2) {
    return null;
  }

  const resolved = [];
  for (const sheet of targets) {
    const gid = String(sheet?.gid || '').trim();
    if (!gid) {
      continue;
    }

    try {
      await page.goto(buildSheetEditUrl(docId, gid));
      if (typeof page.wait === 'function') {
        await page.wait(1);
      }

      const title = await page.evaluate(`
        (() => new Promise((resolve) => {
          // verify title by gid navigation
          const targetGid = ${JSON.stringify(gid)};
          const deadline = Date.now() + 8000;

          const normalizeTitle = (value) => {
            const raw = String(value || '').replace(/\\s+/g, ' ').trim();
            if (!raw) return '';
            return raw.replace(/\\s*[-|]\\s*Google(?:\\s+Sheets|\\s+\\S+表格)?\\s*$/i, '').trim();
          };

          const gidFromText = (text) => {
            const match = String(text || '').match(/[?#&]gid=(\\d+)/);
            return match ? match[1] : '';
          };

          const gidFromNode = (node) => {
            if (!node) return '';
            const idMatch = String(node.id || '').match(/sheet-button-(\\d+)/);
            const fromId = idMatch ? idMatch[1] : '';
            const fromData = String(
              (node.dataset && (node.dataset.gid || node.dataset.sheetId))
              || (node.getAttribute && (node.getAttribute('data-gid') || node.getAttribute('data-sheet-id')))
              || ''
            ).trim();
            const fromHref = gidFromText(node.getAttribute && node.getAttribute('href'));
            return String(fromId || fromData || fromHref || '').trim();
          };

          const readTitle = (node) => {
            if (!node) return '';
            return normalizeTitle(
              (node.getAttribute && (node.getAttribute('aria-label') || node.getAttribute('data-tooltip') || node.getAttribute('title')))
              || node.textContent
              || ''
            );
          };

          const isActiveNode = (node) => {
            if (!node) return false;
            const ariaSelected = String(node.getAttribute && node.getAttribute('aria-selected') || '').toLowerCase() === 'true';
            const ariaCurrent = String(node.getAttribute && node.getAttribute('aria-current') || '').toLowerCase() === 'true';
            const className = String(node.className || '');
            return ariaSelected || ariaCurrent || /active|selected|current/i.test(className);
          };

          const getTabNodes = () => Array.from(
            document.querySelectorAll('[id^="sheet-button-"], [role="tab"], a[href*="gid="], [data-gid], [data-sheet-id]')
          );

          const readActiveTargetTitle = () => {
            for (const node of getTabNodes()) {
              if (!isActiveNode(node)) continue;
              const nodeGid = gidFromNode(node);
              if (nodeGid && nodeGid !== targetGid) continue;
              const title = readTitle(node);
              if (title) return title;
            }
            return '';
          };

          const clickTargetTab = () => {
            for (const node of getTabNodes()) {
              if (gidFromNode(node) !== targetGid) continue;
              try {
                if (typeof node.click === 'function') {
                  node.click();
                  return true;
                }
                if (typeof node.dispatchEvent === 'function') {
                  node.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                  return true;
                }
              } catch {
                // ignore click errors
              }
            }
            return false;
          };

          const readTargetNodeTitle = () => {
            for (const node of getTabNodes()) {
              if (gidFromNode(node) !== targetGid) continue;
              const title = readTitle(node);
              if (title) return title;
            }
            return '';
          };

          const tick = () => {
            const activeTitle = readActiveTargetTitle();
            if (activeTitle) {
              resolve(activeTitle);
              return;
            }

            const clicked = clickTargetTab();
            if (Date.now() >= deadline) {
              const byNode = readTargetNodeTitle();
              if (byNode) {
                resolve(byNode);
                return;
              }

              const fromUrl = gidFromText(location.href);
              if (fromUrl === targetGid) {
                resolve(normalizeTitle(document.title || ''));
                return;
              }
              resolve('');
              return;
            }

            setTimeout(tick, clicked ? 350 : 200);
          };

          tick();
        }))()
      `);

      const normalized = normalizeWorksheetTitle(title);
      if (normalized && !isGenericWorksheetTitle(normalized, gid)) {
        resolved.push({ gid, title: normalized });
      }
    } catch {
      // keep best-effort behavior
    }
  }

  if (resolved.length < 2) {
    return null;
  }

  const uniqueTitles = new Set(resolved.map((item) => String(item.title || '').toLowerCase()));
  if (uniqueTitles.size < Math.max(2, Math.floor(resolved.length * 0.6))) {
    return null;
  }

  return resolved;
}

/**
 * Final fallback for row extraction.
 * Reads currently visible grid cells from DOM after navigating to the sheet.
 *
 * 行数据提取的最后兜底方案。
 * 跳转到目标工作表后，从 DOM 中读取当前可见网格单元格。
 */
async function tryExtractVisibleRowsFromLoadedPage(page, docId, gid) {
  if (!page || typeof page.goto !== 'function' || typeof page.evaluate !== 'function') {
    return null;
  }

  try {
    await page.goto(buildSheetEditUrl(docId, gid));
    if (typeof page.wait === 'function') {
      await page.wait(2);
    }

    const rows = await page.evaluate(`
      (() => {
        // Read visible grid cells from DOM as a final fallback path.
        //
        // 作为最后兜底，从页面 DOM 里直接采集当前可见网格单元格。
        const cells = Array.from(document.querySelectorAll('[role="gridcell"][aria-rowindex][aria-colindex]'));
        if (cells.length === 0) {
          return [];
        }

        const points = [];
        for (const node of cells) {
          const row = Number(node.getAttribute('aria-rowindex'));
          const col = Number(node.getAttribute('aria-colindex'));
          if (!Number.isFinite(row) || !Number.isFinite(col)) {
            continue;
          }
          const text = String(node.textContent || '').trim();
          points.push({ row, col, text });
        }

        const nonEmpty = points.filter((point) => point.text !== '');
        if (nonEmpty.length === 0) {
          return [];
        }

        const minRow = Math.min(...nonEmpty.map((point) => point.row));
        const maxRow = Math.max(...nonEmpty.map((point) => point.row));
        const minCol = Math.min(...nonEmpty.map((point) => point.col));
        const maxCol = Math.max(...nonEmpty.map((point) => point.col));

        const rowCount = Math.min(500, maxRow - minRow + 1);
        const colCount = Math.min(200, maxCol - minCol + 1);
        const out = Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => ''));

        for (const point of points) {
          const r = point.row - minRow;
          const c = point.col - minCol;
          if (r >= 0 && r < rowCount && c >= 0 && c < colCount) {
            out[r][c] = point.text;
          }
        }

        return out;
      })()
    `);

    if (Array.isArray(rows) && rows.length > 0) {
      return rows;
    }
  } catch {
    return null;
  }

  return null;
}

/**
 * Try extracting worksheet list from a fully loaded edit page.
 * Returns both result list and confidence metadata.
 *
 * 尝试从已加载完成的 edit 页面提取工作表列表。
 * 返回结果列表以及置信度元信息。
 */
async function tryExtractSheetsFromLoadedPage(page, editUrl) {
  if (!page || typeof page.goto !== 'function' || typeof page.evaluate !== 'function') {
    return null;
  }

  try {
    await page.goto(editUrl);

    if (typeof page.wait === 'function') {
      await page.wait(2);
    }

    const currentUrl = await page.evaluate('location.href');
    if (isAuthRedirect(String(currentUrl || ''))) {
      throw createGoogleSheetError(
        'AUTH_REQUIRED',
        `Google account login required for this sheet (page redirect: ${String(currentUrl)})`,
        'Open docs.google.com in Chrome/Chromium, ensure login is valid, then retry',
      );
    }

    const extraction = await page.evaluate(`
      (() => new Promise((resolve) => {
        const deadline = Date.now() + 8000;

        const normalizeTitle = (value) => {
          const raw = String(value || '').trim();
          if (!raw) return '';
          return raw.replace(/\\s*[-|]\\s*Google Sheets.*$/i, '').trim();
        };

        const gidFromText = (text) => {
          const match = String(text || '').match(/[?#]gid=(\\d+)/);
          return match ? match[1] : '';
        };

        const sliceLikelyObjectTail = (value) => {
          const text = String(value || '');
          const boundaryMatch = /}\\s*,?\\s*[{[]/.exec(text);
          if (!boundaryMatch) return text;
          return text.slice(0, boundaryMatch.index + 1);
        };

        const collect = () => {
          const out = [];
          const seen = new Set();
          let fromDomCount = 0;
          let fromHtmlCount = 0;

          const add = (gidValue, titleValue, source) => {
            const gid = String(gidValue || '').trim();
            if (!gid) return;
            if (seen.has(gid)) {
              return;
            }
            seen.add(gid);
            out.push({
              gid,
              title: normalizeTitle(titleValue) || ('Sheet_' + gid),
              index: out.length,
            });
            if (source === 'dom') fromDomCount += 1;
            if (source === 'html') fromHtmlCount += 1;
          };

          // Strategy 1 - visible tab controls in DOM.
          //
          // 策略 1：优先读取 DOM 中已渲染的 tab 控件。
          const tabNodes = document.querySelectorAll('[id^="sheet-button-"], [role="tab"], a[href*="#gid="]');
          for (const node of tabNodes) {
            const idMatch = String(node.id || '').match(/sheet-button-(\\d+)/);
            const fromId = idMatch ? idMatch[1] : '';
            const fromHref = gidFromText(node.getAttribute && node.getAttribute('href'));
            const fromData = String(
              (node.dataset && (node.dataset.gid || node.dataset.sheetId))
              || (node.getAttribute && (node.getAttribute('data-gid') || node.getAttribute('data-sheet-id')))
              || ''
            );
            const gid = fromId || fromHref || fromData;
            if (!gid) continue;

            const title =
              (node.getAttribute && (node.getAttribute('aria-label') || node.getAttribute('data-tooltip'))) ||
              node.textContent ||
              '';
            add(gid, title, 'dom');
          }

          // Strategy 2 - parse bootstrap metadata from full HTML.
          //
          // 策略 2：从整页 HTML 的初始化数据中解析。
          if (out.length === 0) {
            const html = document.documentElement ? document.documentElement.innerHTML : '';
            const re = /(?:["']?sheetId["']?)\\s*:\\s*(\\d+)([\\s\\S]*?)(?=(?:["']?sheetId["']?\\s*:)|$)/g;
            let match;
            while ((match = re.exec(html)) !== null) {
              const gid = String(match[1] || '');
              if (!gid || seen.has(gid)) {
                continue;
              }

              const chunk = sliceLikelyObjectTail(match[2] || '');
              const titleDouble = /(?:["']?(?:title|name)["']?)\\s*:\\s*"((?:\\\\.|[^"\\\\])*)"/.exec(chunk);
              const titleSingle = /(?:["']?(?:title|name)["']?)\\s*:\\s*'((?:\\\\.|[^'\\\\])*)'/.exec(chunk);
              let title = '';
              if (titleDouble && titleSingle) {
                title = titleDouble.index <= titleSingle.index ? String(titleDouble[1] || '') : String(titleSingle[1] || '').replace(/\\\\'/g, "'");
              } else if (titleDouble) {
                title = String(titleDouble[1] || '');
              } else if (titleSingle) {
                title = String(titleSingle[1] || '').replace(/\\\\'/g, "'");
              }

              if (!title) {
                continue;
              }

              try {
                title = JSON.parse('"' + String(title).replace(/"/g, '\\\\\\"') + '"');
              } catch {
                // keep raw title
              }
              add(gid, title, 'html');
              if (out.length >= 500) {
                break;
              }
            }

            const reverseDouble = /(?:["']?(?:title|name)["']?)\\s*:\\s*"((?:\\\\.|[^"\\\\])*)"([\\s\\S]{0,800}?)(?:["']?sheetId["']?)\\s*:\\s*(\\d+)/g;
            while ((match = reverseDouble.exec(html)) !== null) {
              const gid = String(match[3] || '').trim();
              if (!gid || seen.has(gid)) {
                continue;
              }
              let title = String(match[1] || '');
              try {
                title = JSON.parse('"' + title.replace(/"/g, '\\\\\\"') + '"');
              } catch {
                // keep raw title
              }
              add(gid, title, 'html');
              if (out.length >= 500) {
                break;
              }
            }

            if (out.length < 500) {
              const reverseSingle = /(?:["']?(?:title|name)["']?)\\s*:\\s*'((?:\\\\.|[^'\\\\])*)'([\\s\\S]{0,800}?)(?:["']?sheetId["']?)\\s*:\\s*(\\d+)/g;
              while ((match = reverseSingle.exec(html)) !== null) {
                const gid = String(match[3] || '').trim();
                if (!gid || seen.has(gid)) {
                  continue;
                }
                const title = String(match[1] || '').replace(/\\\\'/g, "'");
                add(gid, title, 'html');
                if (out.length >= 500) {
                  break;
                }
              }
            }
          }

          // Strategy 3 - fallback to the current gid only.
          //
          // 策略 3：只回退到当前 URL 的 gid（低置信度结果）。
          if (out.length === 0) {
            const currentGid = gidFromText(location.href);
            if (currentGid) {
              add(currentGid, document.title || 'Current Sheet', 'fallback');
            }
          }

          const confident = out.length > 1 || fromHtmlCount > 0 || fromDomCount > 1;
          return { out, confident };
        };

        const tick = () => {
          const { out, confident } = collect();
          if (confident || Date.now() >= deadline) {
            resolve({
              sheets: out,
              confident,
            });
            return;
          }
          setTimeout(tick, 250);
        };

        tick();
      }))
    `);

    if (Array.isArray(extraction) && extraction.length > 0) {
      return {
        sheets: extraction,
        confident: extraction.length > 1,
      };
    }

    if (extraction && typeof extraction === 'object' && Array.isArray(extraction.sheets) && extraction.sheets.length > 0) {
      return extraction;
    }
  } catch (error) {
    if (error && typeof error === 'object' && error.code === 'AUTH_REQUIRED') {
      throw error;
    }
  }

  return null;
}

/**
 * Extract tab titles from the live edit page DOM for title hydration only.
 * This path ignores generic fallbacks and focuses on explicit tab labels.
 *
 * 仅用于标题补全：从 edit 页 DOM 提取 tab 标题。
 * 该路径不使用占位兜底，只关注明确可见的 tab 文本。
 */
async function tryExtractTabTitlesFromLoadedPage(page, docId) {
  if (!page || typeof page.goto !== 'function' || typeof page.evaluate !== 'function') {
    return null;
  }

  try {
    await page.goto(buildDocUrl(docId, 'edit'));

    if (typeof page.wait === 'function') {
      await page.wait(2);
    }

    const titles = await page.evaluate(`
      (() => {
        // Read visible worksheet tabs from DOM to recover real titles.
        //
        // 从 DOM 中读取可见工作表 tab，补全真实标题。
        const normalizeTitle = (value) => {
          const raw = String(value || '').replace(/\\s+/g, ' ').trim();
          if (!raw) return '';
          return raw.replace(/\\s*[-|]\\s*Google(?:\\s+Sheets|\\s+\\S+表格)?\\s*$/i, '').trim();
        };

        const gidFromText = (text) => {
          const match = String(text || '').match(/[?#&]gid=(\\d+)/);
          return match ? match[1] : '';
        };

        const readTitle = (node) => {
          return normalizeTitle(
            (node.getAttribute && (node.getAttribute('aria-label') || node.getAttribute('data-tooltip') || node.getAttribute('title')))
            || node.textContent
            || ''
          );
        };

        const out = [];
        const seen = new Set();
        const tabNodes = document.querySelectorAll('[id^="sheet-button-"], [role="tab"], a[href*="gid="], [data-gid], [data-sheet-id]');
        for (const node of tabNodes) {
          const idMatch = String(node.id || '').match(/sheet-button-(\\d+)/);
          const fromId = idMatch ? idMatch[1] : '';
          const fromHref = gidFromText(node.getAttribute && node.getAttribute('href'));
          const fromData = String(
            (node.dataset && (node.dataset.gid || node.dataset.sheetId))
            || (node.getAttribute && (node.getAttribute('data-gid') || node.getAttribute('data-sheet-id')))
            || ''
          );
          const gid = String(fromId || fromHref || fromData || '').trim();
          if (!gid || seen.has(gid)) {
            continue;
          }

          const title = readTitle(node);
          if (!title) {
            continue;
          }

          seen.add(gid);
          out.push({ gid, title });
        }

        return out;
      })()
    `);

    if (Array.isArray(titles) && titles.length > 0) {
      return titles;
    }
  } catch {
    return null;
  }

  return null;
}

/**
 * Hydrate placeholder titles with better metadata from other sources.
 * Optionally re-reads tab titles from the live edit page when needed.
 *
 * 用更多来源补全占位标题。
 * 如仍有占位名，则再尝试从 edit 页实时 tab 补全。
 */
async function hydrateWorksheetTitles(page, docId, baseSheets, ...sourceLists) {
  const initial = mergeSheetTitles(baseSheets, ...sourceLists);
  if (!hasGenericWorksheetTitles(initial)) {
    return initial;
  }

  const liveTitles = await tryExtractTabTitlesFromLoadedPage(page, docId);
  if (!liveTitles || liveTitles.length === 0) {
    return initial;
  }

  return mergeSheetTitles(initial, liveTitles);
}

/**
 * Convert auth-like HTTP states into domain-specific errors.
 *
 * 将鉴权相关 HTTP 状态转换为插件领域错误。
 */
function ensureAuthorized(response) {
  const redirectedToLogin = isAuthRedirect(response.url);
  if (redirectedToLogin || response.status === 401 || response.status === 403) {
    throw createGoogleSheetError(
      'AUTH_REQUIRED',
      `Google account login required for this sheet (fetch status=${response.status}, url=${response.url})`,
      'Open docs.google.com in Chrome/Chromium, ensure login is valid, then retry',
    );
  }
}

/**
 * Query htmlview endpoints for fuller worksheet metadata.
 * Falls back to page navigation when fetch result is weak.
 *
 * 尝试通过 htmlview 端点补全工作表元数据。
 * 当 fetch 结果偏弱时，回退到页面导航再解析。
 */
async function tryFetchWorkbookMetaFromHtmlView(page, docId) {
  const htmlViewUrls = [buildHtmlViewUrl(docId), buildDocUrl(docId, 'htmlview')];

  let best = [];
  for (const url of htmlViewUrls) {
    let response;
    try {
      response = await fetchWithSession(page, url);
    } catch {
      continue;
    }

    if (!response || !response.ok) {
      continue;
    }

    if (isAuthRedirect(response.url) || response.status === 401 || response.status === 403) {
      continue;
    }

    const parsed = extractSheetListFromHtml(response.body);
    if (isBetterSheetListCandidate(parsed, best)) {
      best = parsed;
    }

    const score = scoreSheetListQuality(best);
    // Stop early only when we already have multiple sheets with all titles named.
    //
    // 仅在“多工作表且标题都不是占位名”时提前结束。
    if (score.count > 1 && score.namedCount === score.count) {
      break;
    }
  }

  // If fetch-based htmlview probing is weak, navigate and parse live DOM HTML.
  //
  // 如果 fetch 到的 htmlview 信息不足，则直接打开页面解析实时 DOM。
  const bestScore = scoreSheetListQuality(best);
  const shouldUseDomFallback = bestScore.count <= 1 || bestScore.namedCount === 0;
  if (shouldUseDomFallback && page && typeof page.goto === 'function' && typeof page.evaluate === 'function') {
    for (const url of htmlViewUrls) {
      try {
        await page.goto(url);
        if (typeof page.wait === 'function') {
          await page.wait(2);
        }
        const html = await page.evaluate('document.documentElement ? document.documentElement.outerHTML : ""');
        const parsed = extractSheetListFromHtml(html);
        if (isBetterSheetListCandidate(parsed, best)) {
          best = parsed;
        }

        const score = scoreSheetListQuality(best);
        if (score.count > 1 && score.namedCount === score.count) {
          break;
        }
      } catch {
        // keep best-effort behavior
      }
    }
  }

  return best.length > 0 ? best : null;
}

/**
 * Fetch workbook worksheet metadata using progressive fallbacks.
 * Preference order:
 * 1) loaded page extraction;
 * 2) edit HTML metadata;
 * 3) htmlview metadata.
 *
 * 通过逐级回退策略获取工作簿工作表元数据。
 * 优先级顺序：
 * 1）已加载页面提取；
 * 2）edit 页 HTML 元数据；
 * 3）htmlview 元数据。
 */
export async function fetchWorkbookMeta(page, docId) {
  const url = buildDocUrl(docId, 'edit');
  const pageExtraction = await tryExtractSheetsFromLoadedPage(page, url);
  const pageSheets = Array.isArray(pageExtraction?.sheets) ? pageExtraction.sheets : [];
  const hasWeakPageSheets = pageSheets.length > 0 && !Boolean(pageExtraction?.confident);

  let response;
  try {
    response = await fetchWithSession(page, url);
  } catch (error) {
    if (pageSheets.length > 0) {
      const hydrated = await hydrateWorksheetTitles(page, docId, pageSheets);
      return hydrated;
    }
    throw error;
  }
  const body = response.body;

  if (response.status === 404) {
    throw createGoogleSheetError('DOC_NOT_FOUND', `Google Sheet not found: ${docId}`, 'Check docId and access permission');
  }

  try {
    ensureAuthorized(response);
  } catch (error) {
    if (pageSheets.length > 0) {
      const hydrated = await hydrateWorksheetTitles(page, docId, pageSheets);
      return hydrated;
    }
    throw error;
  }

  if (!response.ok) {
    throw createGoogleSheetError(
      'FETCH_FAILED',
      `Google Sheets metadata request failed with HTTP ${response.status}`,
      'Try again later or verify the document is accessible in browser',
    );
  }

  const sheets = extractSheetListFromHtml(body);
  let selectedSheets = sheets;
  if (isBetterSheetListCandidate(pageSheets, selectedSheets)) {
    selectedSheets = pageSheets;
  }

  if (selectedSheets.length <= 1) {
    // edit page may only expose one tab; htmlview usually contains full tab list.
    //
    // edit 页面有时只暴露一个 tab，htmlview 通常能拿到完整列表。
    const htmlViewSheets = await tryFetchWorkbookMetaFromHtmlView(page, docId);
    if (isBetterSheetListCandidate(htmlViewSheets, selectedSheets) && htmlViewSheets.length > pageSheets.length) {
      selectedSheets = htmlViewSheets;
    }
  }

  if (!selectedSheets.length) {
    if (/accounts\.google\.com\/ServiceLogin/i.test(body)) {
      if (hasWeakPageSheets) {
        return pageSheets;
      }
      throw createGoogleSheetError(
        'AUTH_REQUIRED',
        'Google account login required for this sheet (metadata missing and login marker found in HTML)',
        'Open docs.google.com in Chrome/Chromium, ensure login is valid, then retry',
      );
    }
    throw createGoogleSheetError(
      'PARSE_FAILED',
      'Failed to parse worksheet list from Google Sheets page',
      'Google Sheets page structure may have changed',
    );
  }

  const hydrated = await hydrateWorksheetTitles(page, docId, selectedSheets, sheets, pageSheets);
  const verifiedTitles = await tryResolveTitlesByGidNavigation(page, docId, hydrated);
  if (!verifiedTitles || verifiedTitles.length === 0) {
    return hydrated;
  }

  const titleByGid = new Map(
    verifiedTitles.map((entry) => [String(entry.gid || '').trim(), String(entry.title || '').trim()]),
  );

  let diffCount = 0;
  const corrected = hydrated.map((sheet, index) => {
    const gid = String(sheet?.gid || '').trim();
    const nextTitle = normalizeWorksheetTitle(titleByGid.get(gid) || '');
    const currentTitle = normalizeWorksheetTitle(sheet?.title || '');
    if (!gid || !nextTitle) {
      return {
        gid,
        title: currentTitle || `Sheet_${gid}`,
        index: Number.isFinite(sheet?.index) ? sheet.index : index,
      };
    }
    if (nextTitle !== currentTitle) {
      diffCount += 1;
    }
    return {
      gid,
      title: nextTitle,
      index: Number.isFinite(sheet?.index) ? sheet.index : index,
    };
  });

  // Require at least one concrete diff before overriding.
  const finalSheets = diffCount > 0 ? corrected : hydrated;
  return finalSheets;
}

/**
 * Map gviz errors into plugin error codes with user-facing hints.
 *
 * 将 gviz 错误映射为插件错误码并附带可操作提示。
 */
function mapGvizError(error, extra = '') {
  const reason = String(error?.reason || '').toLowerCase();
  const message = String(error?.message || '').trim();
  if (reason.includes('access') || reason.includes('auth') || message.toLowerCase().includes('denied')) {
    throw createGoogleSheetError(
      'AUTH_REQUIRED',
      `${message || 'Access denied to Google Sheet'}${extra}`,
      'Confirm your browser session can access this sheet',
    );
  }
  throw createGoogleSheetError(
    'FETCH_FAILED',
    message || 'Google gviz endpoint returned an error',
    'Try another sheet gid or verify permissions',
  );
}

/**
 * Fetch worksheet rows by gid using layered strategies.
 * Preference order:
 * 1) CSV export in browser context;
 * 2) CSV export via Node + browser cookies;
 * 3) gviz endpoint;
 * 4) visible DOM grid fallback when access denied.
 *
 * 按 gid 获取工作表行数据，采用分层策略。
 * 优先级顺序：
 * 1）浏览器上下文 CSV 导出；
 * 2）Node + 浏览器 cookie 的 CSV 导出；
 * 3）gviz 接口；
 * 4）遇到权限拒绝时回退到可见 DOM 网格读取。
 */
export async function fetchSheetRows(page, docId, gid) {
  let csvAttempt = 'csv:not-tried';

  // First try CSV export in browser context; private sheets often allow this
  // even when gviz returns ACCESS_DENIED.
  //
  // 先尝试浏览器上下文的 CSV 导出；很多私有表格这里可读，而 gviz 会拒绝。
  if (page && typeof page.evaluate === 'function') {
    const exportUrl = buildExportCsvUrl(docId, gid);
    const exportResp = await fetchWithinBrowserContext(page, exportUrl);
    if (exportResp) {
      csvAttempt = `csv:status=${exportResp.status},ok=${Boolean(exportResp.ok)}`;
      const csvBody = String(exportResp.body || '');
      const denied = /ACCESS_DENIED|PERMISSION_DENIED|NOT_AUTHORIZED/i.test(csvBody);
      if (exportResp.ok && !denied && !looksLikeHtml(csvBody)) {
        const csvRows = parseCsvRows(csvBody);
        if (csvRows.length > 0) {
          return csvRows;
        }
        csvAttempt = `${csvAttempt},rows=0`;
      } else {
        csvAttempt = `${csvAttempt},denied=${denied},html=${looksLikeHtml(csvBody)},preview=${compactPreview(csvBody)}`;
      }
    } else {
      csvAttempt = 'csv:context-fetch-unavailable';
    }
  }

  // If browser-context fetch is unavailable, retry with Node fetch + merged cookies.
  //
  // 若浏览器上下文 fetch 不可用，则退回 Node fetch + 浏览器 cookie 合并方案。
  if (csvAttempt.includes('context-fetch-unavailable') && page && typeof page.getCookies === 'function') {
    const exportUrl = buildExportCsvUrl(docId, gid);
    const nodeCsvResp = await fetchWithNodeCookies(page, exportUrl, buildSheetEditUrl(docId, gid));
    csvAttempt = `${csvAttempt};node:status=${nodeCsvResp.status},ok=${Boolean(nodeCsvResp.ok)}`;
    const csvBody = String(nodeCsvResp.body || '');
    const denied = /ACCESS_DENIED|PERMISSION_DENIED|NOT_AUTHORIZED/i.test(csvBody);
    if (nodeCsvResp.ok && !denied && !looksLikeHtml(csvBody)) {
      const csvRows = parseCsvRows(csvBody);
      if (csvRows.length > 0) {
        return csvRows;
      }
      csvAttempt = `${csvAttempt},rows=0`;
    } else {
      csvAttempt = `${csvAttempt},denied=${denied},html=${looksLikeHtml(csvBody)},preview=${compactPreview(csvBody)}`;
    }
  }

  const url = `${buildDocUrl(docId, 'gviz/tq')}?gid=${encodeURIComponent(gid)}&tqx=out:json`;
  const response = await fetchWithSession(page, url);
  const body = response.body;

  if (response.status === 404) {
    throw createGoogleSheetError('DOC_NOT_FOUND', `Google Sheet not found: ${docId}`, 'Check docId and gid');
  }

  ensureAuthorized(response);

  if (!response.ok) {
    throw createGoogleSheetError(
      'FETCH_FAILED',
      `Google gviz request failed with HTTP ${response.status}`,
      'Try again later or verify sheet access',
    );
  }

  let payload;
  try {
    payload = parseGvizResponse(body);
  } catch {
    throw createGoogleSheetError(
      'PARSE_FAILED',
      'Failed to parse Google gviz response',
      'The sheet response format may have changed',
    );
  }

  if (payload?.status === 'error') {
    const reason = String(payload?.errors?.[0]?.reason || '').toLowerCase();
    const message = String(payload?.errors?.[0]?.message || '').trim();
    const accessDenied = reason.includes('access') || reason.includes('auth') || message.toLowerCase().includes('denied');
    if (accessDenied) {
      const domRows = await tryExtractVisibleRowsFromLoadedPage(page, docId, gid);
      if (domRows && domRows.length > 0) {
        return domRows;
      }
      mapGvizError(payload.errors?.[0], ` (${csvAttempt}; dom:unavailable)`);
    }
    mapGvizError(payload.errors?.[0], ` (${csvAttempt})`);
  }

  return parseGvizTableRows(payload);
}
