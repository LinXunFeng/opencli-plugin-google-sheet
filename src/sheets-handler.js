import { fetchWorkbookMeta } from './sheets-api.js';

/**
 * Create sheets command handler with injectable dependencies.
 *
 * 创建 sheets 命令处理器，并支持依赖注入。
 */
export function createSheetsHandler(deps = {}) {
  const getWorkbookMeta = deps.fetchWorkbookMeta ?? fetchWorkbookMeta;

  /**
   * Return worksheet metadata list for a workbook.
   *
   * 返回工作簿的工作表元数据列表。
   */
  return async function sheetsHandler(page, kwargs) {
    return getWorkbookMeta(page, kwargs.docId);
  };
}
