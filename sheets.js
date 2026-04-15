import { cli, Strategy } from '@jackwener/opencli/registry';

import { createSheetsHandler } from './src/sheets-handler.js';
import { toCliError } from './src/opencli-error.js';

/**
 * Register `google-sheet sheets` command.
 *
 * 注册 `google-sheet sheets` 命令。
 */
const handler = createSheetsHandler();

cli({
  site: 'google-sheet',
  name: 'sheets',
  description: 'List worksheets by Google Sheets docId',
  domain: 'docs.google.com',
  strategy: Strategy.COOKIE,
  browser: true,
  args: [
    { name: 'docId', required: true, help: 'Google Sheets document ID' },
  ],
  columns: ['gid', 'title', 'index'],

  /**
   * Execute sheet metadata listing and normalize thrown errors.
   *
   * 执行工作表元数据列表读取，并统一转换抛出的异常。
   */
  func: async (page, kwargs) => {
    try {
      return await handler(page, kwargs);
    } catch (error) {
      throw toCliError(error);
    }
  },
});
