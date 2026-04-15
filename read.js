import { cli, Strategy } from '@jackwener/opencli/registry';

import { createReadHandler } from './src/read-handler.js';
import { toCliError } from './src/opencli-error.js';

/**
 * Register `google-sheet read` command.
 *
 * 注册 `google-sheet read` 命令。
 */
const handler = createReadHandler();

cli({
  site: 'google-sheet',
  name: 'read',
  description: 'Read worksheet rows by docId and sheet name/gid',
  domain: 'docs.google.com',
  strategy: Strategy.COOKIE,
  browser: true,
  defaultFormat: 'json',
  args: [
    { name: 'docId', required: true, help: 'Google Sheets document ID' },
    { name: 'sheet', help: 'Worksheet selector: exact title or gid' },
  ],

  /**
   * Execute read workflow and normalize errors for OpenCLI output.
   *
   * 执行 read 读取流程，并将异常转换为 OpenCLI 统一错误输出。
   */
  func: async (page, kwargs) => {
    try {
      return await handler(page, kwargs);
    } catch (error) {
      throw toCliError(error);
    }
  },
});
