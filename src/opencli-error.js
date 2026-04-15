import { CliError, EXIT_CODES } from '@jackwener/opencli/errors';

import { GoogleSheetError } from './error.js';

/**
 * Map plugin error code to OpenCLI process exit code.
 *
 * 将插件错误码映射到 OpenCLI 进程退出码。
 */
function resolveExitCode(code) {
  if (code === 'AUTH_REQUIRED') {
    return EXIT_CODES.NOPERM;
  }
  if (code === 'SHEET_NOT_FOUND' || code === 'DOC_NOT_FOUND') {
    return EXIT_CODES.EMPTY_RESULT;
  }
  return EXIT_CODES.GENERIC_ERROR;
}

/**
 * Normalize unknown errors into OpenCLI CliError shape.
 *
 * 将各种异常统一转换为 OpenCLI 的 CliError 结构。
 */
export function toCliError(error) {
  if (error instanceof CliError) {
    return error;
  }

  if (error instanceof GoogleSheetError) {
    return new CliError(error.code, error.message, error.hint, resolveExitCode(error.code));
  }

  if (error instanceof Error) {
    return new CliError('FETCH_FAILED', error.message, 'Run with -v for debug details');
  }

  return new CliError('FETCH_FAILED', String(error), 'Run with -v for debug details');
}
