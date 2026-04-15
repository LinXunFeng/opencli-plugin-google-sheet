/**
 * Domain error type for Google Sheet operations.
 *
 * Google Sheet 操作的领域错误类型。
 */
export class GoogleSheetError extends Error {
  constructor(code, message, hint) {
    super(message);
    this.name = 'GoogleSheetError';
    this.code = code;
    this.hint = hint;
  }
}

/**
 * Create a typed GoogleSheetError instance.
 *
 * 创建带类型信息的 GoogleSheetError 实例。
 */
export function createGoogleSheetError(code, message, hint) {
  return new GoogleSheetError(code, message, hint);
}
