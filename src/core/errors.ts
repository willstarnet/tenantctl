/**
 * Every failure the CLI reports maps to one code and one exit status, so agents can branch on either.
 */
export const EXIT = {
  ok: 0,
  error: 1,
  usage: 2,
  confirmationRequired: 3,
  auth: 4,
  forbidden: 5,
  notFound: 6,
  writesDisabled: 7,
  ambiguous: 8,
  throttled: 9,
  denied: 10,
} as const;

export type ErrorCode =
  | 'USAGE'
  | 'CONFIG'
  | 'AUTH_REQUIRED'
  | 'AUTH_FAILED'
  | 'FORBIDDEN'
  | 'NOT_FOUND'
  | 'AMBIGUOUS'
  | 'WRITES_DISABLED'
  | 'CONFIRM_INVALID'
  | 'THROTTLED'
  | 'DENIED'
  | 'GRAPH_ERROR'
  | 'INTERNAL';

const EXIT_BY_CODE: Record<ErrorCode, number> = {
  USAGE: EXIT.usage,
  CONFIG: EXIT.usage,
  AUTH_REQUIRED: EXIT.auth,
  AUTH_FAILED: EXIT.auth,
  FORBIDDEN: EXIT.forbidden,
  NOT_FOUND: EXIT.notFound,
  AMBIGUOUS: EXIT.ambiguous,
  WRITES_DISABLED: EXIT.writesDisabled,
  CONFIRM_INVALID: EXIT.usage,
  THROTTLED: EXIT.throttled,
  DENIED: EXIT.denied,
  GRAPH_ERROR: EXIT.error,
  INTERNAL: EXIT.error,
};

export interface CliErrorOptions {
  hint?: string;
  status?: number;
  graphCode?: string;
  requestId?: string;
  /** Structured extras, e.g. candidate matches for AMBIGUOUS. */
  details?: unknown;
}

export class CliError extends Error {
  readonly code: ErrorCode;
  readonly hint?: string;
  readonly status?: number;
  readonly graphCode?: string;
  readonly requestId?: string;
  readonly details?: unknown;

  constructor(code: ErrorCode, message: string, options: CliErrorOptions = {}) {
    super(message);
    this.name = 'CliError';
    this.code = code;
    this.hint = options.hint;
    this.status = options.status;
    this.graphCode = options.graphCode;
    this.requestId = options.requestId;
    this.details = options.details;
  }

  get exitCode(): number {
    return EXIT_BY_CODE[this.code];
  }

  toJSON(): Record<string, unknown> {
    return {
      code: this.code,
      message: this.message,
      ...(this.hint ? { hint: this.hint } : {}),
      ...(this.status ? { status: this.status } : {}),
      ...(this.graphCode ? { graphCode: this.graphCode } : {}),
      ...(this.requestId ? { requestId: this.requestId } : {}),
      ...(this.details !== undefined ? { details: this.details } : {}),
    };
  }
}

export function toCliError(error: unknown): CliError {
  if (error instanceof CliError) return error;
  const message = error instanceof Error ? error.message : String(error);
  return new CliError('INTERNAL', message);
}
