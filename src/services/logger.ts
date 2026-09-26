import { appendFileSync, mkdirSync } from 'node:fs';
import path from 'node:path';

const sensitiveField = /token|secret|credential|password|authorization/i;
const credentialValue = /\bBearer\s+[^\s,;]+|\b((?:access|refresh|oauth)?[_-]?token|client[_-]?secret|password|authorization(?:code)?)\s*[:=]\s*(?:Bearer\s+)?["']?[^\s,"';]+/gi;

function redactString(value: string): string {
  return value.replace(credentialValue, (_match, key: string | undefined) => key ? `${key}=[REDACTED]` : 'Bearer [REDACTED]');
}

function redact(value: unknown): unknown {
  if (typeof value === 'string') return redactString(value);
  if (value instanceof Error) {
    return {
      name: value.name,
      message: redactString(value.message),
      stack: value.stack && redactString(value.stack),
    };
  }
  if (Array.isArray(value)) return value.map(redact);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, nested]) => [
      key,
      sensitiveField.test(key) ? '[REDACTED]' : redact(nested),
    ]));
  }
  return value;
}

export interface LoggerOptions {
  directory: string;
  now?: () => Date;
}

export function createLogger(options: LoggerOptions) {
  const logFile = path.join(options.directory, 'application.log');
  mkdirSync(options.directory, { recursive: true });
  const now = options.now ?? (() => new Date());

  function write(level: 'info' | 'error', event: string, fields: Record<string, unknown>): void {
    const sanitizedFields = redact(fields) as Record<string, unknown>;
    const line = JSON.stringify({ timestamp: now().toISOString(), level, event: redactString(event), ...sanitizedFields });
    appendFileSync(logFile, `${line}\n`, { encoding: 'utf8', mode: 0o600 });
  }

  return {
    info(event: string, fields: Record<string, unknown> = {}): void {
      write('info', event, fields);
    },
    error(event: string, error: unknown, fields: Record<string, unknown> = {}): void {
      write('error', event, { ...fields, error });
    },
  };
}
