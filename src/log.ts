type Level = 'debug' | 'info' | 'warn' | 'error';
const ORDER: Record<Level, number> = { debug: 10, info: 20, warn: 30, error: 40 };

function current(): Level {
  const v = (process.env.LOG_LEVEL ?? 'info').toLowerCase();
  return v in ORDER ? (v as Level) : 'info';
}

function emit(level: Level, args: unknown[]): void {
  if (ORDER[level] < ORDER[current()]) return;
  // stdout is reserved for the MCP protocol: log to stderr only.
  process.stderr.write(`[outlook-mcp] ${level.toUpperCase()} ${args.map(String).join(' ')}\n`);
}

export const log = {
  debug: (...a: unknown[]) => emit('debug', a),
  info: (...a: unknown[]) => emit('info', a),
  warn: (...a: unknown[]) => emit('warn', a),
  error: (...a: unknown[]) => emit('error', a),
};
