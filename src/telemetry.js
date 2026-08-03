import { performance } from 'node:perf_hooks';

const MAX_STRING_LENGTH = 80;
const MAX_ARRAY_ITEMS = 32;
const MAX_OBJECT_KEYS = 64;

function round(value) {
  return Math.round(value * 100) / 100;
}

// Telemetry crosses the worker IPC boundary and is written to production logs. Keep it deliberately
// structural: bounded scalar values and small metric objects only. This prevents a future caller from
// accidentally attaching RDL XML, dataset values, paths, or output bytes to a phase event.
export function sanitizeTelemetryValue(value, depth = 0) {
  if (depth > 3 || value === null || value === undefined) return undefined;
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') return value.replace(/[\r\n\t]/g, ' ').slice(0, MAX_STRING_LENGTH);
  if (Array.isArray(value)) {
    return value.slice(0, MAX_ARRAY_ITEMS)
      .map((entry) => sanitizeTelemetryValue(entry, depth + 1))
      .filter((entry) => entry !== undefined);
  }
  if (typeof value === 'object') {
    const result = {};
    for (const [key, entry] of Object.entries(value).slice(0, MAX_OBJECT_KEYS)) {
      const safe = sanitizeTelemetryValue(entry, depth + 1);
      if (safe !== undefined) result[String(key).slice(0, MAX_STRING_LENGTH)] = safe;
    }
    return result;
  }
  return undefined;
}

export function processTelemetryMetrics(startedCpu, startedEventLoop) {
  const memory = process.memoryUsage();
  const cpu = process.cpuUsage(startedCpu);
  const eventLoop = performance.eventLoopUtilization(startedEventLoop);
  return {
    pid: process.pid,
    rssMb: round(memory.rss / 1024 / 1024),
    heapUsedMb: round(memory.heapUsed / 1024 / 1024),
    heapTotalMb: round(memory.heapTotal / 1024 / 1024),
    externalMb: round(memory.external / 1024 / 1024),
    cpuUserMs: round(cpu.user / 1000),
    cpuSystemMs: round(cpu.system / 1000),
    eventLoopUtilization: round(eventLoop.utilization),
  };
}

export function createTelemetryClock(source, emit = () => {}) {
  const startedAt = performance.now();
  let previousAt = startedAt;
  const startedCpu = process.cpuUsage();
  const startedEventLoop = performance.eventLoopUtilization();
  return {
    mark(phase, status = 'completed', metrics = {}) {
      const now = performance.now();
      const event = sanitizeTelemetryValue({
        source,
        phase,
        status,
        phaseDurationMs: round(now - previousAt),
        totalDurationMs: round(now - startedAt),
        ...processTelemetryMetrics(startedCpu, startedEventLoop),
        metrics,
      });
      previousAt = now;
      try {
        emit(event);
      } catch {
        // Observability must never change render success or failure semantics.
      }
      return event;
    },
  };
}
