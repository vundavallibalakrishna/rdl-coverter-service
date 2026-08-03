#!/usr/bin/env node
import os from 'node:os';
import { buildApp } from './app.js';

const app = await buildApp();
const { host, port } = app.converterConfig;

async function shutdown(signal) {
  app.log.info({ signal }, 'Shutting down RDL converter');
  await app.close();
  process.exit(0);
}

process.once('SIGTERM', () => shutdown('SIGTERM'));
process.once('SIGINT', () => shutdown('SIGINT'));

try {
  await app.listen({ host, port });
  app.log.info({
    event: 'runtime.profile',
    platform: process.platform,
    osRelease: os.release(),
    architecture: process.arch,
    nodeVersion: process.version,
    availableParallelism: os.availableParallelism(),
    totalMemoryMb: Math.round(os.totalmem() / (1024 * 1024)),
    maxConcurrency: app.converterConfig.maxConcurrency,
    workerMemoryMb: app.converterConfig.workerMemoryMb,
    workerMemoryMaxMb: app.converterConfig.workerMemoryMaxMb,
    renderTimeoutMs: app.converterConfig.renderTimeoutMs,
  }, 'RDL converter runtime profile');
} catch (error) {
  app.log.error({ error: error.message }, 'Failed to start RDL converter');
  process.exit(1);
}
