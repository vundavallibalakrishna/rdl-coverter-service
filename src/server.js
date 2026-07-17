#!/usr/bin/env node
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
} catch (error) {
  app.log.error({ error: error.message }, 'Failed to start RDL converter');
  process.exit(1);
}
