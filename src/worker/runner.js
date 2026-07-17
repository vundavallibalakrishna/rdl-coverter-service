import fs from 'node:fs/promises';
import path from 'node:path';
import { fork } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ServiceError } from '../errors.js';

const workerPath = fileURLToPath(new URL('./renderWorker.js', import.meta.url));

export class RenderRunner {
  constructor(config) {
    this.config = config;
    this.active = new Set();
    this.inFlight = 0;
    this.shuttingDown = false;
  }

  async render({ rdlBuffer, request, signal }) {
    if (this.shuttingDown) throw new ServiceError('BUSY', 'Service is shutting down', 503);
    if (this.inFlight >= this.config.maxConcurrency) throw new ServiceError('BUSY', 'Render capacity is currently full', 503, { retryAfterSeconds: 5 });
    this.inFlight += 1;
    let tempDir;
    let child;
    try {
      tempDir = await fs.mkdtemp(path.join(this.config.tempRoot, 'request-'));
      await fs.chmod(tempDir, 0o700);
      const rdlPath = path.join(tempDir, 'input.rdl');
      const requestPath = path.join(tempDir, 'request.json');
      await Promise.all([
        fs.writeFile(rdlPath, rdlBuffer, { mode: 0o600 }),
        fs.writeFile(requestPath, JSON.stringify(request), { mode: 0o600 }),
      ]);

      child = fork(workerPath, [], {
        stdio: ['ignore', 'ignore', 'ignore', 'ipc'],
        env: process.env,
        execArgv: [`--max-old-space-size=${this.config.workerMemoryMb}`],
      });
      this.active.add(child);
      const metadata = await new Promise((resolve, reject) => {
        let settled = false;
        const finish = (callback, value) => {
          if (settled) return;
          settled = true;
          clearTimeout(timeout);
          signal?.removeEventListener('abort', abort);
          callback(value);
        };
        const timeout = setTimeout(() => {
          child.kill('SIGKILL');
          finish(reject, new ServiceError('RENDER_TIMEOUT', 'Rendering exceeded the configured timeout', 504));
        }, this.config.renderTimeoutMs);
        const abort = () => {
          child.kill('SIGKILL');
          finish(reject, new ServiceError('RENDER_FAILED', 'Client disconnected before rendering completed', 499));
        };
        signal?.addEventListener('abort', abort, { once: true });
        if (signal?.aborted) abort();
        child.once('error', () => finish(reject, new ServiceError('RENDER_FAILED', 'Render worker could not be started', 500)));
        child.once('exit', (code, exitSignal) => {
          if (!settled && code !== 0) finish(reject, new ServiceError('RENDER_FAILED', `Render worker stopped unexpectedly (${exitSignal || code})`, 500));
        });
        child.on('message', (message) => {
          if (message?.type === 'failed') finish(reject, new ServiceError(message.error.code, message.error.message, message.error.statusCode, message.error.details));
          if (message?.type === 'completed') finish(resolve, message);
        });
        child.send({ type: 'render', tempDir, rdlPath, requestPath, config: this.config });
      });
      const buffer = await fs.readFile(metadata.outputPath);
      return { ...metadata, buffer };
    } finally {
      if (child?.connected) child.disconnect();
      if (child && !child.killed) child.kill('SIGTERM');
      if (child) this.active.delete(child);
      this.inFlight -= 1;
      if (tempDir) await fs.rm(tempDir, { recursive: true, force: true });
    }
  }

  async shutdown() {
    this.shuttingDown = true;
    for (const child of this.active) child.kill('SIGTERM');
    await Promise.allSettled([...this.active].map((child) => new Promise((resolve) => child.once('exit', resolve))));
    this.active.clear();
  }
}
