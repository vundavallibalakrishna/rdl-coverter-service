import assert from 'node:assert/strict';
import test from 'node:test';
import { classifyWorkerFatalStderr } from '../src/worker/runner.js';

test('worker fatal diagnostics distinguish V8 heap exhaustion from other aborts', () => {
  assert.equal(classifyWorkerFatalStderr(
    'FATAL ERROR: Reached heap limit Allocation failed - JavaScript heap out of memory',
    134,
  ), 'V8_HEAP_OUT_OF_MEMORY');
  assert.equal(classifyWorkerFatalStderr('Check failed: isolate != nullptr', 134), 'NATIVE_RUNTIME_ABORT');
  assert.equal(classifyWorkerFatalStderr('', 134), 'PROCESS_ABORT');
  assert.equal(classifyWorkerFatalStderr('', 1), 'WORKER_EXIT');
});

test('worker fatal diagnostics return only stable categories and never echo stderr content', () => {
  const sensitive = 'C:\\private\\request-123\\input.rdl customer-secret-value';
  const category = classifyWorkerFatalStderr(sensitive, 134);
  assert.equal(category, 'PROCESS_ABORT');
  assert.equal(category.includes('private'), false);
  assert.equal(category.includes('customer-secret-value'), false);
});
