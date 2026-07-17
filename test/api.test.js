import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { buildApp } from '../src/app.js';
import { loadConfig } from '../src/config.js';

const fixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url));
const renderOptions = {
  fileName: 'basic.rdl', outputFileName: 'basic-output', output: 'PDF',
  parameters: { Title: 'Sales', Choice: 'A' },
  datasets: { Sales: [{ Name: 'North', Amount: 10 }] },
};

async function application(context, overrides = {}) {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-api-test-'));
  const config = loadConfig({ ...process.env, RDL_TEMP_ROOT: tempRoot, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000', ...overrides });
  const app = await buildApp({ config, logger: false });
  context.after(async () => { await app.close(); await fs.rm(tempRoot, { recursive: true, force: true }); });
  return { app, tempRoot };
}

function multipartBody(rdl, options) {
  const boundary = `----rdl-${Date.now()}`;
  const chunks = [
    Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="request"\r\nContent-Type: application/json\r\n\r\n${JSON.stringify(options)}\r\n`),
    Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="rdl"; filename="basic.rdl"\r\nContent-Type: application/xml\r\n\r\n`),
    rdl,
    Buffer.from(`\r\n--${boundary}--\r\n`),
  ];
  return { boundary, payload: Buffer.concat(chunks) };
}

test('health, readiness, and analysis do not expose RDL queries', async (context) => {
  const { app } = await application(context);
  assert.equal((await app.inject({ method: 'GET', url: '/healthz' })).statusCode, 200);
  assert.equal((await app.inject({ method: 'GET', url: '/readyz' })).statusCode, 200);
  const analysis = await app.inject({ method: 'POST', url: '/v1/analyze', payload: { rdlBase64: fixture.toString('base64') } });
  assert.equal(analysis.statusCode, 200);
  assert.equal(analysis.body.includes('select secret'), false);
});

test('renders equivalent PDF contracts for JSON and multipart requests and cleans temp files', async (context) => {
  const { app, tempRoot } = await application(context);
  const json = await app.inject({ method: 'POST', url: '/v1/render', payload: { ...renderOptions, rdlBase64: fixture.toString('base64') } });
  assert.equal(json.statusCode, 200, json.body);
  assert.equal(json.headers['content-type'], 'application/pdf');
  assert.equal(json.rawPayload.subarray(0, 4).toString(), '%PDF');
  assert.match(json.headers['content-disposition'], /basic-output\.pdf/);
  assert.equal(Number(json.headers['content-length']), json.rawPayload.length);
  assert.equal(Number(json.headers['x-page-count']) >= 1, true);
  assert.equal(Number(json.headers['x-render-duration-ms']) >= 0, true);
  assert.equal(typeof json.headers['x-request-id'], 'string');

  const multipart = multipartBody(fixture, renderOptions);
  const response = await app.inject({ method: 'POST', url: '/v1/render', headers: { 'content-type': `multipart/form-data; boundary=${multipart.boundary}` }, payload: multipart.payload });
  assert.equal(response.statusCode, 200, response.body);
  assert.equal(response.rawPayload.subarray(0, 4).toString(), '%PDF');
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('returns stable validation errors without leaking input', async (context) => {
  const { app } = await application(context);
  const response = await app.inject({ method: 'POST', url: '/v1/render', payload: { ...renderOptions, datasets: {}, rdlBase64: fixture.toString('base64') } });
  assert.equal(response.statusCode, 400);
  assert.equal(response.json().error.code, 'DATASET_MISSING');
  assert.equal(response.body.includes('select secret'), false);
  assert.deepEqual(await fs.readdir(app.converterConfig.tempRoot), []);
});

test('reports rejected XML paths during analysis and blocks rendering fail-closed', async (context) => {
  const { app, tempRoot } = await application(context);
  const unsupported = Buffer.from(fixture.toString().replace('<CanGrow>true</CanGrow>', '<CanGrow>true</CanGrow><ToolTip>not rendered</ToolTip>'));
  const analysis = await app.inject({ method: 'POST', url: '/v1/analyze', payload: { rdlBase64: unsupported.toString('base64') } });
  assert.equal(analysis.statusCode, 200);
  assert.equal(analysis.json().compatible, false);
  assert.equal(analysis.json().capabilities.rejected.some(({ path }) => path.endsWith('.Textbox.ToolTip')), true);

  const render = await app.inject({
    method: 'POST', url: '/v1/render', payload: { ...renderOptions, rdlBase64: unsupported.toString('base64') },
  });
  assert.equal(render.statusCode, 400);
  assert.equal(render.json().error.code, 'UNSUPPORTED_FEATURE');
  assert.equal(render.json().error.details.features.some((feature) => feature.endsWith('.Textbox.ToolTip')), true);
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('renders both DOCX modes through the public API', async (context) => {
  const { app, tempRoot } = await application(context);
  for (const output of ['DOCX_EDITABLE', 'DOCX_VISUAL']) {
    const response = await app.inject({ method: 'POST', url: '/v1/render', payload: { ...renderOptions, output, rdlBase64: fixture.toString('base64') } });
    assert.equal(response.statusCode, 200, response.body);
    assert.equal(response.headers['content-type'], 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    assert.equal(response.rawPayload.subarray(0, 2).toString(), 'PK');
    if (output === 'DOCX_VISUAL') assert.equal(Number(response.headers['x-page-count']) >= 1, true);
  }
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('rejects saturated renders with Retry-After without creating temporary files', async (context) => {
  const { app, tempRoot } = await application(context, { RDL_MAX_CONCURRENCY: '1' });
  app.renderRunner.inFlight = 1;
  const response = await app.inject({ method: 'POST', url: '/v1/render', payload: { ...renderOptions, rdlBase64: fixture.toString('base64') } });
  app.renderRunner.inFlight = 0;
  assert.equal(response.statusCode, 503);
  assert.equal(response.json().error.code, 'BUSY');
  assert.equal(response.headers['retry-after'], '5');
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('terminates timed-out and pre-aborted workers and guarantees cleanup', async (context) => {
  const { app, tempRoot } = await application(context, { RDL_RENDER_TIMEOUT_MS: '1' });
  const timedOut = await app.inject({ method: 'POST', url: '/v1/render', payload: { ...renderOptions, rdlBase64: fixture.toString('base64') } });
  assert.equal(timedOut.statusCode, 504);
  assert.equal(timedOut.json().error.code, 'RENDER_TIMEOUT');
  assert.deepEqual(await fs.readdir(tempRoot), []);

  const controller = new AbortController();
  controller.abort();
  await assert.rejects(
    app.renderRunner.render({ rdlBuffer: fixture, request: renderOptions, signal: controller.signal }),
    (error) => error.code === 'RENDER_FAILED' && error.statusCode === 499,
  );
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('enforces request, row, parameter, and exact-field limits with stable errors', async (context) => {
  const { app } = await application(context, { RDL_MAX_ROWS: '1' });
  const excessiveRows = await app.inject({ method: 'POST', url: '/v1/render', payload: {
    ...renderOptions,
    datasets: { Sales: [{ Name: 'North', Amount: 1 }, { Name: 'South', Amount: 2 }] },
    rdlBase64: fixture.toString('base64'),
  } });
  assert.equal(excessiveRows.statusCode, 413);
  assert.equal(excessiveRows.json().error.code, 'RDL_INVALID');

  const missingField = await app.inject({ method: 'POST', url: '/v1/render', payload: {
    ...renderOptions, datasets: { Sales: [{ Name: 'North' }] }, rdlBase64: fixture.toString('base64'),
  } });
  assert.equal(missingField.json().error.code, 'FIELD_MISSING');

  const missingParameter = await app.inject({ method: 'POST', url: '/v1/render', payload: {
    ...renderOptions, parameters: { Title: 'Sales' }, rdlBase64: fixture.toString('base64'),
  } });
  assert.equal(missingParameter.json().error.code, 'PARAMETER_INVALID');

  const limited = await application(context, { RDL_MAX_REQUEST_BYTES: '2048', RDL_MAX_RDL_BYTES: '1024' });
  const oversized = await limited.app.inject({ method: 'POST', url: '/v1/analyze', payload: { rdlBase64: fixture.toString('base64') } });
  assert.equal(oversized.statusCode, 413);
  assert.equal(oversized.json().error.code, 'RDL_INVALID');
});
