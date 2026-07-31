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

test('serves the open end-to-end render page with isolated same-origin browser behavior', async (context) => {
  const { app } = await application(context);
  for (const url of ['/', '/test-ui']) {
    const response = await app.inject({ method: 'GET', url });
    assert.equal(response.statusCode, 200);
    assert.match(response.headers['content-type'], /^text\/html/);
    assert.equal(response.headers['cache-control'], 'no-store');
    assert.match(response.headers['content-security-policy'], /connect-src 'self'/);
    assert.match(response.headers['content-security-policy'], /frame-ancestors 'none'/);
    assert.match(response.body, /id="rdl-file"/);
    assert.match(response.body, /id="json-file"/);
    assert.match(response.body, /DOCX_EDITABLE/);
    assert.match(response.body, /DOCX_VISUAL/);
    assert.equal(response.body.includes("fetch('/v1/render'"), true);
    assert.equal(response.body.includes('authorization'), false);
  }
});

test('health, readiness, and analysis do not expose RDL queries', async (context) => {
  const { app } = await application(context);
  assert.equal((await app.inject({ method: 'GET', url: '/healthz' })).statusCode, 200);
  assert.equal((await app.inject({ method: 'GET', url: '/readyz' })).statusCode, 200);
  const analysis = await app.inject({ method: 'POST', url: '/v1/analyze', payload: { rdlBase64: fixture.toString('base64') } });
  assert.equal(analysis.statusCode, 200);
  assert.equal(analysis.body.includes('select secret'), false);
  assert.match(analysis.json().identity.definitionSha256, /^[a-f0-9]{64}$/);
  assert.equal(analysis.json().windowsWordEditable.layoutMode, 'windows-paged-editable');
  assert.equal(analysis.json().windowsWordEditable.platform, 'Microsoft Word for Windows');
  assert.equal(analysis.json().windowsWordEditable.pdfLayoutAuthority, true);
  assert.equal(analysis.json().windowsWordEditable.page.eligible, true);
  assert.equal(analysis.json().windowsWordEditable.tableGrid.maximumColumns, 63);
  assert.equal(Array.isArray(analysis.json().windowsWordEditable.fontEmbedding), true);
});

test('public analyze and render endpoints accept the native Code.CalculateColor compatibility mapping', async (context) => {
  const { app } = await application(context);
  const mapped = Buffer.from(fixture.toString()
    .replace('<ReportSections>', '<Code>Public Function CalculateColor(y, x) As String : Return \"ignored\" : End Function</Code><ReportSections>')
    .replace('<BackgroundColor>#dddddd</BackgroundColor>', '<BackgroundColor>=Code.CalculateColor(CStr(1), CStr(1))</BackgroundColor>'));
  const analysis = await app.inject({
    method: 'POST',
    url: '/v1/analyze',
    payload: { rdlBase64: mapped.toString('base64') },
  });
  assert.equal(analysis.statusCode, 200, analysis.body);
  assert.ok(analysis.json().capabilities.expressions.detected
    .some((entry) => entry.name === 'Code.CalculateColor' && entry.status === 'SUPPORTED'));

  const rendered = await app.inject({
    method: 'POST',
    url: '/v1/render',
    payload: { ...renderOptions, rdlBase64: mapped.toString('base64') },
  });
  assert.equal(rendered.statusCode, 200, rendered.body);
  assert.equal(rendered.rawPayload.subarray(0, 4).toString(), '%PDF');
});

test('obsolete DOCX pagination flags and profiles are rejected explicitly', async (context) => {
  const { app } = await application(context);
  const obsoleteAnalysis = await app.inject({
    method: 'POST',
    url: '/v1/analyze',
    payload: { rdlBase64: fixture.toString('base64'), docx: { nativePageFragments: false } },
  });
  assert.equal(obsoleteAnalysis.statusCode, 400);
  assert.equal(obsoleteAnalysis.json().error.code, 'RDL_INVALID');
  assert.deepEqual(obsoleteAnalysis.json().error.details.obsolete, ['docx.nativePageFragments']);

  const obsoleteRender = await app.inject({
    method: 'POST',
    url: '/v1/render',
    payload: {
      ...renderOptions,
      output: 'DOCX_EDITABLE',
      rdlBase64: fixture.toString('base64'),
      docxProfile: 'removed-profile',
    },
  });
  assert.equal(obsoleteRender.statusCode, 400);
  assert.equal(obsoleteRender.json().error.code, 'RDL_INVALID');
  assert.deepEqual(obsoleteRender.json().error.details.obsolete, ['docxProfile']);
});

test('removed DOCX profile environment settings do not enter the runtime config', () => {
  const config = loadConfig({
    ...process.env,
    RDL_DOCX_NATIVE_PAGE_FRAGMENTS: 'false',
    RDL_DOCX_PROFILE_PATH: '/tmp/obsolete.json',
    RDL_DOCX_PROFILE_AUTO: 'true',
  });
  assert.equal('docxNativePageFragments' in config, false);
  assert.equal('docxProfilePath' in config, false);
  assert.equal('docxProfileAuto' in config, false);
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
  const unsupported = Buffer.from(fixture.toString().replace('<CanGrow>true</CanGrow>', '<CanGrow>true</CanGrow><UnsupportedTextboxProp>x</UnsupportedTextboxProp>'));
  const analysis = await app.inject({ method: 'POST', url: '/v1/analyze', payload: { rdlBase64: unsupported.toString('base64') } });
  assert.equal(analysis.statusCode, 200);
  assert.equal(analysis.json().compatible, false);
  assert.equal(analysis.json().capabilities.rejected.some(({ path }) => path.endsWith(".Textbox.UnsupportedTextboxProp")), true);

  const render = await app.inject({
    method: 'POST', url: '/v1/render', payload: { ...renderOptions, rdlBase64: unsupported.toString('base64') },
  });
  assert.equal(render.statusCode, 400);
  assert.equal(render.json().error.code, 'UNSUPPORTED_FEATURE');
  assert.equal(render.json().error.details.features.some((feature) => feature.endsWith(".Textbox.UnsupportedTextboxProp")), true);
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('renders all DOCX modes through the public API with explicit layout and editability headers', async (context) => {
  const { app, tempRoot } = await application(context);
  const expected = {
    DOCX_EDITABLE: { layout: 'windows-paged-editable', ratio: '1', numericPages: true },
    DOCX_VISUAL: { layout: 'visual', ratio: '0', numericPages: true },
  };
  for (const output of Object.keys(expected)) {
    const response = await app.inject({ method: 'POST', url: '/v1/render', payload: { ...renderOptions, output, rdlBase64: fixture.toString('base64') } });
    assert.equal(response.statusCode, 200, response.body);
    assert.equal(response.headers['content-type'], 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    assert.equal(response.rawPayload.subarray(0, 2).toString(), 'PK');
    assert.equal(response.headers['x-docx-layout-mode'], expected[output].layout);
    assert.equal(response.headers['x-docx-editable-text-ratio'], expected[output].ratio);
    if (output === 'DOCX_EDITABLE') {
      assert.equal(response.headers['x-docx-native-page-fragments'], undefined);
      assert.equal(response.headers['x-docx-profile-id'], undefined);
    }
    if (expected[output].numericPages) assert.equal(Number(response.headers['x-page-count']) >= 1, true);
    else assert.equal(response.headers['x-page-count'], 'unknown');
  }
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('renders XLSX with an Excel-only layout header and DATA compatibility modes', async (context) => {
  const { app, tempRoot } = await application(context);
  const cases = [
    { excel: undefined, expected: 'report-sections' },
    { excel: { layoutMode: 'data' }, expected: 'data-stacked' },
    { excel: { sheetPerTablix: true }, expected: 'data-per-tablix' },
  ];
  for (const entry of cases) {
    const response = await app.inject({
      method: 'POST', url: '/v1/render',
      payload: { ...renderOptions, output: 'XLSX', excel: entry.excel, rdlBase64: fixture.toString('base64') },
    });
    assert.equal(response.statusCode, 200, response.body);
    assert.equal(response.headers['x-xlsx-layout-mode'], entry.expected);
    assert.equal(response.headers['x-docx-layout-mode'], undefined);
    assert.equal(response.headers['x-page-count'], 'unknown');
    assert.equal(response.rawPayload.subarray(0, 2).toString(), 'PK');
  }
  const conflict = await app.inject({
    method: 'POST', url: '/v1/render',
    payload: { ...renderOptions, output: 'XLSX', excel: { layoutMode: 'REPORT', sheetPerTablix: true }, rdlBase64: fixture.toString('base64') },
  });
  assert.equal(conflict.statusCode, 400);
  assert.equal(conflict.json().error.code, 'RDL_INVALID');
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('renders page-locked DOCX without obsolete profile or fragmentation headers', async (context) => {
  const { app, tempRoot } = await application(context);
  const response = await app.inject({
    method: 'POST',
    url: '/v1/render',
    payload: { ...renderOptions, output: 'DOCX_EDITABLE', rdlBase64: fixture.toString('base64') },
  });
  assert.equal(response.statusCode, 200, response.body);
  assert.equal(response.headers['x-docx-layout-mode'], 'windows-paged-editable');
  assert.equal(Number(response.headers['x-page-count']) >= 1, true);
  assert.equal(response.headers['x-docx-profile-id'], undefined);
  assert.equal(response.headers['x-docx-profile-certified'], undefined);
  assert.equal(response.headers['x-docx-native-page-fragments'], undefined);
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
