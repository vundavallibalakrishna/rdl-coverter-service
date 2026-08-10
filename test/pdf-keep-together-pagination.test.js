import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const fixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url));
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' });
const tmpRoot = path.resolve(new URL('../tmp/', import.meta.url).pathname);

function keepTogetherScenario(lines, withLeadingItem) {
  const model = parseRdl(fixture);
  model.page.height = 240;
  model.page.marginTop = 20;
  model.page.marginBottom = 20;
  model.page.header = null;
  model.page.footer = null;

  const tablix = model.body.items.find((item) => item.type === 'Tablix');
  const bodyColumns = [...tablix.columns];
  const staticHeader = structuredClone(tablix.rows[0].cells[0]);
  staticHeader.items[0].value = 'Group';
  staticHeader.items[0].paragraphs = [['Group']];
  const dynamicHeader = structuredClone(tablix.rows[1].cells[0]);
  dynamicHeader.items[0].value = '=Fields!Group.Value';
  dynamicHeader.items[0].paragraphs = [['=Fields!Group.Value']];
  const staticMember = {
    group: null,
    repeatOnNewPage: true,
    keepTogether: true,
    keepWithGroup: 'After',
    fixedData: false,
    hidden: 'false',
    sortExpressions: [],
    header: { size: 50, cell: staticHeader },
    children: [],
  };
  const dynamicMember = {
    group: { name: 'KeepGroup', expressions: ['=Fields!Group.Value'], pageBreak: 'None' },
    repeatOnNewPage: false,
    keepTogether: false,
    keepWithGroup: 'None',
    fixedData: false,
    hidden: 'false',
    sortExpressions: [],
    header: { size: 50, cell: dynamicHeader },
    children: [],
  };
  tablix.rowHeaderColumns = [50];
  tablix.columns = [50, ...bodyColumns];
  tablix.rowMembers = [staticMember, dynamicMember];
  tablix.rowMemberPaths = [[staticMember], [dynamicMember]];
  tablix.rows[1].cells.forEach((cell) => {
    const textbox = cell.items.find((item) => item.type === 'Textbox');
    if (textbox) textbox.keepTogether = true;
  });
  tablix.top = withLeadingItem ? 80 : 0;

  if (withLeadingItem) {
    const leading = structuredClone(model.body.items.find((item) => item.type === 'Textbox'));
    leading.name = 'LeadingItem';
    leading.top = 0;
    leading.left = 0;
    leading.height = 80;
    leading.value = 'LEADING_PAGE_CONTENT';
    leading.paragraphs = [['LEADING_PAGE_CONTENT']];
    model.body.items = [leading, tablix];
  } else {
    model.body.items = [tablix];
  }

  const longText = ['KEEP_ROW_START', ...Array.from({ length: lines }, (_, index) => `KEEP_LINE_${String(index + 1).padStart(3, '0')} wrapped content`), 'KEEP_ROW_END'].join('\n');
  const request = {
    outputFileName: 'keep-together-pagination',
    parameters: { Title: 'Keep together', Choice: 'A' },
    datasets: {
      Sales: [
        { Name: longText, Amount: 1, Group: 'Oversized logical group' },
        { Name: 'FOLLOWING_ROW_1', Amount: 2, Group: 'Oversized logical group' },
        { Name: 'FOLLOWING_ROW_2', Amount: 3, Group: 'Oversized logical group' },
        { Name: 'FOLLOWING_ROW_3', Amount: 4, Group: 'Oversized logical group' },
      ],
    },
  };
  return { model, request };
}

async function renderAndExtract(context, scenario, name) {
  await fs.mkdir(tmpRoot, { recursive: true });
  const pdfPath = path.join(tmpRoot, `${name}-${process.pid}.pdf`);
  context.after(() => fs.rm(pdfPath, { force: true }));
  const result = await renderPdf(scenario.model, scenario.request, config);
  await fs.writeFile(pdfPath, result.buffer);
  const all = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  return { result, pdfPath, text: all.stdout };
}

test('moves a keep-together row to a fresh page before splitting its text', async (context) => {
  const scenario = keepTogetherScenario(10, true);
  const rendered = await renderAndExtract(context, scenario, 'keep-together-row');
  const firstPage = await execFileAsync('pdftotext', ['-f', '1', '-l', '1', '-layout', rendered.pdfPath, '-']);
  const secondPage = await execFileAsync('pdftotext', ['-f', '2', '-l', '2', '-layout', rendered.pdfPath, '-']);

  assert.match(firstPage.stdout, /LEADING_PAGE_CONTENT/);
  assert.doesNotMatch(firstPage.stdout, /KEEP_ROW_START/);
  assert.match(secondPage.stdout, /KEEP_ROW_START/);
  assert.equal((rendered.text.match(/KEEP_ROW_START/g) || []).length, 1);
  assert.equal((rendered.text.match(/KEEP_ROW_END/g) || []).length, 1);
});

test('splits an oversized keep-together row when it already starts on a fresh page', async (context) => {
  const scenario = keepTogetherScenario(80, false);
  const rendered = await renderAndExtract(context, scenario, 'oversized-keep-together-row');

  assert.equal(rendered.result.pageCount > 1, true);
  assert.equal(rendered.result.pageCount < 20, true);
  assert.equal((rendered.text.match(/KEEP_ROW_START/g) || []).length, 1);
  assert.equal((rendered.text.match(/KEEP_ROW_END/g) || []).length, 1);
  for (let index = 1; index <= 80; index += 1) {
    const marker = `KEEP_LINE_${String(index).padStart(3, '0')}`;
    assert.equal((rendered.text.match(new RegExp(marker, 'g')) || []).length, 1);
  }
});

test('moves a multi-row repeatable header with its first data row instead of replaying part of it', async (context) => {
  const model = parseRdl(fixture);
  model.page.height = 240;
  model.page.marginTop = 20;
  model.page.marginBottom = 20;
  model.page.header = null;
  model.page.footer = null;

  const tablix = model.body.items.find((item) => item.type === 'Tablix');
  const firstHeader = tablix.rows[0];
  const secondHeader = structuredClone(firstHeader);
  const detail = tablix.rows[1];
  firstHeader.height = 24;
  secondHeader.height = 24;
  detail.height = 30;
  const setText = (row, value) => {
    const textbox = row.cells[0].items.find((item) => item.type === 'Textbox');
    textbox.value = value;
    textbox.paragraphs = [[{ value, markupType: 'None', style: textbox.style }]];
  };
  setText(firstHeader, 'HEADER_ONE');
  setText(secondHeader, 'HEADER_TWO');

  const firstStaticMember = tablix.rowMembers[0];
  const secondStaticMember = structuredClone(firstStaticMember);
  const dynamicMember = tablix.rowMembers[1];
  firstStaticMember.repeatOnNewPage = true;
  firstStaticMember.keepWithGroup = 'After';
  secondStaticMember.repeatOnNewPage = true;
  secondStaticMember.keepWithGroup = 'After';
  tablix.rows = [firstHeader, secondHeader, detail];
  tablix.rowMembers = [firstStaticMember, secondStaticMember, dynamicMember];
  tablix.rowMemberPaths = [[firstStaticMember], [secondStaticMember], [dynamicMember]];
  tablix.top = 160;

  const leading = structuredClone(model.body.items.find((item) => item.type === 'Textbox'));
  leading.name = 'LeadingItem';
  leading.top = 0;
  leading.left = 0;
  leading.height = 160;
  leading.value = 'LEADING_PAGE_CONTENT';
  leading.paragraphs = [[{ value: 'LEADING_PAGE_CONTENT', markupType: 'None', style: leading.style }]];
  model.body.items = [leading, tablix];

  const scenario = {
    model,
    request: {
      outputFileName: 'header-block-pagination',
      parameters: { Title: 'Header block', Choice: 'A' },
      datasets: { Sales: [{ Name: 'FIRST_DATA_ROW', Amount: 1 }] },
    },
  };
  const rendered = await renderAndExtract(context, scenario, 'header-block-pagination');
  const firstPage = await execFileAsync('pdftotext', ['-f', '1', '-l', '1', '-layout', rendered.pdfPath, '-']);
  const secondPage = await execFileAsync('pdftotext', ['-f', '2', '-l', '2', '-layout', rendered.pdfPath, '-']);

  assert.match(firstPage.stdout, /LEADING_PAGE_CONTENT/);
  assert.doesNotMatch(firstPage.stdout, /HEADER_ONE|HEADER_TWO|FIRST_DATA_ROW/,
    'no orphaned header row may remain on the preceding page');
  assert.equal((secondPage.stdout.match(/HEADER_ONE/g) || []).length, 1);
  assert.equal((secondPage.stdout.match(/HEADER_TWO/g) || []).length, 1);
  assert.equal((secondPage.stdout.match(/FIRST_DATA_ROW/g) || []).length, 1);
});
