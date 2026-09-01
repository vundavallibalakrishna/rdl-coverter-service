import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import test from 'node:test';
import JSZip from 'jszip';
import PDFDocument from 'pdfkit';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { loadConfig } from '../src/config.js';
import { textForItem } from '../src/render/common.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { pdfFont } from '../src/render/fonts.js';
import { measureTextboxHeight, renderPdf } from '../src/render/pdf.js';

const fixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url), 'utf8');
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = {
  parameters: { Title: 'SHOULD_NOT_BE_USED', Choice: 'A' },
  datasets: { Sales: [{ Name: 'North', Amount: 1 }], Choices: [{ Value: 'A' }] },
  outputFileName: 'paragraph-layout',
};

function paragraphRdl(value = '=Parameters!Title.Value', evaluationMode = 'Constant') {
  return fixture
    .replace(
      '<Value>=Parameters!Title.Value</Value>',
      `<Value EvaluationMode="${evaluationMode}">${value}</Value>`,
    )
    .replace(
      '</TextRuns></Paragraph></Paragraphs>',
      '</TextRuns><SpaceBefore>=IIF(1=1,"3pt","0pt")</SpaceBefore><SpaceAfter>10pt</SpaceAfter><Style><LineHeight>=IIF(1=1,"20pt","10pt")</LineHeight></Style></Paragraph></Paragraphs>',
    );
}

test('TextRun EvaluationMode=Constant preserves a leading equals sign without executing it', async () => {
  const rdl = paragraphRdl('=Code.MustNeverRun()');
  const analysis = analyzeRdl(rdl);
  assert.equal(analysis.compatible, true);
  assert.equal(analysis.blockingErrors.some((entry) => /CustomCode|MustNeverRun/.test(entry.feature)), false);

  const model = parseRdl(rdl);
  const title = model.body.items.find((item) => item.name === 'TitleBox');
  assert.equal(title.paragraphs[0][0].evaluationMode, 'Constant');
  assert.equal(textForItem(title, { parameters: request.parameters, globals: {}, fields: {} }), '=Code.MustNeverRun()');
  assert.equal((await renderPdf(model, request, config)).buffer.subarray(0, 4).toString(), '%PDF');
});

test('invalid TextRun EvaluationMode fails closed', () => {
  assert.throws(
    () => parseRdl(paragraphRdl('literal', 'Sometimes')),
    (error) => error?.code === 'RDL_INVALID' && /EvaluationMode/.test(error.message),
  );
});

test('paragraph SpaceBefore, SpaceAfter, and LineHeight are normalized and rendered natively', async () => {
  const rdl = paragraphRdl();
  const analysis = analyzeRdl(rdl);
  assert.equal(analysis.compatible, true);
  const used = new Map(analysis.capabilities.expressions.properties.map((entry) => [entry.property, entry]));
  assert.equal(used.get('Paragraph.SpaceBefore')?.handled, true);
  assert.equal(used.get('Style.LineHeight')?.handled, true);

  const model = parseRdl(rdl);
  const title = model.body.items.find((item) => item.name === 'TitleBox');
  assert.equal(title.paragraphStyles[0].spaceAfter, 10);
  assert.equal(title.paragraphStyles[0].spaceBefore, '=IIF(1=1,"3pt","0pt")');
  assert.equal(title.paragraphStyles[0].lineHeight, '=IIF(1=1,"20pt","10pt")');

  const measurementDoc = new PDFDocument({ autoFirstPage: false });
  measurementDoc.on('data', () => {});
  measurementDoc.addPage();
  const height = measureTextboxHeight(
    measurementDoc,
    config,
    title,
    { parameters: request.parameters, globals: {}, fields: {} },
    '=Parameters!Title.Value',
    title.width,
  );
  measurementDoc.end();
  assert.ok(height >= 33, `expected 20pt line + 3pt before + 10pt after, got ${height}`);

  const docx = await renderEditableDocx(model, request, config);
  const xml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  assert.match(xml, /<w:spacing[^>]*w:before="60"/);
  // The 10pt paragraph SpaceAfter remains native and the textbox's 2pt bottom padding is represented as
  // trailing paragraph space so Word does not add tcMar/bottom to the exact row height.
  assert.match(xml, /<w:spacing[^>]*w:after="240"/);
  assert.match(xml, /<w:spacing[^>]*w:line="400"/);
  assert.match(xml, /<w:spacing[^>]*w:lineRule="exact"/);
});

test('embedded TextRun newlines receive paragraph spacing once in canonical PDF and editable Word', async () => {
  const model = parseRdl(paragraphRdl('Line 1\nLine 2', 'Constant'));
  const title = model.body.items.find((item) => item.name === 'TitleBox');
  const context = { parameters: request.parameters, globals: {}, fields: {} };
  const text = textForItem(title, context);
  assert.equal(text, 'Line 1\nLine 2');

  const measurementDoc = new PDFDocument({ autoFirstPage: false });
  measurementDoc.on('data', () => {});
  measurementDoc.addPage();
  const height = measureTextboxHeight(measurementDoc, config, title, context, text, title.width);
  measurementDoc.end();
  // 3pt SpaceBefore + two 20pt lines + one 10pt SpaceAfter. The defect charged SpaceAfter after both
  // embedded lines and measured 63pt.
  assert.ok(height >= 52.75 && height <= 53.25, `expected one paragraph spacing contribution, got ${height}`);

  const docx = await renderEditableDocx(model, request, config);
  const xml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  assert.match(xml, />Line<\/w:t>/);
  assert.match(xml, />1<\/w:t>/);
  assert.match(xml, />2<\/w:t>/);
  assert.match(xml, /<w:br\/>/);
});

test('implicit text line height excludes a font external leading gap, while explicit RDL LineHeight remains authoritative', () => {
  const model = parseRdl(fixture);
  const title = model.body.items.find((item) => item.name === 'TitleBox');
  title.style = {
    ...title.style,
    fontFamily: 'Arial',
    fontSize: 11,
    lineHeight: null,
    paddingTop: 0,
    paddingBottom: 0,
    paddingLeft: 0,
    paddingRight: 0,
  };
  title.paragraphStyles = [{ ...title.style, spaceBefore: 0, spaceAfter: 0 }];
  title.paragraphs = [[{ value: 'First line\nSecond line', markupType: 'None', style: title.style }]];
  const context = { parameters: request.parameters, globals: {}, fields: {} };
  const measurementDoc = new PDFDocument({ autoFirstPage: false });
  measurementDoc.on('data', () => {});
  measurementDoc.addPage();
  measurementDoc.font(pdfFont(config, 'Arial', false, false, 'First line')).fontSize(11);
  const ssrsDefaultLineHeight = measurementDoc.currentLineHeight();
  const measured = measureTextboxHeight(measurementDoc, config, title, context, 'First line\nSecond line', title.width);
  measurementDoc.end();

  assert.ok(
    Math.abs(measured - (2 * ssrsDefaultLineHeight)) < 0.01,
    `expected two SSRS default line boxes (${2 * ssrsDefaultLineHeight}), got ${measured}`,
  );
});
