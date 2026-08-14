// Two hard rules the service enforces regardless of what the RDL declares:
//  1. A tablix's static column-header rows repeat on every page.
//  2. The last row of a bordered data tablix is closed; borderless static or dynamic layout tablixes honor
//     Border=None.
// For a tablix that uses vertically-merged (rowSpan) cells — where Word disables native repeat-header — the
// header is repeated by physically redrawing it per page (page-fragment mode). These tests assert the model
// flagging, the shared border helper, and the emitted OpenXML on synthetic RDLs isolating each construct.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { PNG } from 'pngjs';
import { parseRdl } from '../src/rdl/parser.js';
import { tablixRows, enforcedBottomBorder, shouldEnforceTablixBottom } from '../src/render/common.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderVisualDocx } from '../src/render/visualDocx.js';
import { loadConfig } from '../src/config.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const documentXml = async (buffer) => (await JSZip.loadAsync(buffer)).file('word/document.xml').async('string');
const tmpRoot = path.resolve(new URL('../tmp/', import.meta.url).pathname);

// A tablix with one static column-header row and a dynamic detail row grouped by V (no merged cells).
const flatTablixRdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="V"><DataField>V</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody><TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
   <TablixRows>
    <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents><Textbox Name="h"><Paragraphs><Paragraph><TextRuns><TextRun><Value>COLHDR</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
    <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents><Textbox Name="d"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!V.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
   </TablixRows></TablixBody>
   <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
   <TablixRowHierarchy><TablixMembers><TablixMember/><TablixMember><Group Name="g"><GroupExpressions><GroupExpression>=Fields!V.Value</GroupExpression></GroupExpressions></Group></TablixMember></TablixMembers></TablixRowHierarchy>
   <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Width>3in</Width></Tablix>
 </ReportItems><Height>10in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;

const request = (rowCount) => ({ parameters: {}, datasets: { D: Array.from({ length: rowCount }, (_, i) => ({ V: `Row ${i}` })) } });

test('a tablix static column-header row is marked to repeat, even without RepeatColumnHeaders in the RDL', () => {
  const m = parseRdl(flatTablixRdl);
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  const { rows } = tablixRows(tablix, request(5), { PageNumber: 1, TotalPages: 1 }, m);
  assert.equal(rows[0].isHeader, true); // the column header repeats
  assert.equal(rows[1].isHeader, false); // the detail rows do not
});

test('the page-locked DOCX materializes the column header as native text in its page table', async () => {
  const m = parseRdl(flatTablixRdl);
  const xml = await documentXml((await renderEditableDocx(m, request(5), config)).buffer);
  assert.match(xml, /COLHDR/);
  assert.match(xml, /<w:tblLayout w:type="fixed"\/>/);
  assert.match(xml, /<w:trHeight[^>]*w:hRule="exact"/);
});

test('enforcedBottomBorder keeps a declared visible bottom, else synthesizes one', () => {
  const declared = { style: 'Solid', color: '#000000', width: 1 };
  assert.deepEqual(enforcedBottomBorder({ borders: { bottom: declared } }), declared);
  // None bottom -> synthesized Solid, matching an existing solid side's colour/width.
  const synth = enforcedBottomBorder({ borders: { bottom: { style: 'None' }, left: { style: 'Solid', color: '#123456', width: 2 } } });
  assert.equal(synth.style, 'Solid');
  assert.equal(synth.color, '#123456');
  assert.equal(synth.width, 2);
  // Nothing declared -> default black 1pt.
  assert.deepEqual(enforcedBottomBorder({}), { style: 'Solid', color: '#000000', width: 1 });
});

test('editable DOCX returns the canonical PDF page count instead of an estimated page range', async () => {
  const m = parseRdl(flatTablixRdl);
  const renderRequest = request(50);
  const [pdf, docx] = await Promise.all([
    renderPdf(m, renderRequest, config),
    renderEditableDocx(m, renderRequest, config),
  ]);
  assert.equal(docx.pageCount, pdf.pageCount);
  assert.equal(docx.layoutMode, 'windows-paged-editable');
});

test('the first visible body Top is preserved across PDF, native Word, visual Word, and XLSX', async (context) => {
  const atPageTop = parseRdl(flatTablixRdl);
  const belowPriorContent = parseRdl(flatTablixRdl.replace('<Top>0in</Top>', '<Top>2in</Top>'));
  const explicitPageStart = parseRdl(flatTablixRdl.replace(
    '<Top>0in</Top>',
    '<PageBreak><BreakLocation>Start</BreakLocation></PageBreak><Top>2in</Top>',
  ));
  for (const model of [atPageTop, belowPriorContent, explicitPageStart]) {
    const tablix = model.body.items.find((item) => item.type === 'Tablix');
    tablix.style.borders.left = { style: 'Solid', color: '#000000', width: 1 };
  }
  const renderRequest = request(1);
  const canonicalTop = async (model) => {
    const canonical = await renderPdf(model, renderRequest, config, { captureLayoutTrace: true });
    const page = canonical.layoutTrace.pages.find((candidate) => candidate.items.some((item) => (
      item.kind === 'tablixCell' && item.text === 'COLHDR'
    )));
    const cell = page?.items.find((item) => item.kind === 'tablixCell' && item.text === 'COLHDR');
    assert.ok(page && cell);
    const docx = await renderEditableDocx(model, renderRequest, config);
    assert.equal(docx.pageCount, canonical.pageCount);
    return { canonical, offset: cell.y - page.bodyTop, xml: await documentXml(docx.buffer) };
  };

  const top = await canonicalTop(atPageTop);
  const offset = await canonicalTop(belowPriorContent);
  const explicitStart = await canonicalTop(explicitPageStart);
  assert.equal(top.offset, 0);
  assert.equal(offset.offset, 144, 'the canonical PDF retains the declared two-inch body offset');
  assert.equal(explicitStart.offset, 0, 'an explicit Start break establishes a page-local origin');

  const firstWordRowHeight = (xml) => Number(xml.match(/<w:tbl>[\s\S]*?<w:trHeight[^>]*w:val="(\d+)"/)?.[1]);
  assert.notEqual(firstWordRowHeight(top.xml), 2880);
  assert.equal(firstWordRowHeight(offset.xml), 2880, 'native Word receives an exact two-inch spacer row');
  assert.notEqual(firstWordRowHeight(explicitStart.xml), 2880);

  const excelPlacement = async (model) => {
    const rendered = await renderExcel(model, { ...renderRequest, excelLayoutMode: 'REPORT' }, config, null);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(rendered.buffer);
    const worksheet = workbook.worksheets[0];
    let target = null;
    worksheet.eachRow((row) => row.eachCell((cell) => {
      if (cell.value === 'COLHDR') target = cell;
    }));
    assert.ok(target);
    return { row: target.row, leadingHeight: worksheet.getRow(1).height };
  };
  const topExcel = await excelPlacement(atPageTop);
  const offsetExcel = await excelPlacement(belowPriorContent);
  const explicitStartExcel = await excelPlacement(explicitPageStart);
  assert.equal(topExcel.row, 1);
  assert.deepEqual(offsetExcel, { row: 2, leadingHeight: 144 });
  assert.equal(explicitStartExcel.row, 1);

  const visualInkTop = async (model, label) => {
    const tempDir = await fs.mkdtemp(path.join(tmpRoot, `initial-body-top-${label}-`));
    context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
    const visual = await renderVisualDocx(model, renderRequest, config, tempDir);
    const zip = await JSZip.loadAsync(visual.buffer);
    const mediaName = Object.keys(zip.files).find((name) => /^word\/media\/.+\.png$/.test(name));
    assert.ok(mediaName);
    const png = PNG.sync.read(await zip.file(mediaName).async('nodebuffer'));
    for (let y = 0; y < png.height; y += 1) {
      for (let x = 0; x < png.width; x += 1) {
        const pixel = (y * png.width + x) * 4;
        if (png.data[pixel] < 220 || png.data[pixel + 1] < 220 || png.data[pixel + 2] < 220) return y;
      }
    }
    return null;
  };
  const topInk = await visualInkTop(atPageTop, 'origin');
  const offsetInk = await visualInkTop(belowPriorContent, 'offset');
  const explicitStartInk = await visualInkTop(explicitPageStart, 'page-start');
  assert.ok(Math.abs((offsetInk - topInk) - 600) <= 3,
    'the 300-DPI visual Word page inherits the two-inch PDF gap within raster antialiasing tolerance');
  assert.equal(explicitStartInk, topInk);

  const overflow = parseRdl(flatTablixRdl.replace('<Top>0in</Top>', '<Top>2in</Top>'));
  overflow.page.height = 100;
  const overflowPdf = await renderPdf(overflow, renderRequest, config, { captureLayoutTrace: true });
  const continuationPage = overflowPdf.layoutTrace.pages.find((page) => page.items.some((item) => (
    item.kind === 'tablixCell' && item.text === 'COLHDR'
  )));
  const continuationHeader = continuationPage?.items.find((item) => item.kind === 'tablixCell' && item.text === 'COLHDR');
  assert.ok(continuationPage && continuationHeader);
  assert.equal(continuationHeader.y, continuationPage.bodyTop,
    'an initial offset that overflows resumes at the next printable body origin');
});

test('the last row of a bordered data tablix closes when its bottom edge is None', async () => {
  const m = parseRdl(flatTablixRdl);
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  const NONE = { style: 'None', color: '#000000', width: 1 };
  const strip = (style) => {
    if (!style) return;
    delete style.border;
    style.borders = { top: { ...NONE }, right: { ...NONE }, bottom: { ...NONE }, left: { ...NONE } };
  };
  strip(tablix.style);
  for (const row of tablix.rows) for (const cell of row.cells) for (const item of cell.items) strip(item.style);
  tablix.style.borders.left = { style: 'Solid', color: '#123456', width: 2 };
  const xml = await documentXml((await renderEditableDocx(m, {
    ...request(3),
  }, config)).buffer);
  assert.match(xml, /<w:tcBorders>[\s\S]*?<w:bottom w:val="single" w:color="123456" w:sz="16"/);
});

test('a static borderless layout tablix honors Border=None instead of receiving a synthetic line', async () => {
  const m = parseRdl(flatTablixRdl.replaceAll(
    '<Style><Border><Style>Solid</Style></Border></Style>',
    '<Style/>',
  ));
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  tablix.rows = [tablix.rows[0]];
  tablix.rowMembers = [tablix.rowMembers[0]];
  tablix.rowMemberPaths = [[tablix.rowMembers[0]]];
  const NONE = { style: 'None', color: '#000000', width: 1 };
  const strip = (style) => {
    if (!style) return;
    delete style.border;
    style.borders = { top: { ...NONE }, right: { ...NONE }, bottom: { ...NONE }, left: { ...NONE } };
  };
  strip(tablix.style);
  for (const cell of tablix.rows[0].cells) for (const item of cell.items) strip(item.style);

  const renderRequest = request(0);
  const { rows } = tablixRows(tablix, renderRequest, { PageNumber: 1, TotalPages: 1 }, m);
  assert.equal(shouldEnforceTablixBottom(rows, tablix), false);
  const xml = await documentXml((await renderEditableDocx(m, renderRequest, config)).buffer);
  const table = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)?.[0] || '';
  assert.match(table, /COLHDR/);
  assert.doesNotMatch(table, /<w:bottom w:val="single"/);
});

test('a dynamic borderless narrative tablix honors Border=None instead of receiving a synthetic line', async () => {
  const m = parseRdl(flatTablixRdl.replaceAll(
    '<Style><Border><Style>Solid</Style></Border></Style>',
    '<Style/>',
  ));
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  const NONE = { style: 'None', color: '#000000', width: 1 };
  const strip = (style) => {
    if (!style) return;
    delete style.border;
    style.borders = { top: { ...NONE }, right: { ...NONE }, bottom: { ...NONE }, left: { ...NONE } };
  };
  strip(tablix.style);
  for (const row of tablix.rows) for (const cell of row.cells) for (const item of cell.items) strip(item.style);

  const renderRequest = request(1);
  const { rows } = tablixRows(tablix, renderRequest, { PageNumber: 1, TotalPages: 1 }, m);
  assert.equal(rows.some((row) => row.isStatic === false), true, 'the narrative row remains data-bound');
  assert.equal(shouldEnforceTablixBottom(rows, tablix), false);
  const xml = await documentXml((await renderEditableDocx(m, renderRequest, config)).buffer);
  assert.match(xml.replace(/<[^>]+>/g, ''), /Row 0/);
  assert.doesNotMatch(xml, /<w:bottom w:val="single"/);
});

test('a trailing static matrix axis row does not receive a synthetic bottom rule', () => {
  const bordered = {
    items: [{
      type: 'Textbox',
      style: {
        borders: {
          top: { style: 'Solid', color: '#000000', width: 1 },
          right: { style: 'Solid', color: '#000000', width: 1 },
          bottom: { style: 'Solid', color: '#000000', width: 1 },
          left: { style: 'Solid', color: '#000000', width: 1 },
        },
      },
    }],
  };
  const rows = [
    { isStatic: false, cells: [bordered] },
    { isStatic: true, role: 'static', cells: [{ items: [], values: [] }] },
  ];

  assert.equal(shouldEnforceTablixBottom(rows, { style: {} }), false);
});

test('a vertical-merge owner reaching the last row carries the enforced bottom border', async () => {
  const m = parseRdl(flatTablixRdl);
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  const detailRow = tablix.rows[1];
  const dynamicMember = tablix.rowMembers[1];
  const rowHeaderCell = structuredClone(detailRow.cells[0]);
  dynamicMember.header = { size: 72, cell: rowHeaderCell };
  tablix.rows = [detailRow];
  tablix.rowMembers = [dynamicMember];
  tablix.rowMemberPaths = [[dynamicMember]];
  tablix.rowHeaderColumns = [72];
  tablix.columns = [72, ...tablix.bodyColumns];
  tablix.width += 72;
  const NONE = { style: 'None', color: '#000000', width: 1 };
  const strip = (style) => {
    if (!style) return;
    delete style.border;
    style.borders = { top: { ...NONE }, right: { ...NONE }, bottom: { ...NONE }, left: { ...NONE } };
  };
  strip(tablix.style);
  for (const cell of tablix.rows[0].cells) for (const item of cell.items) strip(item.style);
  for (const item of rowHeaderCell.items) strip(item.style);

  const xml = await documentXml((await renderEditableDocx(m, {
    parameters: {},
    datasets: { D: [{ V: 'SAME_GROUP' }, { V: 'SAME_GROUP' }] },
  }, config)).buffer);
  assert.match(xml.replace(/<[^>]+>/g, ''), /SAME_GROUP/);
  assert.match(xml, /<w:bottom w:val="single"/);
});

test('oversized tablix output uses one bordered native page table per canonical PDF page', async () => {
  const m = parseRdl(flatTablixRdl);
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  const detailRow = tablix.rows[1];
  detailRow.height = 100;
  const dynamicMember = tablix.rowMembers[1];
  const rowHeaderCell = structuredClone(detailRow.cells[0]);
  const rowHeaderTextbox = rowHeaderCell.items.find((item) => item.type === 'Textbox');
  rowHeaderTextbox.value = 'GROUP_HEADER_ONLY';
  rowHeaderTextbox.paragraphs = [[{ value: 'GROUP_HEADER_ONLY', markupType: 'None', style: rowHeaderTextbox.style }]];
  dynamicMember.header = { size: 72, cell: rowHeaderCell };
  tablix.rows = [detailRow];
  tablix.rowMembers = [dynamicMember];
  tablix.rowMemberPaths = [[dynamicMember]];
  tablix.rowHeaderColumns = [72];
  tablix.columns = [72, ...tablix.bodyColumns];
  tablix.width += 72;

  const renderRequest = {
    parameters: {},
    pagination: { continuationMarkers: true },
    datasets: { D: Array.from({ length: 8 }, () => ({ V: 'OVERSIZED_GROUP' })) },
  };
  const [ordinary, canonical] = await Promise.all([
    renderPdf(m, renderRequest, config),
    renderPdf(m, renderRequest, config, { captureLayoutTrace: true }),
  ]);
  assert.equal(canonical.pageCount, ordinary.pageCount, 'trace capture must not change continuation pagination');
  // PDF metadata timestamps prevent byte-for-byte equality across independent renders. The identical byte
  // length plus the repository-wide trace raster gate proves that recording itself emits no PDF operators.
  assert.equal(canonical.buffer.length, ordinary.buffer.length,
    'recording a continuation marker must not add PDF drawing content');
  const tracedMarkers = canonical.layoutTrace.pages
    .flatMap((page) => page.items)
    .filter((item) => item.traceRole === 'continuationMarker');
  assert.equal(tracedMarkers.length, canonical.pageCount - 1);
  assert.equal(tracedMarkers.every((item) => item.text === 'Continued from previous page'), true);
  for (const page of canonical.layoutTrace.pages.slice(1)) {
    const marker = page.items.find((item) => item.traceRole === 'continuationMarker');
    const firstTableTop = Math.min(...page.items.filter((item) => item.kind === 'tablixCell').map((item) => item.y));
    assert.ok(Math.abs(marker.y + marker.height - firstTableTop) <= 0.25,
      'the marker must end at the existing repeated-header boundary without moving table cells');
  }
  const xml = await documentXml((await renderEditableDocx(m, renderRequest, config)).buffer);
  const tables = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  assert.equal(tables.length, canonical.pageCount, 'each canonical page must become one explicit native page table');
  assert.equal((xml.match(/<w:sectPr(?:\s|>)/g) || []).length, tables.length);
  assert.equal((xml.match(/Continued from previous page/g) || []).length, tracedMarkers.length,
    'Word must fill the already-reserved canonical marker region with native text');
  for (const table of tables) {
    const physicalRows = [...table.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map((match) => match[0]);
    assert.equal(physicalRows.every((row) => /<w:cantSplit\/>/.test(row)), true);
    assert.match(table, /<w:bottom w:val="single"/);
  }

});

test('ordinary DOCX table fragmentation does not claim that adjacent rows are continuations', async () => {
  const m = parseRdl(flatTablixRdl);
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  tablix.rows[1].height = 100;

  const renderRequest = {
    parameters: {},
    pagination: { continuationMarkers: true },
    datasets: { D: Array.from({ length: 8 }, (_, index) => ({ V: `ROW_${index + 1}` })) },
  };
  const canonical = await renderPdf(m, renderRequest, config, { captureLayoutTrace: true });
  assert.equal(canonical.layoutTrace.pages.flatMap((page) => page.items)
    .some((item) => item.traceRole === 'continuationMarker'), false);
  const xml = await documentXml((await renderEditableDocx(m, renderRequest, config)).buffer);
  assert.equal((xml.match(/<w:tbl>/g) || []).length > 1, true, 'the native table should have multiple page fragments');
  assert.doesNotMatch(xml, /Continued (?:on next|from previous) page/);
});
