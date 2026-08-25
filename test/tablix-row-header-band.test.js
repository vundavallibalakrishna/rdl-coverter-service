import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows } from '../src/rdl/validation.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

// A tablix can lay its whole visible grid out through ROW HEADERS, with the body column hidden — SSRS's
// usual shape for a grouped "plan" table. Two rules that shape depends on:
//
//  * Every TablixHeader is as wide as its declared Size. A group's own header branch is shallower than the
//    detail branch, so its single wide cell spans the leaf hierarchy columns that size covers — a quarter
//    band across the whole table, not a narrow cell in one column with a blank tail beside it.
//  * A member SortExpression reorders the group instances, and running aggregates (RowNumber,
//    RunningValue) accumulate in the order the region emits rows — not in dataset arrival order.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const NARROW_PT = 30;
const WIDE_PT = 100;
const BAND_PT = WIDE_PT * 2; // the band's declared size covers both wide leaf columns

const textbox = (name, value, background) => `<Textbox Name="${name}"><CanGrow>true</CanGrow>`
  + `<Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value>`
  + '<Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>'
  + '<Style><Border><Style>Solid</Style></Border>'
  + '<TopBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>'
  + '<BottomBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></BottomBorder>'
  + '<LeftBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>'
  + '<RightBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></RightBorder>'
  + `${background ? `<BackgroundColor>${background}</BackgroundColor>` : ''}</Style></Textbox>`;

const header = (size, contents) => `<TablixHeader><Size>${size}pt</Size><CellContents>${contents}</CellContents></TablixHeader>`;

// The detail branch: three sized headers, innermost last.
const detailHeaderChain = (prefix) => `${header(NARROW_PT, textbox(`${prefix}Num`, '=RunningValue(Fields!ItemId.Value, CountDistinct, "Quarter")'))}
  <TablixMembers><TablixMember>
    ${header(WIDE_PT, textbox(`${prefix}Name`, '=Fields!Name.Value'))}
    <TablixMembers><TablixMember>
      ${header(WIDE_PT, textbox(`${prefix}When`, '=Fields!Quarter.Value'))}
      <TablixMembers><TablixMember/></TablixMembers>
    </TablixMember></TablixMembers>
  </TablixMember></TablixMembers>`;

const columnHeaderChain = `${header(NARROW_PT, textbox('HeadNum', '#', 'DarkBlue'))}
  <TablixMembers><TablixMember>
    ${header(WIDE_PT, textbox('HeadName', 'Name', 'DarkBlue'))}
    <TablixMembers><TablixMember>
      ${header(WIDE_PT, textbox('HeadWhen', 'When', 'DarkBlue'))}
      <TablixMembers><TablixMember/></TablixMembers>
    </TablixMember></TablixMembers>
  </TablixMember></TablixMembers>`;

// The band branch: two headers, the second sized across both wide leaf columns.
const bandChain = `${header(NARROW_PT, textbox('BandPad', '', '#3cb7e0'))}
  <TablixMembers><TablixMember>
    ${header(BAND_PT, textbox('BandLabel', '=Fields!Quarter.Value', '#3cb7e0'))}
    <TablixMembers><TablixMember/></TablixMembers>
  </TablixMember></TablixMembers>`;

// A counterexample branch: two headers, the second sized to ONE leaf column, so it must not widen.
const narrowBandChain = `${header(NARROW_PT, textbox('BandPad', '', '#3cb7e0'))}
  <TablixMembers><TablixMember>
    ${header(WIDE_PT, textbox('BandLabel', '=Fields!Quarter.Value', '#3cb7e0'))}
    <TablixMembers><TablixMember/></TablixMembers>
  </TablixMember></TablixMembers>`;

const bodyRow = (name) => `<TablixRow><Height>16pt</Height><TablixCells><TablixCell><CellContents>
  ${textbox(name, '')}
</CellContents></TablixCell></TablixCells></TablixRow>`;

const rdl = (band) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Quarter"><DataField>Quarter</DataField></Field>
    <Field Name="ItemId"><DataField>ItemId</DataField></Field>
    <Field Name="Name"><DataField>Name</DataField></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Plan"><TablixBody>
      <TablixColumns><TablixColumn><Width>2pt</Width></TablixColumn></TablixColumns>
      <TablixRows>${bodyRow('BodyHead')}${bodyRow('BodyBand')}${bodyRow('BodyDetail')}</TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember><Visibility><Hidden>true</Hidden></Visibility></TablixMember></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers>
      <TablixMember>${columnHeaderChain}<RepeatOnNewPage>true</RepeatOnNewPage></TablixMember>
      <TablixMember>
        <Group Name="Quarter"><GroupExpressions><GroupExpression>=Fields!Quarter.Value</GroupExpression></GroupExpressions></Group>
        <TablixMembers>
          <TablixMember>${band}</TablixMember>
          <TablixMember><TablixMembers><TablixMember>
            <Group Name="Item"><GroupExpressions><GroupExpression>=Fields!ItemId.Value</GroupExpression></GroupExpressions></Group>
            <SortExpressions><SortExpression><Value>=Fields!ItemId.Value</Value></SortExpression></SortExpressions>
            ${detailHeaderChain('Detail')}
          </TablixMember></TablixMembers></TablixMember>
        </TablixMembers>
      </TablixMember>
    </TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0pt</Top><Left>0pt</Left><Height>48pt</Height><Width>232pt</Width>
    <Style><Border><Style>Solid</Style></Border></Style>
    </Tablix>
  </ReportItems><Height>200pt</Height><Style/></Body><Width>240pt</Width>
  <Page><PageWidth>300pt</PageWidth><PageHeight>400pt</PageHeight><TopMargin>10pt</TopMargin>
    <BottomMargin>10pt</BottomMargin><LeftMargin>10pt</LeftMargin><RightMargin>10pt</RightMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

// Deliberately out of ItemId order, so arrival order and sorted order differ.
const rows = [
  { Quarter: 'Q1', ItemId: 30, Name: 'Gamma' },
  { Quarter: 'Q1', ItemId: 10, Name: 'Alpha' },
  { Quarter: 'Q1', ItemId: 20, Name: 'Beta' },
];
const request = { outputFileName: 'row-header-band', parameters: {}, datasets: { D: rows }, excel: { layoutMode: 'REPORT' } };

function planOf(rdlText) {
  const model = parseRdl(rdlText);
  const tablix = model.body.items.find((item) => item.name === 'Plan');
  return { model, tablix, rows: materializeTablixRows(tablix, rows, {}, {}, { D: rows }) };
}

const cellText = (cell) => (cell.values || []).map((value) => String(value ?? '')).filter(Boolean).join('');

test('the row-header grid comes from the deepest branch and the band spans its declared size', () => {
  const { tablix, rows: materialized } = planOf(rdl(bandChain));
  assert.deepEqual(tablix.rowHeaderColumns, [NARROW_PT, WIDE_PT, WIDE_PT]);
  const band = materialized.find((row) => row.cells.some((cell) => cellText(cell) === 'Q1'));
  assert.ok(band, 'the group band row must materialize');
  const labelCell = band.cells.find((cell) => cellText(cell) === 'Q1');
  assert.equal(labelCell.colSpan, 2, 'the band spans the two leaf columns its declared size covers');
  // The pad cell plus the band cover the whole row-header grid, so no blank tail cell is synthesized.
  assert.equal(band.cells.reduce((sum, cell) => sum + (cell.colSpan || 1), 0), 3);
  assert.equal(band.cells.length, 2);
});

test('a band whose declared size covers one column keeps its single column', () => {
  const { rows: materialized } = planOf(rdl(narrowBandChain));
  const band = materialized.find((row) => row.cells.some((cell) => cellText(cell) === 'Q1'));
  const labelCell = band.cells.find((cell) => cellText(cell) === 'Q1');
  assert.equal(labelCell.colSpan, 1, 'a one-column band must not widen');
  // The uncovered leaf column is still materialized as a blank grid cell.
  assert.equal(band.cells.reduce((sum, cell) => sum + (cell.colSpan || 1), 0), 3);
});

test('a running aggregate numbers the SORTED group instances, not dataset arrival order', () => {
  const { rows: materialized } = planOf(rdl(bandChain));
  const detail = materialized.filter((row) => /^(Alpha|Beta|Gamma)$/.test(row.cells.map(cellText).join('')) === false
    && row.cells.some((cell) => /^(Alpha|Beta|Gamma)$/.test(cellText(cell))));
  assert.equal(detail.length, 3, 'three detail rows');
  assert.deepEqual(
    detail.map((row) => cellText(row.cells[0])),
    ['1', '2', '3'],
    'RunningValue must count in the emitted (sorted) order',
  );
  assert.deepEqual(
    detail.map((row) => cellText(row.cells[1])),
    ['Alpha', 'Beta', 'Gamma'],
    'the ItemId sort must order the rows',
  );
});

test('PDF, editable DOCX, and XLSX all render the band span and the sorted numbering', async (context) => {
  const model = parseRdl(rdl(bandChain));
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-row-band-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const owned = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000', RDL_TEMP_ROOT: tempRoot });

  const pdf = await renderPdf(model, request, owned, { captureLayoutTrace: true });
  const traced = pdf.layoutTrace.pages.flatMap((page) => (page.tablixFragments || []).flatMap((fragment) => fragment.cells || []));
  const bandCell = traced.find((cell) => (cell.text || '').trim() === 'Q1');
  assert.ok(bandCell, 'the band renders');
  assert.equal(bandCell.colSpan, 2);
  // The grid scales the columns to the tablix's declared width, so compare against a single column rather
  // than the raw declared size: the band must clearly cover both wide columns, not one.
  assert.ok(
    bandCell.width > WIDE_PT * 1.5,
    `the band is ${bandCell.width}pt, which is not wider than the ${WIDE_PT}pt column it used to sit in`,
  );
  const numbers = traced.filter((cell) => /^[123]$/.test((cell.text || '').trim()))
    .sort((left, right) => left.y - right.y)
    .map((cell) => cell.text.trim());
  assert.deepEqual(numbers, ['1', '2', '3']);

  const docx = await renderEditableDocx(model, request, owned);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const wordRows = [...documentXml.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((match) => match[0]);
  const text = (fragment) => [...fragment.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join('').trim();
  const bandRow = wordRows.find((row) => text(row) === 'Q1');
  assert.ok(bandRow, 'the band reaches Word');
  const bandWordCell = [...bandRow.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((match) => match[0]).find((cellXml) => text(cellXml) === 'Q1');
  assert.ok(/<w:gridSpan w:val="[2-9]\d*"\/>/.test(bandWordCell), 'the band must be a merged Word cell, not one column');
  const wordCellTexts = wordRows.flatMap((row) => [...row.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((match) => text(match[0])));
  assert.deepEqual(wordCellTexts.filter((value) => /^[123]$/.test(value)), ['1', '2', '3']);

  const xlsx = await renderExcel(model, request, owned, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const sheet = workbook.worksheets[0];
  let bandMaster = null;
  const excelNumbers = [];
  sheet.eachRow((row) => row.eachCell((cell) => {
    const value = typeof cell.value === 'object' && cell.value?.richText
      ? cell.value.richText.map((run) => run.text).join('')
      : String(cell.value ?? '');
    if (value.trim() === 'Q1' && !bandMaster) bandMaster = cell.master ? cell.master.address : cell.address;
    if (/^[123]$/.test(value.trim())) excelNumbers.push(value.trim());
  }));
  assert.ok(bandMaster, 'the band reaches Excel');
  const bandRange = (sheet.model.merges || []).find((range) => range.startsWith(`${bandMaster}:`));
  assert.ok(bandRange, `the band must be merged across its columns (merges: ${(sheet.model.merges || []).join(' ')})`);
  assert.deepEqual([...new Set(excelNumbers)], ['1', '2', '3']);
});
