import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

// Where a child data region held by a MERGED cell lives in the worksheet.
//
// A cell that spans several tablix rows holds its child region in the whole block of rows it covers: the
// fixed-layout renderer draws the child from the block's top and lets it flow past the first row's bottom,
// so the child's rows and the spanned rows interleave. The worksheet has to reproduce that, because it is
// the same grid — only with Excel's own measured heights.
//
// This renderer gave the child every boundary of the row its cell STARTS in. That made the starting row as
// tall as the entire child grid and pushed every later row of the block below it, so the columns beside the
// child ran out of step with it: the second and third rows of the block began under the whole child table
// instead of beside its second and third rows, and their rules no longer lined up with anything. It is the
// worksheet form of the defect the fixed-layout renderer fixed by growing a merge's LAST spanned row.
//
// The oracle is the canonical PDF layout: for every pair of content blocks, the vertical relationship the
// PDF gives them (one above the other, one inside the other's span, or the two sharing a span) must be the
// relationship the worksheet gives them.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });

const textbox = (name, value) => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>`
  + '</TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border>'
  + '<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight></Style></Textbox>';

// One group whose merged header cell holds a CHILD tablix, beside DETAILS ordinary detail rows. Both stacks
// start at the group's top and have their own row heights, so their boundaries interleave.
const report = ({ childRows = 3, childRowHeight = 0.4 } = {}) => Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="DS"><Fields>
    <Field Name="G"><DataField>G</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="D"><DataField>D</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="E"><DataField>E</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="T"><TablixBody>
      <TablixColumns><TablixColumn><Width>1.6in</Width></TablixColumn><TablixColumn><Width>1.6in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
        <TablixCell><CellContents>${textbox('Detail', '=Fields!D.Value')}</CellContents></TablixCell>
        <TablixCell><CellContents>${textbox('Side', '=Fields!E.Value')}</CellContents></TablixCell>
      </TablixCells></TablixRow></TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers>
        <TablixMember><Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group>
          <TablixHeader><Size>1.6in</Size><CellContents>
            <Tablix Name="Child"><TablixBody>
              <TablixColumns><TablixColumn><Width>1.6in</Width></TablixColumn></TablixColumns>
              <TablixRows>${Array.from({ length: childRows }, (unused, index) => (
    `<TablixRow><Height>${childRowHeight}in</Height><TablixCells><TablixCell><CellContents>`
                + `${textbox(`ChildCell${index + 1}`, `CHILD_${index + 1}`)}</CellContents></TablixCell></TablixCells></TablixRow>`
  )).join('')}</TablixRows>
            </TablixBody>
            <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
            <TablixRowHierarchy><TablixMembers>${'<TablixMember/>'.repeat(childRows)}</TablixMembers></TablixRowHierarchy>
            <Top>0in</Top><Left>0in</Left><Height>${childRowHeight * childRows}in</Height><Width>1.6in</Width><Style/></Tablix>
          </CellContents></TablixHeader>
          <TablixMembers>
            <TablixMember><Group Name="D"><GroupExpressions><GroupExpression>=Fields!D.Value</GroupExpression></GroupExpressions></Group></TablixMember>
          </TablixMembers>
        </TablixMember>
      </TablixMembers></TablixRowHierarchy>
      <DataSetName>DS</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>4.8in</Width><Style/></Tablix>
  </ReportItems><Height>9in</Height><Style/></Body><Width>8in</Width>
  <Page><PageHeight>14in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.3in</LeftMargin><RightMargin>0.3in</RightMargin><TopMargin>0.3in</TopMargin><BottomMargin>0.3in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`, 'utf8');

// Detail rows of visibly different heights, so the two stacks cannot line up by accident.
const DETAILS = ['DETAIL_1', 'DETAIL_2\nsecond line\nthird line', 'DETAIL_3'];
const rows = DETAILS.map((value, index) => ({ G: 'G1', D: value, E: `SIDE_${index + 1}` }));
const request = (output) => ({ output, outputFileName: 'merged-child-grid', parameters: {}, datasets: { DS: rows } });

const label = (text) => String(text || '').split('\n')[0].trim();

const pdfBlocks = async (definition) => {
  const rendered = await renderPdf(parseRdl(definition), request('PDF'), config, { captureLayoutTrace: true });
  assert.equal(rendered.pageCount, 1, 'the fixture must fit one page, so pagination cannot explain a difference');
  const blocks = new Map();
  for (const item of rendered.layoutTrace.pages[0].items || []) {
    if (item.kind !== 'tablixCell' || !(item.text || '').trim()) continue;
    const key = label(item.text);
    if (!blocks.has(key)) blocks.set(key, { top: item.y, bottom: item.y + item.height });
  }
  return blocks;
};

const excelBlocks = async (definition) => {
  const excel = await renderExcel(parseRdl(definition), request('XLSX'), config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  const sheet = workbook.worksheets[0];
  const tops = [0];
  for (let row = 1; row <= sheet.rowCount; row += 1) tops.push(tops[row - 1] + (sheet.getRow(row).height ?? 15));
  const spans = new Map();
  for (const range of sheet.model.merges) {
    const [from, to] = range.split(':');
    spans.set(from, { startRow: Number(from.replace(/\D+/g, '')), endRow: Number(to.replace(/\D+/g, '')) });
  }
  const blocks = new Map();
  sheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, column) => {
      if (cell.isMerged && cell.master !== cell) return;
      const value = cell.value?.richText ? cell.value.richText.map((run) => run.text).join('') : cell.value;
      if (!value || !String(value).trim()) return;
      const span = spans.get(cell.address) || { startRow: rowNumber, endRow: rowNumber };
      const key = label(String(value));
      if (!blocks.has(key)) blocks.set(key, { top: tops[span.startRow - 1], bottom: tops[span.endRow], column });
    });
  });
  return blocks;
};

// How two blocks sit relative to each other, at a scale-free level: the answer must not depend on the
// renderer's measured heights, only on the structure they share.
const relation = (left, right) => {
  if (right.top >= left.bottom - 0.5) return 'above';
  if (left.top >= right.bottom - 0.5) return 'below';
  if (Math.abs(left.top - right.top) < 0.5 && Math.abs(left.bottom - right.bottom) < 0.5) return 'same';
  return 'overlap';
};

test('a child grid in a merged cell keeps the vertical structure the fixed layout gives it', async () => {
  const definition = report();
  const [pdf, xlsx] = await Promise.all([pdfBlocks(definition), excelBlocks(definition)]);

  const keys = [...pdf.keys()].filter((key) => xlsx.has(key));
  assert.ok(
    ['CHILD_1', 'CHILD_3', 'DETAIL_1', 'DETAIL_3'].every((key) => keys.includes(key)),
    `both outputs must carry the whole fixture, got ${keys.join(', ')}`,
  );

  // The construct is only meaningful if the two stacks really do interleave in the fixed layout: their
  // boundaries fall at different heights, so no child row simply matches a detail row.
  const childKeys = keys.filter((key) => key.startsWith('CHILD_'));
  const detailKeys = keys.filter((key) => key.startsWith('DETAIL_'));
  const crossRelations = childKeys.flatMap((child) => detailKeys.map((detail) => relation(pdf.get(child), pdf.get(detail))));
  assert.ok(crossRelations.includes('overlap'), `the two stacks must interleave, got ${crossRelations.join(',')}`);
  assert.ok(crossRelations.includes('below'), 'and the child stack must reach past the first detail row');

  const mismatches = [];
  for (let i = 0; i < keys.length; i += 1) {
    for (let j = i + 1; j < keys.length; j += 1) {
      const inPdf = relation(pdf.get(keys[i]), pdf.get(keys[j]));
      const inXlsx = relation(xlsx.get(keys[i]), xlsx.get(keys[j]));
      if (inPdf !== inXlsx) mismatches.push(`${keys[i]} vs ${keys[j]}: pdf=${inPdf} xlsx=${inXlsx}`);
    }
  }
  assert.deepEqual(mismatches, [], 'the worksheet must place every block the way the canonical layout does');
});

test('the rows beside a child grid are not pushed below it', async () => {
  const definition = report({ childRows: 4 });
  const xlsx = await excelBlocks(definition);
  // The narrow statement of the defect: with the child grid confined to the row its cell starts in, every
  // detail row after the first began below the complete child table.
  assert.equal(relation(xlsx.get('CHILD_4'), xlsx.get('DETAIL_2')), 'below', 'the last child row starts after the second detail row');
  assert.notEqual(relation(xlsx.get('CHILD_4'), xlsx.get('DETAIL_2')), 'above');
  assert.equal(xlsx.get('CHILD_1').column, 1, 'the child grid stays in the merged header column');
  assert.equal(xlsx.get('DETAIL_1').column, 2, 'and the detail rows stay in theirs');
});

test('a child grid shorter than its block leaves the rest of the merged cell empty', async () => {
  // One short child row inside a three-row block: the child closes at its own height, and the block's
  // remaining rows are not swallowed by it.
  const definition = report({ childRows: 1, childRowHeight: 0.2 });
  const [pdf, xlsx] = await Promise.all([pdfBlocks(definition), excelBlocks(definition)]);
  assert.equal(relation(pdf.get('CHILD_1'), pdf.get('DETAIL_3')), 'above');
  assert.equal(relation(xlsx.get('CHILD_1'), xlsx.get('DETAIL_3')), 'above');
});
