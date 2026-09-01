import assert from 'node:assert/strict';
import test from 'node:test';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

// What a continuation label means, and what it must never mean.
//
// There is exactly one annotation, and it is driven by detected layout state rather than by a page break:
// this physical tablix row's own content was cut by the break and resumes on this page. Everything else
// that can cross a page boundary — a group instance, a repeated header, a child region sliced at a row
// boundary — starts its page on a whole row and is therefore NOT a continuation and carries nothing.
//
// Everything below is built from a generic fixture builder — column count, grouping shape, page geometry,
// fonts and padding are all parameters — because the feature lives in the tablix pagination engine and
// applies to every report. No assertion depends on a report name, a dataset name, a column name, or a
// rows-per-page constant.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '120000' });

const columnName = (index) => String.fromCharCode(65 + index); // A, B, C, ... — the spec's A–E is index 0..4

// One textbox. `style` lets a single column differ in family/size/padding so the split cell can be
// measured under a different font than its neighbours.
const textbox = (name, value, style = {}, extra = '') => {
  const { fontFamily = 'Arial', fontSize = 9, padding = 2 } = style;
  return `<Textbox Name="${name}"><CanGrow>true</CanGrow>${extra}<Paragraphs><Paragraph><TextRuns>`
    + `<TextRun><Value>${value}</Value><Style><FontFamily>${fontFamily}</FontFamily>`
    + `<FontSize>${fontSize}pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>`
    + '<Style><Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border>'
    + `<PaddingLeft>${padding}pt</PaddingLeft><PaddingRight>${padding}pt</PaddingRight>`
    + `<PaddingTop>${padding}pt</PaddingTop><PaddingBottom>${padding}pt</PaddingBottom></Style></Textbox>`;
};

const cell = (contents) => `<TablixCell><CellContents>${contents}</CellContents></TablixCell>`;

// A tablix in one of the three shapes a group can take: none, a merged row-header cell (one level or
// nested), or a group header ROW. Column count, row heights and cell styling are parameters.
function tablixXml({
  name,
  columns,
  grouping,
  rowHeightIn,
  headerHeightIn,
  columnWidthIn,
  groupWidthIn,
  columnStyles = {},
  keepTogether = false,
}) {
  const keep = keepTogether ? '<KeepTogether>true</KeepTogether>' : '';
  const headerCells = Array.from({ length: columns }, (unused, index) => cell(textbox(
    `${name}Head${index}`, `HEAD_${columnName(index)}`,
  ))).join('');
  const detailCells = Array.from({ length: columns }, (unused, index) => cell(textbox(
    `${name}Body${index}`, `=Fields!${columnName(index)}.Value`, columnStyles[index] || {}, keep,
  ))).join('');
  const bodyColumns = Array.from({ length: columns }, () => `<TablixColumn><Width>${columnWidthIn}in</Width></TablixColumn>`).join('');
  const row = (heightIn, cells) => `<TablixRow><Height>${heightIn}in</Height><TablixCells>${cells}</TablixCells></TablixRow>`;

  // A group header ROW is a static leaf member nested inside a group, so the region carries an extra row
  // template that renders once per group instance.
  const groupHeaderRow = grouping === 'headerRow'
    ? row(headerHeightIn, cell(textbox(`${name}GroupRow`, '=Fields!GRP.Value'))
      + Array.from({ length: columns - 1 }, (unused, index) => cell(textbox(`${name}GroupPad${index}`, ''))).join(''))
    : '';

  const detailMember = '<TablixMember><Group Name="Detail"><GroupExpressions>'
    + '<GroupExpression>=Fields!KEY.Value</GroupExpression></GroupExpressions></Group></TablixMember>';
  const groupHeaderCell = (label, value) => `<TablixHeader><Size>${groupWidthIn}in</Size>`
    + `<CellContents>${textbox(label, value)}</CellContents></TablixHeader>`;

  let hierarchy;
  if (grouping === 'merged') {
    hierarchy = `<TablixMember>${groupHeaderCell(`${name}HeadGroup`, 'GROUP')}<TablixMembers><TablixMember/></TablixMembers></TablixMember>`
      + '<TablixMember><Group Name="Grp"><GroupExpressions><GroupExpression>=Fields!GRP.Value</GroupExpression></GroupExpressions></Group>'
      + groupHeaderCell(`${name}GroupCell`, '=Fields!GRP.Value')
      + `<TablixMembers>${detailMember}</TablixMembers></TablixMember>`;
  } else if (grouping === 'nested') {
    hierarchy = `<TablixMember>${groupHeaderCell(`${name}HeadOuter`, 'OUTER')}<TablixMembers>`
      + `<TablixMember>${groupHeaderCell(`${name}HeadInner`, 'INNER')}<TablixMembers><TablixMember/></TablixMembers></TablixMember>`
      + '</TablixMembers></TablixMember>'
      + '<TablixMember><Group Name="Outer"><GroupExpressions><GroupExpression>=Fields!GRP.Value</GroupExpression></GroupExpressions></Group>'
      + groupHeaderCell(`${name}OuterCell`, '=Fields!GRP.Value')
      + '<TablixMembers><TablixMember><Group Name="Inner"><GroupExpressions><GroupExpression>=Fields!DEPT.Value</GroupExpression></GroupExpressions></Group>'
      + groupHeaderCell(`${name}InnerCell`, '=Fields!DEPT.Value')
      + `<TablixMembers>${detailMember}</TablixMembers></TablixMember></TablixMembers></TablixMember>`;
  } else if (grouping === 'headerRow') {
    hierarchy = '<TablixMember/>'
      + '<TablixMember><Group Name="Grp"><GroupExpressions><GroupExpression>=Fields!GRP.Value</GroupExpression></GroupExpressions></Group>'
      + `<TablixMembers><TablixMember/>${detailMember}</TablixMembers></TablixMember>`;
  } else {
    hierarchy = `<TablixMember/>${detailMember}`;
  }

  const rowHeaderColumns = grouping === 'merged' ? 1 : (grouping === 'nested' ? 2 : 0);
  const width = columns * columnWidthIn + rowHeaderColumns * groupWidthIn;
  return `<Tablix Name="${name}"><TablixBody><TablixColumns>${bodyColumns}</TablixColumns>`
    + `<TablixRows>${row(headerHeightIn, headerCells)}${groupHeaderRow}${row(rowHeightIn, detailCells)}</TablixRows></TablixBody>`
    + '<TablixColumnHierarchy><TablixMembers>'
    + Array.from({ length: columns }, () => '<TablixMember/>').join('')
    + '</TablixMembers></TablixColumnHierarchy>'
    + `<TablixRowHierarchy><TablixMembers>${hierarchy}</TablixMembers></TablixRowHierarchy>`
    + `<DataSetName>D</DataSetName><Top>{TOP}in</Top><Left>0in</Left><Width>${width}in</Width><Style/></Tablix>`;
}

function report({
  columns = 5,
  grouping = 'none',
  pageHeightIn = 2.5,
  pageWidthIn = 8.5,
  marginIn = 0.25,
  rowHeightIn = 0.25,
  headerHeightIn = 0.25,
  columnWidthIn = 1.2,
  groupWidthIn = 1,
  columnStyles = {},
  keepTogether = false,
  tablixCount = 1,
} = {}) {
  const fields = ['KEY', 'GRP', 'DEPT', ...Array.from({ length: columns }, (unused, index) => columnName(index))];
  const bodies = Array.from({ length: tablixCount }, (unused, index) => tablixXml({
    name: `T${index + 1}`, columns, grouping, rowHeightIn, headerHeightIn, columnWidthIn, groupWidthIn, columnStyles, keepTogether,
  }).replace('{TOP}', index === 0 ? '0' : String(index * 0.01)));
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>${fields
    .map((field) => `<Field Name="${field}"><DataField>${field}</DataField><TypeName>System.String</TypeName></Field>`)
    .join('')}</Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>${bodies.join('')}</ReportItems>
    <Height>${pageHeightIn}in</Height><Style/></Body><Width>${pageWidthIn}in</Width>
  <Page><PageHeight>${pageHeightIn}in</PageHeight><PageWidth>${pageWidthIn}in</PageWidth>
    <TopMargin>${marginIn}in</TopMargin><BottomMargin>${marginIn}in</BottomMargin>
    <LeftMargin>${marginIn}in</LeftMargin><RightMargin>${marginIn}in</RightMargin>
  </Page></ReportSection></ReportSections></Report>`, 'utf8');
}

// Explicit line breaks make a cell's measured height deterministic under any installed font, so "this row
// is taller than a page" is a property of the fixture rather than of the test machine.
const lines = (prefix, count) => Array.from({ length: count }, (unused, index) => `${prefix}_${String(index + 1).padStart(3, '0')}`).join('\n');

const dataRow = ({ key, group = 'G1', dept = 'D1', columns = 5, tall = 0, tallColumn = 4, prefix = 'X' }) => {
  const row = { KEY: key, GRP: group, DEPT: dept };
  for (let index = 0; index < columns; index += 1) {
    row[columnName(index)] = index === tallColumn && tall > 0 ? lines(prefix, tall) : `${columnName(index)}_${key}`;
  }
  return row;
};

const render = async (rdl, rows, extra = {}) => renderPdf(
  parseRdl(rdl),
  {
    outputFileName: 'continuation-labels',
    parameters: {},
    datasets: { D: rows },
    pagination: { continuationMarkers: true },
    ...extra,
  },
  config,
  { captureLayoutTrace: true },
);

const items = (rendered) => rendered.layoutTrace.pages.flatMap((page) => (page.items || []).map((item) => ({ ...item, page: page.number })));
const rowLabels = (rendered) => items(rendered).filter((item) => item.traceRole === 'continuationMarker');
const rowLabelPages = (rendered) => [...new Set(rowLabels(rendered).map((item) => item.page))].sort((a, b) => a - b);
// No fixture writes "continu" into its data, so any cell containing it would be an annotation the
// renderer added to something that is not a split row.
const annotatedCells = (rendered) => items(rendered)
  .filter((item) => item.kind === 'tablixCell' && /continu/i.test(item.text || ''));
const pagesWith = (rendered, needle) => [...new Set(items(rendered)
  .filter((item) => (item.text || '').includes(needle))
  .map((item) => item.page))].sort((a, b) => a - b);

// ------------------------------------------------------------------ 1. a row that fits is never labelled

test('1: rows that fit produce no label of either kind', async () => {
  const rdl = report();
  const rows = Array.from({ length: 5 }, (unused, index) => dataRow({ key: `R${index + 1}` }));
  const rendered = await render(rdl, rows);
  assert.equal(rendered.pageCount, 1);
  assert.deepEqual(rowLabelPages(rendered), []);
  assert.deepEqual(annotatedCells(rendered), []);
});

// ------------------------------------------------------------------ 2. one split → label on page 2 only

test('2: a row split once carries the label on the page it resumes on, and nowhere else', async () => {
  const rdl = report();
  const rows = [dataRow({ key: 'R1', tall: 18, prefix: 'SPLIT' })];
  const rendered = await render(rdl, rows);
  assert.equal(rendered.pageCount, 2, `the row must cross exactly one boundary (got ${rendered.pageCount} pages)`);
  assert.deepEqual(rowLabelPages(rendered), [2]);
  assert.equal(rowLabels(rendered)[0].text, 'Continued from previous page');
  // Nothing is lost or duplicated by the split.
  for (let index = 1; index <= 18; index += 1) {
    assert.equal(pagesWith(rendered, `SPLIT_${String(index).padStart(3, '0')}`).length, 1);
  }
});

// ------------------------------------------------------------------ 3. only the last column overflows

test('3: the row is split where its overflowing cell overflows, not pushed whole to the next page', async () => {
  const rdl = report({ columns: 5 });
  const rows = [
    dataRow({ key: 'R1' }),
    dataRow({ key: 'R2', tall: 18, tallColumn: 4, prefix: 'ECOL' }),
  ];
  const rendered = await render(rdl, rows);
  // A–D are short and render once, on the page the row starts on; E continues onto the next page.
  for (const index of [0, 1, 2, 3]) {
    assert.deepEqual(pagesWith(rendered, `${columnName(index)}_R2`), [1], `column ${columnName(index)} belongs to page 1`);
  }
  assert.deepEqual(pagesWith(rendered, 'ECOL_001'), [1], 'the fitting part of E stays on page 1');
  assert.deepEqual(pagesWith(rendered, 'ECOL_018'), [2], 'the rest of E continues on page 2');
  assert.deepEqual(rowLabelPages(rendered), [2]);
});

// ------------------------------------------------------------------ 3b. use a remainder before a fresh page

test('3b: a growable row that fits a fresh page uses the current remainder and resumes below the repeated header', async () => {
  const rdl = report({ keepTogether: true });
  const rows = [
    ...Array.from({ length: 4 }, (unused, index) => dataRow({ key: `LEAD_${index + 1}` })),
    dataRow({ key: 'R2', tall: 8, tallColumn: 4, prefix: 'REMAINDER' }),
    dataRow({ key: 'R3' }),
  ];
  const rendered = await render(rdl, rows);
  const firstLinePage = pagesWith(rendered, 'REMAINDER_001')[0];
  const lastLinePage = pagesWith(rendered, 'REMAINDER_008')[0];

  assert.equal(firstLinePage, 1, 'the row starts in the otherwise unused remainder');
  assert.equal(lastLinePage, 2, 'the row continues on the following page even though it fits a fresh page');
  assert.deepEqual(rowLabelPages(rendered), [2], 'only the split row receives a continuation label');

  const secondPageItems = items(rendered).filter((item) => item.page === 2);
  const repeatedHeader = secondPageItems.find((item) => item.kind === 'tablixCell' && item.text.includes('HEAD_A'));
  const label = secondPageItems.find((item) => item.traceRole === 'continuationMarker');
  const tail = secondPageItems.find((item) => item.kind === 'tablixCell' && item.text.includes('REMAINDER_008'));
  assert.ok(repeatedHeader && label && tail, 'the continued page contains its repeated header, label and tail');
  assert.ok(repeatedHeader.y < label.y, 'the repeated table header precedes the continuation label');
  assert.ok(label.y + label.height <= tail.y + 0.5, 'the continuation label precedes the remaining row content');
  assert.equal(pagesWith(rendered, 'A_R3')[0], 2, 'the following whole row uses the same continuation page');
});

// ------------------------------------------------------------------ 4. a row taller than a whole page

test('4: a row spanning three pages is labelled on every continuation page, under repeated headers', async () => {
  const rdl = report();
  const rows = [dataRow({ key: 'R1', tall: 40, prefix: 'GIANT' })];
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 3, `the row must span 3+ pages (got ${rendered.pageCount})`);
  const expected = Array.from({ length: rendered.pageCount - 1 }, (unused, index) => index + 2);
  assert.deepEqual(rowLabelPages(rendered), expected, 'every page the row resumes on is labelled');
  for (let page = 1; page <= rendered.pageCount; page += 1) {
    assert.ok(pagesWith(rendered, 'HEAD_A').includes(page), `page ${page} must still repeat the column header`);
  }
  for (let index = 1; index <= 40; index += 1) {
    assert.equal(pagesWith(rendered, `GIANT_${String(index).padStart(3, '0')}`).length, 1);
  }
});

// ------------------------------------------------------------------ 5. consecutive split rows

test('5: consecutive split rows each label the page they continue onto', async () => {
  const rdl = report();
  const rows = [
    dataRow({ key: 'R1', tall: 18, prefix: 'FIRST' }),
    dataRow({ key: 'R2', tall: 18, prefix: 'SECOND' }),
  ];
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 3);
  const firstResumes = pagesWith(rendered, 'FIRST_018')[0];
  const secondResumes = pagesWith(rendered, 'SECOND_018')[0];
  assert.notEqual(firstResumes, secondResumes, 'the two rows must resume on different pages');
  for (const page of [firstResumes, secondResumes]) {
    assert.ok(rowLabelPages(rendered).includes(page), `page ${page} resumes a row and must be labelled`);
  }
  // ...and a page that only starts a fresh row is not labelled.
  const startsFresh = pagesWith(rendered, 'SECOND_001')[0];
  if (startsFresh !== firstResumes) {
    assert.ok(!rowLabelPages(rendered).includes(startsFresh), 'a page that begins a row is not a continuation');
  }
});

// ------------------------------------------------------------------ 6. fonts, wrapping, padding, margins

test('6: the split cell is measured under its own font, wrapping and padding', async () => {
  const rdl = report({
    marginIn: 0.4,
    columnStyles: { 4: { fontFamily: 'Times New Roman', fontSize: 11, padding: 6 } },
  });
  // Wrapped prose rather than explicit breaks, so the split point is decided by real line wrapping.
  const prose = Array.from({ length: 120 }, (unused, index) => `wrapped-${index + 1}`).join(' ');
  const rows = [{ KEY: 'R1', GRP: 'G1', DEPT: 'D1', A: 'A', B: 'B', C: 'C', D: 'D', E: `WRAPSTART ${prose} WRAPEND` }];
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 2, 'the wrapped cell must overflow its page');
  assert.deepEqual(pagesWith(rendered, 'WRAPSTART'), [1]);
  assert.deepEqual(pagesWith(rendered, 'WRAPEND'), [rendered.pageCount]);
  assert.deepEqual(
    rowLabelPages(rendered),
    Array.from({ length: rendered.pageCount - 1 }, (unused, index) => index + 2),
  );
});

// ------------------------------------------------------------------ 7. page geometry

test('7: changed page size, orientation and margins keep the labels correct', async () => {
  const rows = [dataRow({ key: 'R1', tall: 24, prefix: 'GEO' })];
  const portrait = await render(report({ pageHeightIn: 4, pageWidthIn: 3, marginIn: 0.5, columnWidthIn: 0.4 }), rows);
  const landscape = await render(report({ pageHeightIn: 3, pageWidthIn: 8.5, marginIn: 0.2, columnWidthIn: 1.2 }), rows);
  assert.notEqual(portrait.pageCount, landscape.pageCount, 'the two geometries must paginate differently');
  for (const rendered of [portrait, landscape]) {
    assert.deepEqual(
      rowLabelPages(rendered),
      Array.from({ length: rendered.pageCount - 1 }, (unused, index) => index + 2),
      'the label follows the measured split, whatever the page is',
    );
  }
});

// ------------------------------------------------------------------ 8. rows-per-page varies

test('8: a document whose rows-per-page varies is labelled only where rows really continue', async () => {
  const rdl = report();
  const rows = [
    dataRow({ key: 'R1' }),
    dataRow({ key: 'R2', tall: 3, prefix: 'MID' }),
    dataRow({ key: 'R3' }),
    dataRow({ key: 'R4', tall: 30, prefix: 'LONG' }),
    dataRow({ key: 'R5' }),
    dataRow({ key: 'R6', tall: 2, prefix: 'SHORT' }),
  ];
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 3);
  // Only the row that cannot fit a page at all continues, so only the pages it resumes on are labelled.
  const longPages = pagesWith(rendered, 'LONG_001').concat(pagesWith(rendered, 'LONG_030'));
  const resumed = rowLabelPages(rendered);
  assert.ok(resumed.length > 0, 'the long row must continue somewhere');
  for (const page of resumed) {
    assert.ok(page > Math.min(...longPages), 'labels appear only after the continuing row started');
  }
  for (const key of ['R1', 'R3', 'R5']) {
    assert.equal(pagesWith(rendered, `A_${key}`).length, 1, `${key} renders once`);
  }
});

// ------------------------------------------------------------------ 9. KeepTogether

test('9: a KeepTogether row moves whole and is never labelled', async () => {
  const rdl = report({ keepTogether: true });
  const rows = Array.from({ length: 9 }, (unused, index) => dataRow({ key: `R${index + 1}` }));
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 2, 'the rows must cross a page boundary');
  assert.deepEqual(rowLabelPages(rendered), [], 'a row moved whole is not a continuation');
  for (let index = 1; index <= 9; index += 1) {
    assert.equal(pagesWith(rendered, `A_R${index}`).length, 1, `row ${index} renders on exactly one page`);
  }
});

// ------------------------------------------------------------------ 10. group continues, no split row

test('10: a group spanning pages with no split row is not annotated', async () => {
  const rdl = report({ grouping: 'merged' });
  const rows = Array.from({ length: 12 }, (unused, index) => dataRow({ key: `R${index + 1}`, group: 'EWSS' }));
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 2, 'the group must cross a page boundary');
  assert.deepEqual(rowLabelPages(rendered), [], 'no row was cut, so nothing is a continuation');
  assert.deepEqual(annotatedCells(rendered), [], 'the re-drawn group cell keeps its plain value');
  // The group cell is still re-drawn on every page it spans — only the annotation is absent.
  assert.equal(pagesWith(rendered, 'EWSS').length, rendered.pageCount);
});

// ------------------------------------------------------------------ 11. both labels on one page

test('11: a split row inside a continuing group is labelled, and the group is not', async () => {
  const rdl = report({ grouping: 'merged' });
  const rows = [
    ...Array.from({ length: 4 }, (unused, index) => dataRow({ key: `R${index + 1}`, group: 'EWSS' })),
    dataRow({ key: 'AU572', group: 'EWSS', tall: 24, prefix: 'AU572E' }),
    ...Array.from({ length: 2 }, (unused, index) => dataRow({ key: `R${index + 9}`, group: 'EWSS' })),
    ...Array.from({ length: 3 }, (unused, index) => dataRow({ key: `R${index + 20}`, group: 'OCRO' })),
  ];
  const rendered = await render(rdl, rows);
  const resumePage = pagesWith(rendered, 'AU572E_024')[0];
  assert.ok(resumePage > 1, 'the tall row must continue past its first page');
  assert.ok(rowLabelPages(rendered).includes(resumePage), 'the continued row is labelled');
  assert.deepEqual(annotatedCells(rendered), [], 'the group it sits in is not');
  // Pages of this report that only start rows carry nothing at all.
  const ocroStart = pagesWith(rendered, 'OCRO')[0];
  assert.ok(!rowLabelPages(rendered).includes(ocroStart), 'a page that starts a group and a row is not a continuation');
});

// ------------------------------------------------------------------ 12. boundary coincidence

test('12: a group boundary that lands on a page boundary produces no label at all', async () => {
  // 2.5in page, 0.25in margins: 144pt of body. A 0.25in header plus seven 0.25in rows fills it exactly,
  // so a group of exactly seven rows ends where the page ends and the next group opens the next page.
  const rdl = report({ grouping: 'merged' });
  const rows = [
    ...Array.from({ length: 7 }, (unused, index) => dataRow({ key: `A${index + 1}`, group: 'FIRST' })),
    ...Array.from({ length: 7 }, (unused, index) => dataRow({ key: `B${index + 1}`, group: 'SECOND' })),
  ];
  const rendered = await render(rdl, rows);
  assert.equal(rendered.pageCount, 2, 'the fixture must fill each page exactly');
  assert.deepEqual(pagesWith(rendered, 'FIRST'), [1]);
  assert.deepEqual(pagesWith(rendered, 'SECOND'), [2]);
  assert.deepEqual(rowLabelPages(rendered), []);
  assert.deepEqual(annotatedCells(rendered), []);
});

// ------------------------------------------------------------------ 13. nested groups

test('13: nested groups crossing a boundary are not annotated at any level', async () => {
  const rdl = report({ grouping: 'nested', columns: 3, columnWidthIn: 1.2, groupWidthIn: 0.9 });
  const rows = [
    ...Array.from({ length: 7 }, (unused, index) => dataRow({ key: `I1_${index + 1}`, group: 'OUT', dept: 'IN1', columns: 3 })),
    ...Array.from({ length: 10 }, (unused, index) => dataRow({ key: `I2_${index + 1}`, group: 'OUT', dept: 'IN2', columns: 3 })),
  ];
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 3, `the fixture must cross two boundaries (got ${rendered.pageCount})`);
  // Page 2 continues the outer group with a fresh inner one; page 3 falls inside both. Neither is a row
  // that was cut, so neither page carries anything.
  assert.deepEqual(rowLabelPages(rendered), []);
  assert.deepEqual(annotatedCells(rendered), []);
  assert.ok(pagesWith(rendered, 'OUT').length >= 3, 'the outer group is still re-drawn on every page it spans');
  assert.ok(pagesWith(rendered, 'IN2').includes(2), 'and a fresh inner group still starts where it starts');
});

// ------------------------------------------------------------------ 13b. group header ROW shape

test('13b: a group expressed as a header row is left exactly as the engine draws it', async () => {
  const rdl = report({ grouping: 'headerRow', columns: 3, columnWidthIn: 1.5 });
  const rows = Array.from({ length: 12 }, (unused, index) => dataRow({ key: `R${index + 1}`, group: 'DIVISION', columns: 3 }));
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 2);
  assert.deepEqual(rowLabelPages(rendered), [], 'no row is cut, so nothing is labelled');
  assert.deepEqual(annotatedCells(rendered), []);
  assert.deepEqual(pagesWith(rendered, 'DIVISION'), [1], 'the header row is neither re-drawn nor annotated for the label');
});

// ------------------------------------------------------------------ 15. a structurally different report

test('15: a different column count, no grouping and a different page behave identically', async () => {
  const rdl = report({ columns: 2, pageHeightIn: 3.5, pageWidthIn: 5, marginIn: 0.3, columnWidthIn: 1.8 });
  const rows = [
    dataRow({ key: 'R1', columns: 2, tallColumn: 1 }),
    dataRow({ key: 'R2', columns: 2, tallColumn: 1, tall: 30, prefix: 'TWOCOL' }),
    dataRow({ key: 'R3', columns: 2, tallColumn: 1 }),
  ];
  const rendered = await render(rdl, rows);
  assert.ok(rendered.pageCount >= 2, `the two-column fixture must cross a boundary (got ${rendered.pageCount})`);
  const resumePages = pagesWith(rendered, 'TWOCOL_030');
  assert.ok(rowLabelPages(rendered).includes(resumePages[0]));
  assert.deepEqual(annotatedCells(rendered), [], 'an ungrouped tablix gets row labels only');
  assert.ok(!rowLabelPages(rendered).includes(1));

  // A 20-column tablix is the same feature: nothing keys off the column count.
  const wide = report({ columns: 20, pageWidthIn: 30, columnWidthIn: 1.2 });
  const wideRendered = await render(wide, [dataRow({ key: 'W1', columns: 20, tallColumn: 19, tall: 20, prefix: 'WIDE' })]);
  assert.ok(wideRendered.pageCount >= 2);
  assert.deepEqual(
    rowLabelPages(wideRendered),
    Array.from({ length: wideRendered.pageCount - 1 }, (unused, index) => index + 2),
  );
});

// ------------------------------------------------------------------ 16. two tablixes in one report

test('16: two tablixes splitting on different pages are labelled independently', async () => {
  const rdl = report({ columns: 3, columnWidthIn: 1.5, tablixCount: 2 });
  const rows = [
    dataRow({ key: 'R1', columns: 3, tallColumn: 2 }),
    dataRow({ key: 'R2', columns: 3, tallColumn: 2, tall: 16, prefix: 'BOTH' }),
  ];
  const rendered = await render(rdl, rows);
  const perTablix = new Map();
  for (const label of rowLabels(rendered)) {
    if (!perTablix.has(label.tablixName)) perTablix.set(label.tablixName, []);
    perTablix.get(label.tablixName).push(label.page);
  }
  assert.deepEqual([...perTablix.keys()].sort(), ['T1', 'T2'], 'both tablixes label their own continuation');
  const [first] = perTablix.get('T1');
  const [second] = perTablix.get('T2');
  assert.notEqual(first, second, 'each tablix splits on its own page and labels only that one');
  for (const [name, pages] of perTablix) {
    assert.equal(new Set(pages).size, pages.length, `${name} labels each page at most once`);
  }
});

// ------------------------------------------------------------------ 17. the reported regression

test('17: a long tablix whose rows all fit carries no row label on any page', async () => {
  // The reported defect: a grouped tablix emitted the band on EVERY continuation page, because an open
  // row-span counted as "this row continues". Here 400 rows fit whole across 50+ pages, so not one of them
  // is cut and nothing may be annotated anywhere.
  const grouped = await render(
    report({ grouping: 'merged', groupWidthIn: 2 }),
    Array.from({ length: 400 }, (unused, index) => dataRow({ key: `R${index + 1}`, group: 'ONE' })),
  );
  assert.ok(grouped.pageCount >= 50, `the regression needs a many-page report (got ${grouped.pageCount})`);
  assert.deepEqual(rowLabelPages(grouped), [], 'every page starts a fresh row, so nothing is a continuation');
  assert.deepEqual(annotatedCells(grouped), []);

  // The same data without grouping behaves identically.
  const flat = await render(report(), Array.from({ length: 400 }, (unused, index) => dataRow({ key: `R${index + 1}` })));
  assert.ok(flat.pageCount >= 50);
  assert.deepEqual(rowLabelPages(flat), []);
  assert.deepEqual(annotatedCells(flat), []);
});

// ------------------------------------------------------------------ configuration

test('config: the label can be renamed or turned off, and off means no geometry change', async () => {
  const rdl = report({ grouping: 'merged' });
  const rows = [
    ...Array.from({ length: 4 }, (unused, index) => dataRow({ key: `R${index + 1}`, group: 'EWSS' })),
    dataRow({ key: 'R9', group: 'EWSS', tall: 20, prefix: 'CFG' }),
  ];
  const withConfig = async (continuation) => render(
    rdl,
    rows,
    {},
  ).then(() => renderPdf(
    parseRdl(rdl),
    { outputFileName: 'cfg', parameters: {}, datasets: { D: rows }, pagination: { continuationMarkers: true } },
    { ...config, continuation },
    { captureLayoutTrace: true },
  ));

  const on = await withConfig(config.continuation);
  assert.ok(rowLabelPages(on).length > 0, 'the fixture splits a row');
  assert.deepEqual(annotatedCells(on), [], 'and annotates nothing else');

  const custom = await withConfig({ rowLabel: { enabled: true, text: 'CONTINUED_ROW' } });
  assert.equal(rowLabels(custom)[0].text, 'CONTINUED_ROW');

  // Disabled must be indistinguishable from never asking for markers at all.
  const off = await withConfig({ rowLabel: { enabled: false, text: 'Continued from previous page' } });
  const never = await renderPdf(
    parseRdl(rdl),
    { outputFileName: 'cfg', parameters: {}, datasets: { D: rows } },
    config,
    { captureLayoutTrace: true },
  );
  assert.equal(off.pageCount, never.pageCount);
  assert.deepEqual(
    items(off).map((item) => `${item.kind}:${item.text || ''}:${Math.round(item.y * 100)}`),
    items(never).map((item) => `${item.kind}:${item.text || ''}:${Math.round(item.y * 100)}`),
    'with both labels off the page is laid out exactly as it is without the option',
  );
});
