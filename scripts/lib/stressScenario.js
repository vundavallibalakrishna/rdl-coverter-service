const PAGE_WIDTH = 841.8897637795275;
const PAGE_HEIGHT = 595.275590551181;

function border() {
  return { style: 'Solid', color: '#666666', width: 0.5 };
}

function openBottomBorders(color = '#666666', width = 0.5) {
  const visible = { style: 'Solid', color, width };
  return {
    top: visible,
    right: { ...visible, style: 'Dashed' },
    bottom: { style: 'None', color, width: 0 },
    left: visible,
  };
}

function style(overrides = {}) {
  return {
    color: '#000000',
    backgroundColor: null,
    fontFamily: 'Arial',
    fontSize: 8,
    fontWeight: 'Normal',
    fontStyle: 'Normal',
    textAlign: 'Left',
    verticalAlign: 'Top',
    format: null,
    paddingLeft: 2,
    paddingRight: 2,
    paddingTop: 2,
    paddingBottom: 2,
    border: border(),
    ...overrides,
  };
}

function textbox(name, value, overrides = {}) {
  return {
    type: 'Textbox',
    name,
    top: 0,
    left: 0,
    width: 0,
    height: 0,
    hidden: 'false',
    style: style(overrides),
    pageBreak: null,
    value,
    paragraphs: [[value]],
    canGrow: true,
    canShrink: false,
  };
}

function cell(name, value, overrides = {}, colSpan = 1, rowSpan = 1) {
  return { colSpan, rowSpan, items: [textbox(name, value, overrides)] };
}

function member(overrides = {}) {
  return {
    group: null,
    repeatOnNewPage: false,
    keepWithGroup: 'None',
    fixedData: false,
    hidden: 'false',
    sortExpressions: [],
    children: [],
    ...overrides,
  };
}

export const STRESS_ROW_COUNT = 470;
export const STRESS_OVERFLOW_LINES = 220;

export function createStressScenario() {
  const fields = [
    ['GroupKey', 'System.String'],
    ['SubgroupKey', 'System.String'],
    ['Sequence', 'System.Int32'],
    ['Description', 'System.String'],
    ['OverflowText', 'System.String'],
    ['Amount', 'System.Decimal'],
    ['Status', 'System.String'],
    ['UniqueMarker', 'System.String'],
  ].map(([name, typeName]) => ({ name, dataField: name, typeName }));
  const columns = [60, 72, 50, 120, 220, 65, 70, 123];
  const headerStyle = { backgroundColor: '#1f4e78', color: '#ffffff', fontFamily: 'Segoe UI', fontWeight: 'Bold', textAlign: 'Center', verticalAlign: 'Middle' };
  const subheaderStyle = { backgroundColor: '#d9eaf7', fontFamily: 'Times New Roman', fontWeight: 'Bold', fontStyle: 'Italic', textAlign: 'Center', verticalAlign: 'Middle' };
  const deliberatelyOpenBottom = { borders: openBottomBorders() };
  const rows = [
    {
      height: 20,
      cells: [
        cell('GroupingHeader', 'Grouping', headerStyle, 2),
        cell('WorkHeader', 'Work details', headerStyle, 2),
        cell('MeasuresHeader', 'Measures', headerStyle, 2),
        cell('TraceHeader', 'Traceability', headerStyle, 2),
      ],
    },
    {
      height: 18,
      cells: [
        cell('GroupHeader', 'Group', subheaderStyle, 1, 2),
        cell('SubgroupHeader', 'Subgroup', subheaderStyle, 1, 2),
        cell('SequenceHeader', 'Seq', subheaderStyle, 1, 2),
        cell('DescriptionHeader', 'Description', subheaderStyle, 1, 2),
        cell('MeasureDetailHeader', 'Measure detail', subheaderStyle, 2),
        cell('StatusHeader', 'Status', subheaderStyle, 1, 2),
        cell('MarkerHeader', 'Unique marker', subheaderStyle, 1, 2),
      ],
    },
    {
      height: 20,
      cells: [
        cell('OverflowHeader', 'Overflow / wrapped content', subheaderStyle),
        cell('AmountHeader', 'Amount', subheaderStyle),
      ],
    },
    {
      height: 18,
      cells: [
        cell('GroupCell', '=Fields!GroupKey.Value', { ...deliberatelyOpenBottom, fontWeight: 'Bold' }),
        cell('SubgroupCell', '=Fields!SubgroupKey.Value', { ...deliberatelyOpenBottom, fontFamily: 'Times New Roman', fontStyle: 'Italic' }),
        cell('SequenceCell', '=Fields!Sequence.Value', { ...deliberatelyOpenBottom, fontFamily: 'Segoe UI', textAlign: 'Right', textDecoration: 'Underline' }),
        cell('DescriptionCell', '=Fields!Description.Value', {
          ...deliberatelyOpenBottom,
          fontFamily: '=IIF(Fields!Status.Value = "Escalated", "Times New Roman", IIF(Fields!Status.Value = "Review", "Segoe UI", "Arial"))',
          fontWeight: '=IIF(Fields!Status.Value = "Escalated", "Bold", "Normal")',
        }),
        cell('OverflowCell', '=Fields!OverflowText.Value', { ...deliberatelyOpenBottom }),
        cell('AmountCell', '=Format(Fields!Amount.Value, "N2")', { ...deliberatelyOpenBottom, fontFamily: 'Times New Roman', textAlign: 'Right' }),
        cell('StatusCell', '=Fields!Status.Value', {
          ...deliberatelyOpenBottom,
          fontFamily: 'Segoe UI',
          fontWeight: 'Bold',
          textAlign: 'Center',
          backgroundColor: '=IIF(Fields!Status.Value = "Escalated", "Red", IIF(Fields!Status.Value = "Review", "Yellow", IIF(Fields!Status.Value = "Closed", "Lime", "White")))',
        }),
        cell('MarkerCell', '=Fields!UniqueMarker.Value', { ...deliberatelyOpenBottom, color: '#17365d' }),
      ],
    },
  ];
  const groupedMember = member({
    group: { name: 'GroupRows', expressions: ['=Fields!GroupKey.Value'], pageBreak: 'None' },
    sortExpressions: [{ value: '=Fields!GroupKey.Value', direction: 'Ascending' }],
    children: [member({
      group: { name: 'SubgroupRows', expressions: ['=Fields!SubgroupKey.Value'], pageBreak: 'None' },
      sortExpressions: [
        { value: '=Fields!SubgroupKey.Value', direction: 'Ascending' },
        { value: '=Fields!Sequence.Value', direction: 'Ascending' },
      ],
    })],
  });
  const table = {
    type: 'Tablix',
    name: 'CertificationTable',
    top: 0,
    left: 0,
    width: columns.reduce((sum, width) => sum + width, 0),
    height: 60,
    hidden: 'false',
    style: style(),
    pageBreak: null,
    datasetName: 'StressData',
    datasetFields: fields,
    columns,
    rows,
    rowMembers: [
      member({ repeatOnNewPage: true, keepWithGroup: 'After' }),
      member({ repeatOnNewPage: true, keepWithGroup: 'After' }),
      member({ repeatOnNewPage: true, keepWithGroup: 'After' }),
      groupedMember,
    ],
    columnMembers: columns.map(() => member()),
    repeatColumnHeaders: true,
    repeatRowHeaders: false,
    filters: [],
    sortExpressions: [],
  };
  const endMarker = textbox('StressDocumentEnd', 'STRESS_DOCUMENT_END', {
    fontFamily: 'Segoe UI',
    fontSize: 16,
    fontWeight: 'Bold',
    color: '#1f4e78',
    border: { style: 'None', color: '#000000', width: 0 },
  });
  Object.assign(endMarker, {
    top: 100,
    left: 0,
    width: 500,
    height: 30,
    pageBreak: { location: 'Start', disabled: 'false' },
  });
  const header = textbox('PageHeader', 'RDL TABLE STRESS CERTIFICATION', {
    fontSize: 11,
    fontWeight: 'Bold',
    color: '#ffffff',
    backgroundColor: '#17365d',
    verticalAlign: 'Middle',
    border: { style: 'None', color: '#000000', width: 0 },
  });
  Object.assign(header, { width: 780, height: 22 });
  const footer = textbox('PageFooter', 'Generated by deterministic RDL stress certification', {
    fontSize: 7,
    color: '#555555',
    border: { style: 'None', color: '#000000', width: 0 },
  });
  Object.assign(footer, { width: 780, height: 14 });
  const model = {
    namespace: 'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition',
    name: 'RDL table stress certification',
    page: {
      width: PAGE_WIDTH,
      height: PAGE_HEIGHT,
      marginTop: 28,
      marginRight: 28,
      marginBottom: 28,
      marginLeft: 28,
      header: { height: 24, printOnFirstPage: true, printOnLastPage: true, items: [header] },
      footer: { height: 18, printOnFirstPage: true, printOnLastPage: true, items: [footer] },
    },
    body: { height: 200, width: 780, items: [table, endMarker] },
    datasets: [{ name: 'StressData', fields, parameterNames: [], hasQuery: false }],
    parameters: [],
    renderingDatasets: ['StressData'],
    parameterDatasets: [],
    embeddedImages: {},
    fonts: ['Arial', 'Times New Roman', 'Segoe UI'],
    features: { textboxes: 22, tablixes: 1, rectangles: 0, lines: 0, images: 0, groups: 2, pageBreaks: 1 },
    unsupported: [],
  };

  const rowsData = Array.from({ length: STRESS_ROW_COUNT }, (_, index) => {
    const sequence = index + 1;
    const group = Math.floor(index / 40) + 1;
    const subgroup = Math.floor((index % 40) / 10) + 1;
    let overflowText = sequence === 1
      ? `GIANT_CELL_START\n${Array.from({ length: STRESS_OVERFLOW_LINES }, (_, line) => `GIANT_LINE_${String(line + 1).padStart(3, '0')} deterministic wrapped content`).join('\n')}\nGIANT_CELL_END`
      : `Normal wrapped content for row ${String(sequence).padStart(4, '0')}.`;
    if (sequence !== 1 && sequence % 13 === 0) {
      overflowText = Array.from({ length: 5 }, (_, line) => `ROW_${String(sequence).padStart(4, '0')}_WRAP_${line + 1}`).join('\n');
    }
    return {
      GroupKey: `GROUP_${String(group).padStart(2, '0')}`,
      SubgroupKey: `SUBGROUP_${String(group).padStart(2, '0')}_${String(subgroup).padStart(2, '0')}`,
      Sequence: sequence,
      Description: sequence % 17 === 0 ? `Long description ${sequence}: grouping, sorting, merged headers, wrapping, conditional style, and pagination.` : `Description ${sequence}`,
      OverflowText: overflowText,
      Amount: sequence * 17.35,
      Status: ['Open', 'Review', 'Closed', 'Escalated'][index % 4],
      UniqueMarker: `UNIQUE_ROW_${String(sequence).padStart(4, '0')}`,
    };
  });
  const request = {
    outputFileName: 'rdl-table-stress-certification',
    parameters: {},
    datasets: { StressData: rowsData },
    pagination: { continuationMarkers: true },
  };
  return { model, request };
}
