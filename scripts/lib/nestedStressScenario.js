const PAGE_WIDTH = 841.8897637795275;
const PAGE_HEIGHT = 595.275590551181;

const LEVEL_COUNT = 5;
const PORTFOLIO_COUNT = 3;
const CHILDREN_PER_LEVEL = 2;
const DETAILS_PER_CONTROL = 2;

function border(style = 'Solid', color = '#5b6573', width = 0.5) {
  return { style, color, width };
}

function borders(color = '#5b6573', bottomStyle = 'None') {
  return {
    top: border('Solid', color),
    right: border('Dashed', color),
    bottom: border(bottomStyle, color, bottomStyle === 'None' ? 0 : 0.5),
    left: border('Solid', color),
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
    borders: borders(),
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
    keepTogether: false,
    keepWithGroup: 'None',
    fixedData: false,
    hideIfNoRows: false,
    hidden: 'false',
    sortExpressions: [],
    header: null,
    children: [],
    ...overrides,
  };
}

function leafPaths(members, parent = [], target = []) {
  for (const entry of members) {
    const path = [...parent, entry];
    if (entry.children.length) leafPaths(entry.children, path, target);
    else target.push(path);
  }
  return target;
}

const levelFields = ['Portfolio', 'Division', 'Domain', 'Process', 'Control'];
const levelColors = ['#d9eaf7', '#e2f0d9', '#fff2cc', '#fce4d6', '#e4dfec'];
const levelFonts = ['Segoe UI', 'Arial', 'Times New Roman', 'Segoe UI', 'Arial'];

function groupHeaderCell(level) {
  const field = levelFields[level];
  return cell(`${field}RowHeader`, `=Fields!${field}.Value`, {
    backgroundColor: levelColors[level],
    fontFamily: levelFonts[level],
    fontWeight: 'Bold',
    verticalAlign: 'Middle',
  });
}

function groupTemplate(level, role) {
  const field = levelFields[level];
  const markerField = `L${level + 1}${role === 'header' ? 'Header' : 'Footer'}`;
  const value = role === 'header'
    ? `=Fields!${markerField}.Value & " | " & Fields!${field}.Value`
    : `=Fields!${markerField}.Value & " | Total " & Format(Sum(Fields!Amount.Value, "L${level + 1}Group"), "N2")`;
  return {
    height: role === 'header' ? 16 : 15,
    hidden: 'false',
    cells: [cell(`L${level + 1}${role}Cell`, value, {
      backgroundColor: role === 'header' ? levelColors[level] : '#f2f2f2',
      fontFamily: levelFonts[level],
      fontWeight: 'Bold',
      fontStyle: role === 'footer' && level % 2 ? 'Italic' : 'Normal',
      color: role === 'header' ? '#17365d' : '#404040',
      borders: borders('#5b6573'),
    }, 6)],
  };
}

function nestedHierarchy(level = 0) {
  const field = levelFields[level];
  const headerLeaf = member();
  const footerLeaf = member();
  const child = level === LEVEL_COUNT - 1
    ? member({ group: { name: 'Details', expressions: [], pageBreak: 'None', parent: null, filters: [] } })
    : nestedHierarchy(level + 1);
  return member({
    group: {
      name: `L${level + 1}Group`,
      expressions: [`=Fields!${field}.Value`],
      pageBreak: level === 0 ? 'Between' : 'None',
      parent: null,
      filters: [],
    },
    keepTogether: level >= 3,
    sortExpressions: [{ value: `=Fields!${field}.Value`, direction: 'Ascending' }],
    header: { size: (LEVEL_COUNT - level) * 54, cell: groupHeaderCell(level) },
    children: [headerLeaf, child, footerLeaf],
  });
}

function staticHeaders() {
  const top = { backgroundColor: '#17365d', color: '#ffffff', fontFamily: 'Segoe UI', fontWeight: 'Bold', textAlign: 'Center', verticalAlign: 'Middle', borders: borders('#ffffff', 'Solid') };
  const band = { backgroundColor: '#2f75b5', color: '#ffffff', fontFamily: 'Arial', fontWeight: 'Bold', textAlign: 'Center', verticalAlign: 'Middle', borders: borders('#ffffff', 'Solid') };
  const labels = { backgroundColor: '#d9eaf7', fontFamily: 'Times New Roman', fontWeight: 'Bold', fontStyle: 'Italic', textAlign: 'Center', verticalAlign: 'Middle', borders: borders('#5b6573', 'Solid') };
  return [
    { height: 21, hidden: 'false', cells: [cell('NestedTitle', 'DEEP NESTED RDL PAGINATION CERTIFICATION', top, 11)] },
    { height: 18, hidden: 'false', cells: [
      cell('HierarchyBand', 'Five-level hierarchy', band, 5),
      cell('EvidenceBand', 'Evidence and measures', band, 3),
      cell('DecisionBand', 'Decision and traceability', band, 3),
    ] },
    { height: 24, hidden: 'false', cells: [
      ...['Portfolio', 'Division', 'Domain', 'Process', 'Control'].map((label, index) => cell(`Label${index}`, label, labels)),
      cell('SequenceLabel', 'Seq', labels),
      cell('NarrativeLabel', 'Narrative', labels),
      cell('OverflowLabel', 'Overflow / wrapping', labels),
      cell('AmountLabel', 'Amount', labels),
      cell('StatusLabel', 'Status', labels),
      cell('MarkerLabel', 'Unique marker', labels),
    ] },
  ];
}

function detailTemplate() {
  const open = { borders: borders('#5b6573') };
  return {
    height: 17,
    hidden: '=Fields!Visibility.Value = "Hidden"',
    cells: [
      cell('SequenceCell', '=Fields!Sequence.Value', { ...open, fontFamily: 'Segoe UI', textAlign: 'Right', textDecoration: 'Underline' }),
      cell('NarrativeCell', '=Fields!Narrative.Value', {
        ...open,
        fontFamily: '=IIF(Fields!Status.Value = "Escalated", "Times New Roman", IIF(Fields!Status.Value = "Review", "Segoe UI", "Arial"))',
        fontWeight: '=IIF(Fields!Status.Value = "Escalated", "Bold", "Normal")',
      }),
      cell('OverflowCell', '=Fields!OverflowText.Value', { ...open }),
      cell('AmountCell', '=Format(Fields!Amount.Value, "N2")', { ...open, fontFamily: 'Times New Roman', textAlign: 'Right' }),
      cell('StatusCell', '=Fields!Status.Value', {
        ...open,
        fontFamily: 'Segoe UI',
        fontWeight: 'Bold',
        textAlign: 'Center',
        backgroundColor: '=IIF(Fields!Status.Value = "Escalated", "Red", IIF(Fields!Status.Value = "Review", "Yellow", IIF(Fields!Status.Value = "Closed", "Lime", "White")))',
      }),
      cell('MarkerCell', '=Fields!UniqueMarker.Value', { ...open, color: '#17365d' }),
    ],
  };
}

function hierarchyTemplates(level = 0, target = []) {
  target.push(groupTemplate(level, 'header'));
  if (level === LEVEL_COUNT - 1) target.push(detailTemplate());
  else hierarchyTemplates(level + 1, target);
  target.push(groupTemplate(level, 'footer'));
  return target;
}

function headerAndFooter() {
  const header = textbox('PageHeader', 'DEEP NESTED TABLE STRESS CERTIFICATION', {
    fontFamily: 'Segoe UI', fontSize: 11, fontWeight: 'Bold', color: '#ffffff', backgroundColor: '#17365d',
    verticalAlign: 'Middle', borders: { top: border('None'), right: border('None'), bottom: border('None'), left: border('None') },
  });
  Object.assign(header, { width: 785, height: 22 });
  const footer = textbox('PageFooter', 'Nested hierarchy certification | Page expressions and borders remain grounded', {
    fontFamily: 'Arial', fontSize: 7, color: '#555555',
    borders: { top: border('None'), right: border('None'), bottom: border('None'), left: border('None') },
  });
  Object.assign(footer, { width: 785, height: 14 });
  return { header, footer };
}

function createRowsData() {
  const rows = [];
  const exactMarkers = [];
  let sequence = 0;
  for (let portfolio = 1; portfolio <= PORTFOLIO_COUNT; portfolio += 1) {
    for (let division = 1; division <= CHILDREN_PER_LEVEL; division += 1) {
      for (let domain = 1; domain <= CHILDREN_PER_LEVEL; domain += 1) {
        for (let process = 1; process <= CHILDREN_PER_LEVEL; process += 1) {
          for (let control = 1; control <= CHILDREN_PER_LEVEL; control += 1) {
            for (let detail = 1; detail <= DETAILS_PER_CONTROL; detail += 1) {
              sequence += 1;
              const keys = [portfolio, division, domain, process, control];
              const identifiers = keys.map((value, index) => `${levelFields[index].toUpperCase()}_${keys.slice(0, index + 1).map((part) => String(part).padStart(2, '0')).join('_')}`);
              const compactKey = (depth) => keys.slice(0, depth + 1).map((part) => String(part).padStart(2, '0')).join('');
              const headerMarkers = identifiers.map((identifier, index) => `NH${index + 1}_${compactKey(index)}`);
              const footerMarkers = identifiers.map((identifier, index) => `NF${index + 1}_${compactKey(index)}`);
              const uniqueMarker = `NR${String(sequence).padStart(4, '0')}`;
              let overflowText = `Normal nested content for sequence ${String(sequence).padStart(4, '0')}.`;
              const giant = sequence === 1 ? { id: 'A', lines: 60 }
                : sequence === Math.floor(NESTED_STRESS_ROW_COUNT / 2) + 1 ? { id: 'B', lines: 70 }
                  : sequence === PORTFOLIO_COUNT * 32 ? { id: 'C', lines: 80 }
                    : null;
              if (giant) {
                overflowText = `G${giant.id}_START\n${Array.from({ length: giant.lines }, (_, line) => `G${giant.id}${String(line + 1).padStart(3, '0')} deterministic content`).join('\n')}\nG${giant.id}_END`;
              } else if (sequence % 7 === 0) {
                overflowText = Array.from({ length: 5 }, (_, line) => `NW${String(sequence).padStart(4, '0')}_${line + 1}`).join('\n');
              } else if (sequence % 5 === 0) {
                overflowText = `NW${String(sequence).padStart(4, '0')}_1\nNW${String(sequence).padStart(4, '0')}_2`;
              }
              rows.push({
                Portfolio: identifiers[0], Division: identifiers[1], Domain: identifiers[2], Process: identifiers[3], Control: identifiers[4],
                Sequence: sequence,
                Narrative: sequence % 11 === 0
                  ? `Nested narrative ${sequence}: merged ancestors, subtotal scopes, conditional formats, and pagination boundaries.`
                  : `Nested narrative ${sequence}`,
                OverflowText: overflowText,
                Amount: sequence * 11.75,
                Status: ['Open', 'Review', 'Closed', 'Escalated'][sequence % 4],
                Visibility: sequence % 53 === 0 ? 'Hidden' : 'Visible',
                UniqueMarker: uniqueMarker,
                ...Object.fromEntries(headerMarkers.map((marker, index) => [`L${index + 1}Header`, marker])),
                ...Object.fromEntries(footerMarkers.map((marker, index) => [`L${index + 1}Footer`, marker])),
              });
              if (sequence % 53 !== 0) exactMarkers.push(uniqueMarker);
              if (detail === 1) {
                exactMarkers.push(headerMarkers[4], footerMarkers[4]);
                if (control === 1) {
                  exactMarkers.push(headerMarkers[3], footerMarkers[3]);
                  if (process === 1) {
                    exactMarkers.push(headerMarkers[2], footerMarkers[2]);
                    if (domain === 1) {
                      exactMarkers.push(headerMarkers[1], footerMarkers[1]);
                      if (division === 1) exactMarkers.push(headerMarkers[0], footerMarkers[0]);
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  return { rows, exactMarkers };
}

export const NESTED_STRESS_ROW_COUNT = PORTFOLIO_COUNT * (CHILDREN_PER_LEVEL ** 4) * DETAILS_PER_CONTROL;

export function createNestedStressScenario() {
  const bodyColumns = [45, 145, 145, 55, 65, 60];
  const rowHeaderColumns = [54, 54, 54, 54, 54];
  const fields = [
    ...levelFields,
    'Sequence', 'Narrative', 'OverflowText', 'Amount', 'Status', 'Visibility', 'UniqueMarker',
    ...Array.from({ length: LEVEL_COUNT }, (_, index) => `L${index + 1}Header`),
    ...Array.from({ length: LEVEL_COUNT }, (_, index) => `L${index + 1}Footer`),
  ].map((name) => ({ name, dataField: name, typeName: name === 'Sequence' ? 'System.Int32' : name === 'Amount' ? 'System.Decimal' : 'System.String' }));
  const staticMembers = [member(), member(), member()];
  const rowMembers = [...staticMembers, nestedHierarchy()];
  const rows = [...staticHeaders(), ...hierarchyTemplates()];
  const table = {
    type: 'Tablix', name: 'DeepNestedCertificationTable', top: 0, left: 0,
    width: [...rowHeaderColumns, ...bodyColumns].reduce((sum, width) => sum + width, 0), height: 120,
    hidden: 'false', style: style(), pageBreak: null, datasetName: 'NestedData', datasetFields: fields,
    bodyColumns, rowHeaderColumns, columns: [...rowHeaderColumns, ...bodyColumns], rows, rowMembers,
    rowMemberPaths: leafPaths(rowMembers),
    columnMembers: bodyColumns.map(() => member()),
    columnMemberPaths: bodyColumns.map((entry, index) => [member({ key: `${entry}:${index}` })]),
    hasColumnGroups: false, tablixCorner: [], repeatColumnHeaders: true, repeatRowHeaders: true,
    filters: [], sortExpressions: [],
  };
  const { rows: rowsData, exactMarkers } = createRowsData();
  const endMarker = textbox('NestedStressDocumentEnd', 'NESTED_STRESS_DOCUMENT_END', {
    fontFamily: 'Segoe UI', fontSize: 16, fontWeight: 'Bold', color: '#17365d',
    borders: { top: border('None'), right: border('None'), bottom: border('None'), left: border('None') },
  });
  Object.assign(endMarker, { top: 150, left: 0, width: 500, height: 30, pageBreak: { location: 'Start', disabled: 'false' } });
  const { header, footer } = headerAndFooter();
  const model = {
    namespace: 'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition',
    name: 'Deep nested RDL table stress certification',
    page: {
      width: PAGE_WIDTH, height: PAGE_HEIGHT, marginTop: 28, marginRight: 28, marginBottom: 28, marginLeft: 28,
      header: { height: 24, printOnFirstPage: true, printOnLastPage: true, items: [header] },
      footer: { height: 18, printOnFirstPage: true, printOnLastPage: true, items: [footer] },
    },
    body: { height: 220, width: 785, items: [table, endMarker] },
    datasets: [{ name: 'NestedData', fields, parameterNames: [], hasQuery: false }],
    parameters: [], renderingDatasets: ['NestedData'], parameterDatasets: [], embeddedImages: {},
    fonts: ['Arial', 'Times New Roman', 'Segoe UI'],
    features: { textboxes: 40, tablixes: 1, rectangles: 0, lines: 0, images: 0, groups: LEVEL_COUNT, pageBreaks: 2 },
    unsupported: [],
  };
  const overflowSpecs = [
    { id: 'A', lines: 60, start: 'GA_START', end: 'GA_END', linePrefix: 'GA' },
    { id: 'B', lines: 70, start: 'GB_START', end: 'GB_END', linePrefix: 'GB' },
    { id: 'C', lines: 80, start: 'GC_START', end: 'GC_END', linePrefix: 'GC' },
  ];
  return {
    model,
    request: {
      outputFileName: 'rdl-deep-nested-stress-certification', parameters: {}, datasets: { NestedData: rowsData },
      pagination: { continuationMarkers: true },
    },
    certification: {
      id: 'rdl-deep-nested-stress',
      targetPages: { minimum: 30, maximum: 40 },
      exactMarkers,
      overflowSpecs,
      endMarker: 'NESTED_STRESS_DOCUMENT_END',
      expectedHeaderRows: 3,
      expectedFonts: model.fonts,
      expectedRowCount: NESTED_STRESS_ROW_COUNT,
    },
  };
}
