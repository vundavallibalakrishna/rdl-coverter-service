import fs from 'node:fs/promises';
import path from 'node:path';
import { parseRdl } from '../src/rdl/parser.js';
import { samplesRoot } from './lib/samples.js';

async function readChunkedJsonCsv(filePath) {
  const source = (await fs.readFile(filePath, 'utf8')).replace(/^\uFEFF/, '').trim();
  const chunks = source.split(/\r?\n/).map((line, index) => {
    if (!line.startsWith('"') || !line.endsWith('"')) {
      throw new Error(`${path.basename(filePath)} chunk ${index + 1} is not a quoted JSON fragment`);
    }
    return line.slice(1, -1);
  });
  const rows = JSON.parse(chunks.join(''));
  if (!Array.isArray(rows)) throw new Error(`${path.basename(filePath)} does not contain a JSON array`);
  return rows;
}

function normalizeRows(model, datasetName, rows) {
  const dataset = model.datasets.find((candidate) => candidate.name === datasetName);
  if (!dataset) throw new Error(`RDL dataset not found: ${datasetName}`);
  return rows.map((row) => Object.fromEntries(dataset.fields.map((field) => [
    field.dataField,
    Object.hasOwn(row, field.dataField) ? row[field.dataField] : null,
  ])));
}

function derivePieChartRows(mainRows) {
  const statuses = ['Open', 'Closed', 'New', 'Not Assessed', 'Under Investigation'];
  const total = mainRows.length;
  return statuses.map((status) => {
    const count = mainRows.filter((row) => row.Status === status).length;
    const percentage = total === 0 ? 0 : (count / total) * 100;
    return {
      'Action Status': status,
      StatusCount: count,
      TotalCount: total,
      Percentage: percentage,
      AdjustedPercentage: Math.round(percentage),
    };
  });
}

function deriveIncidentCreatedRows(mainRows) {
  return mainRows.map((row, index) => ({
    NameM: row['Month/Year'] ?? null,
    Months: row.Order ?? null,
    'Incident Types': row['Incident Type'] ?? null,
    UniverseID: null,
    'Domain Name': row['Division/Department'] ?? null,
    AssessID: null,
    'Assessment Name': row.Located ?? null,
    'Month/Year': row['Month/Year'] ?? null,
    Order: row.Order ?? null,
    PRiskInstID: row.PRiskInstID ?? null,
    HistoryID: null,
    'Incident Name': row['Risk Name'] ?? null,
    Classification: row.Classification ?? null,
    'Incident Type': row['Incident Type'] ?? null,
    Status: row.Status ?? null,
    Row: index + 1,
  }));
}

async function writeRequest(fileName, request) {
  const outputPath = path.join(samplesRoot, fileName);
  await fs.writeFile(outputPath, `${JSON.stringify(request, null, 2)}\n`);
  return outputPath;
}

const combinedRdlName = 'Combined Assurance Reports Excel.rdl';
const combinedModel = parseRdl(await fs.readFile(path.join(samplesRoot, combinedRdlName)));
const combinedSourceRows = await readChunkedJsonCsv(path.join(samplesRoot, 'Combined_assurances_test 1.csv'));
const combinedMainRows = normalizeRows(combinedModel, 'MainData', combinedSourceRows);
const combinedPath = await writeRequest('combined-assurance-excel-request.json', {
  fileName: combinedRdlName,
  output: 'PDF',
  outputFileName: 'combined-assurance-excel',
  parameters: {
    userid: 1,
    Division: ['All'],
    Domain: ['All'],
    Assessment: ['All'],
    DateFrom: '2026-01-01T00:00:00Z',
    DateTo: '2026-12-31T23:59:59Z',
    Submitted: 'System Administrator',
  },
  datasets: {
    MainData: combinedMainRows,
    intro: [],
    User: [{ DisplayName: 'System Administrator' }],
  },
});

const incidentRdlName = 'Incident Dashboard Report.rdl';
const incidentModel = parseRdl(await fs.readFile(path.join(samplesRoot, incidentRdlName)));
const incidentSourceRows = await readChunkedJsonCsv(path.join(samplesRoot, 'Incident Dashboard.csv'));
const incidentMainRows = normalizeRows(incidentModel, 'MainData', incidentSourceRows);
const incidentPath = await writeRequest('incident-dashboard-request.json', {
  fileName: incidentRdlName,
  output: 'PDF',
  outputFileName: 'incident-dashboard',
  parameters: {
    userid: '1',
    Domain: ['All'],
    Assessment: ['All'],
    Year: 2026,
    Month: ['All'],
    Submitted: 'System Administrator',
  },
  datasets: {
    MainData: incidentMainRows,
    User: [{ DisplayName: 'System Administrator' }],
    PieChart: normalizeRows(incidentModel, 'PieChart', derivePieChartRows(incidentSourceRows)),
    IncidentCreated: normalizeRows(incidentModel, 'IncidentCreated', deriveIncidentCreatedRows(incidentSourceRows)),
  },
});

process.stdout.write(`${JSON.stringify({
  combinedAssurance: { path: combinedPath, mainRows: combinedMainRows.length },
  incidentDashboard: {
    path: incidentPath,
    mainRows: incidentMainRows.length,
    pieRows: 5,
    incidentCreatedRows: incidentSourceRows.length,
  },
}, null, 2)}\n`);
