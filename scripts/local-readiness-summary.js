import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);

if (args.includes('--help') || args.includes('-h')) {
  process.stdout.write(`Usage: node scripts/local-readiness-summary.js

Prints a JSON summary of local engineering readiness, stress results, schema coverage, and available DOCX
certification reports. It does not render new artifacts and does not mark SSRS certification complete unless
an existing Word certification report is certified.
`);
  process.exit(0);
}

async function readJsonIfExists(filePath) {
  try {
    return JSON.parse(await fs.readFile(filePath, 'utf8'));
  } catch {
    return null;
  }
}

async function collectWindowsWordReports() {
  const absoluteRoot = path.join(serviceRoot, 'tmp');
  const reports = [];
  try {
    for (const entry of await fs.readdir(absoluteRoot, { withFileTypes: true })) {
      if (!entry.isFile() || !entry.name.endsWith('-windows-word-certification.json')) continue;
      const report = await readJsonIfExists(path.join(absoluteRoot, entry.name));
      if (!report) continue;
      reports.push({
        report: entry.name.replace(/-windows-word-certification\.json$/, ''),
        renderer: report.renderer,
        layoutMode: report.layoutMode,
        authority: report.authority,
        certified: report.passed === true,
        gates: report.gates || {},
        input: report.input || {},
      });
    }
  } catch {
    // Missing tmp/ or certification artifacts are reported as an empty set.
  }
  return reports.sort((left, right) => left.report.localeCompare(right.report));
}

const schema = await readJsonIfExists(path.join(serviceRoot, 'tmp/rdl-2016-capability-catalogue.json'));
const stress = await readJsonIfExists(path.join(serviceRoot, 'tmp/rdl-table-stress-certification.json'));
const wordReports = await collectWindowsWordReports();

const summary = {
  localEngineeringReady: Boolean(stress?.passed && schema),
  formalSsrsCertified: wordReports.some((report) => report.certified),
  schemaCoverage: schema ? {
    supported: schema.SUPPORTED ?? schema.summary?.SUPPORTED,
    metadataOnly: schema.METADATA_ONLY ?? schema.summary?.METADATA_ONLY,
    rejected: schema.REJECTED ?? schema.summary?.REJECTED,
    total: schema.total ?? schema.summary?.total,
  } : null,
  stress: stress ? {
    passed: stress.passed,
    directPdfPages: stress.directPdf?.pages,
    structuredDocxPages: stress.editableDocx?.renderedPages,
    structuredDocxPageRangeAdvisoryPassed: stress.editableDocx?.pageRangeAdvisoryPassed,
    nativeTables: stress.editableDocx?.openXml?.nativeTables,
    rowMarkersPassed: stress.editableDocx?.rowMarkers?.passed,
    overflowPassed: stress.editableDocx?.overflow?.passed,
  } : null,
  docxCertification: {
    word: wordReports,
  },
  remainingBlockers: [
    ...(!wordReports.length ? ['Run npm run certify:windows-word on a Windows QA host with desktop Microsoft Word.'] : []),
    ...(!wordReports.some((report) => report.certified) ? ['Formal SSRS certification still requires exact SSRS reference PDF, parameters, dataset rows, and licensed font versions from the same run.'] : []),
  ],
};

process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
