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

async function collectDocxReports(root) {
  const absoluteRoot = path.join(serviceRoot, root);
  const reports = [];
  try {
    for (const entry of await fs.readdir(absoluteRoot, { withFileTypes: true })) {
      if (!entry.isDirectory()) continue;
      const reportPath = path.join(absoluteRoot, entry.name, 'docx-certification-report.json');
      const report = await readJsonIfExists(reportPath);
      if (!report) continue;
      const best = report.variants?.find((variant) => variant.name === report.bestVariant);
      reports.push({
        report: entry.name,
        renderer: report.renderer,
        certified: report.certified,
        certificationBlocked: report.certificationBlocked,
        bestVariant: report.bestVariant,
        nativePageFragments: report.profileCandidate?.docx?.nativePageFragments ?? null,
        canonicalPages: report.canonicalPdf?.pages ?? null,
        wordExportPages: best?.pages ?? null,
        pageCountMatchesPdf: best?.pageCountMatchesPdf ?? null,
        dimensionsMatchPdf: best?.dimensionsMatchPdf ?? null,
        textCoverage: best?.textCoverage?.ratio ?? null,
        nativeTables: best?.openXml?.nativeTables ?? null,
        bodyPositionedShapes: best?.openXml?.bodyPositionedShapes ?? null,
        blockingDifferences: report.blockingDifferences || [],
      });
    }
  } catch {
    // Missing certification directory is reported as an empty set.
  }
  return reports.sort((left, right) => left.report.localeCompare(right.report));
}

const schema = await readJsonIfExists(path.join(serviceRoot, 'tmp/output/rdl-2016-capability-catalogue.json'));
const stress = await readJsonIfExists(path.join(serviceRoot, 'tmp/output/stress-certification.json'));
const libreOfficeReports = await collectDocxReports('tmp/output/docx-certification-local');
const wordReports = await collectDocxReports('tmp/output/docx-certification-word-local');

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
    libreOffice: libreOfficeReports,
  },
  remainingBlockers: [
    ...(!wordReports.length ? ['Run npm run certify:docx with --renderer=word for local Word certification evidence.'] : []),
    ...(!wordReports.some((report) => report.certified) ? ['Formal SSRS certification still requires exact SSRS reference PDF, parameters, dataset rows, and licensed font versions from the same run.'] : []),
  ],
};

process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
