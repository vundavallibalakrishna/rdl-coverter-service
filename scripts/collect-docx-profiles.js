import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);

if (args.includes('--help') || args.includes('-h')) {
  process.stdout.write(`Usage: node scripts/collect-docx-profiles.js [input-dir] [output.json]

Collects per-report docx-profile-candidate.json files from certification output folders and writes one
validated profile bundle. Defaults to tmp/output/docx-certification-word-local/docx-profiles.candidate.json.

Arguments:
  input-dir     Directory containing one subdirectory per certification run.
  output.json   Combined profile bundle path.
`);
  process.exit(0);
}

const inputRoot = path.resolve(args[0] || path.join(serviceRoot, 'tmp', 'output', 'docx-certification-word-local'));
const outputPath = path.resolve(args[1] || path.join(inputRoot, 'docx-profiles.candidate.json'));

async function readJsonIfExists(filePath) {
  try {
    return JSON.parse(await fs.readFile(filePath, 'utf8'));
  } catch {
    return null;
  }
}

const profiles = [];
for (const entry of await fs.readdir(inputRoot, { withFileTypes: true })) {
  if (!entry.isDirectory()) continue;
  const reportPath = path.join(inputRoot, entry.name, 'docx-certification-report.json');
  const profilePath = path.join(inputRoot, entry.name, 'docx-profile-candidate.json');
  const profileDocument = await readJsonIfExists(profilePath);
  const report = await readJsonIfExists(reportPath);
  for (const profile of profileDocument?.profiles || []) {
    profiles.push({
      ...profile,
      source: path.relative(serviceRoot, reportPath),
      renderer: report?.renderer || null,
      bestVariant: report?.bestVariant || null,
      certified: profile.certified === true && report?.certified === true,
    });
  }
}

profiles.sort((left, right) => String(left.id).localeCompare(String(right.id)));
const duplicateIds = [...new Set(profiles.map((profile) => profile.id)
  .filter((id, index, ids) => ids.indexOf(id) !== index))];
if (duplicateIds.length > 0) throw new Error(`Duplicate DOCX profile ids: ${duplicateIds.join(', ')}`);

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, `${JSON.stringify({ profiles }, null, 2)}\n`);
process.stdout.write(`${JSON.stringify({ outputPath, profiles: profiles.length, ids: profiles.map((profile) => profile.id) }, null, 2)}\n`);
