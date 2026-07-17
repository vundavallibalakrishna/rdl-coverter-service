import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// Client RDLs, request fixtures, and row data are client property: they carry real report definitions,
// queries, and row values, so they must never enter version control. They live under tmp/ (git-ignored in
// full) rather than in the repo, which means a fresh clone legitimately does not have them — tests and
// scripts that need them must skip rather than fail. RDL_SAMPLES_DIR overrides the location for a machine
// that keeps them elsewhere.
const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');

export const samplesRoot = path.resolve(process.env.RDL_SAMPLES_DIR || path.join(serviceRoot, 'tmp', 'samples'));

export function samplePath(...segments) {
  return path.join(samplesRoot, ...segments);
}

/** True only when every named sample is present, so callers can skip cleanly on a fresh clone. */
export function hasSamples(...files) {
  return files.every((file) => fs.existsSync(samplePath(file)));
}

export const MISSING_SAMPLES = `Client samples are not present in ${samplesRoot} (see README "Working files")`;
