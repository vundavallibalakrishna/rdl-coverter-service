#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';

const args = process.argv.slice(2);
const docxPath = args[0];
const outIndex = args.indexOf('--out');
const outputPath = outIndex >= 0 ? args[outIndex + 1] : null;
if (!docxPath || !outputPath) {
  console.error('Usage: audit_docx_ooxml.mjs <docx> --out tmp/<audit>.json');
  process.exit(2);
}
const serviceRoot = path.resolve(path.dirname(new URL(import.meta.url).pathname), '../../..');
const tmpRoot = path.join(serviceRoot, 'tmp');
const resolvedOutput = path.resolve(outputPath);
if (path.dirname(resolvedOutput) !== tmpRoot) throw new Error('Audit must be written directly under repository tmp/');

const zip = await JSZip.loadAsync(await fs.readFile(docxPath));
const read = async (name) => zip.file(name)?.async('string') || '';
const [documentXml, fontTableXml, settingsXml] = await Promise.all([
  read('word/document.xml'),
  read('word/fontTable.xml'),
  read('word/settings.xml'),
]);
const names = Object.keys(zip.files);
const storyPartNames = names.filter((name) => /^word\/(?:document|header\d+|footer\d+)\.xml$/i.test(name));
const storyParts = await Promise.all(storyPartNames.map(read));
const drawingIds = storyParts.flatMap((xml) => (
  [...xml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)].map((match) => match[1])
));
const duplicateDrawingIds = drawingIds.length - new Set(drawingIds).size;
const exactLineInlineDrawings = storyParts.flatMap((xml) => (
  [...xml.matchAll(/<w:p(?:\s|>)[\s\S]*?<\/w:p>/g)]
    .map((match) => match[0])
    .filter((paragraph) => /<wp:inline(?:\s|>)/.test(paragraph))
    .map((paragraph) => {
      const spacing = paragraph.match(/<w:spacing\b[^>]*\/>/)?.[0] || '';
      const line = Number(spacing.match(/\bw:line="(\d+)"/)?.[1] || 0);
      const exact = /\bw:lineRule="exact"/.test(spacing);
      const heightEmu = Number(paragraph.match(/<wp:extent\b[^>]*\bcy="(\d+)"/)?.[1] || 0);
      return { exact, lineTwips: line, heightEmu };
    })
    .filter((drawing) => drawing.exact && drawing.lineTwips > 0 && drawing.heightEmu > 0)
));
// 1 twip = 635 EMUs. An inline object larger than an exact line is clipped by Word until the user
// manually enlarges the row/paragraph, even when the difference originated from harmless-looking
// point-to-pixel rounding.
const oversizedExactLineInlineDrawings = exactLineInlineDrawings.filter(
  (drawing) => drawing.heightEmu > drawing.lineTwips * 635,
).length;
const relationshipParts = names.filter((name) => /(?:^|\/)_rels\/.+\.rels$/i.test(name));
const relationshipXml = await Promise.all(relationshipParts.map(read));
const externalRelationships = relationshipXml.reduce(
  (count, xml) => count + (xml.match(/TargetMode="External"/g) || []).length,
  0,
);
const embeddedFontParts = names.filter((name) => /^word\/fonts\/.+\.odttf$/i.test(name)).length;
const embeddedRegular = (fontTableXml.match(/<w:embedRegular\b/g) || []).length;
const embeddedBold = (fontTableXml.match(/<w:embedBold\b/g) || []).length;
const embeddedItalic = (fontTableXml.match(/<w:embedItalic\b/g) || []).length;
const embeddedBoldItalic = (fontTableXml.match(/<w:embedBoldItalic\b/g) || []).length;
const report = {
  schemaVersion: 1,
  file: path.basename(docxPath),
  nativeTextRuns: (documentXml.match(/<w:t(?:\s|>)/g) || []).length,
  nativeTables: (documentXml.match(/<w:tbl>/g) || []).length,
  sections: (documentXml.match(/<w:sectPr(?:\s|>)/g) || []).length,
  exactRows: (documentXml.match(/<w:trHeight[^>]*w:hRule="exact"/g) || []).length,
  drawings: (documentXml.match(/<w:drawing>/g) || []).length,
  packageDrawingIds: drawingIds.length,
  duplicateDrawingIds,
  exactLineInlineDrawings: exactLineInlineDrawings.length,
  oversizedExactLineInlineDrawings,
  vmlShapes: (documentXml.match(/<v:shape(?:\s|>)/g) || []).length,
  fullPageVisualImages: (documentXml.match(/behindDoc="1"/g) || []).length,
  embeddedFontParts,
  embeddedRegular,
  embeddedBold,
  embeddedItalic,
  embeddedBoldItalic,
  liveFields: (documentXml.match(/<w:(?:instrText|fldSimple)\b/g) || []).length,
  updateFieldsOnOpen: /<w:updateFields\b[^>]*w:val="(?:true|1)"/i.test(settingsXml),
  relationshipParts: relationshipParts.length,
  externalRelationships,
};
report.passed = report.nativeTextRuns > 0
  && report.nativeTables > 0
  && report.sections > 0
  && report.embeddedRegular > 0
  && report.embeddedBold === report.embeddedRegular
  && report.embeddedItalic === report.embeddedRegular
  && report.embeddedBoldItalic === report.embeddedRegular
  && report.embeddedFontParts === report.embeddedRegular * 4
  && report.vmlShapes === 0
  && report.fullPageVisualImages === 0
  && report.duplicateDrawingIds === 0
  && report.oversizedExactLineInlineDrawings === 0
  && report.liveFields === 0
  && report.updateFieldsOnOpen === false
  && report.externalRelationships === 0;
await fs.writeFile(resolvedOutput, `${JSON.stringify(report, null, 2)}\n`, { mode: 0o600 });
console.log(resolvedOutput);
if (!report.passed) process.exitCode = 1;
