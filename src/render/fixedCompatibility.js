import { checkFonts } from './fonts.js';

function countItems(items = []) {
  let count = 0;
  for (const item of items) {
    count += 1;
    if (item.items) count += countItems(item.items);
    if (item.rows) {
      for (const row of item.rows) {
        for (const cell of row.cells || []) count += countItems(cell.items || []);
      }
    }
  }
  return count;
}

export function analyzeFixedEditableCompatibility(model, config) {
  const fonts = checkFonts(config, model.fonts.length ? model.fonts : ['Arial']);
  const staticObjectCount = countItems([
    ...(model.page.header?.items || []),
    ...(model.body.items || []),
    ...(model.page.footer?.items || []),
  ]);
  const analyzeTimePasses = model.unsupported.length === 0 && fonts.ready;
  return {
    // Provisional: this reflects only the checks possible without rendering (RDL-level support + fonts).
    // Fixed-editable re-parses the generated PDF, so gradients/shading, rotated text or transformed images,
    // mixed page sizes, and the per-document object/page/image limits can still cause a render-time
    // UNSUPPORTED_FEATURE even when this is true. Treat `compatible` as necessary, not sufficient.
    compatible: analyzeTimePasses,
    verdict: analyzeTimePasses ? 'provisional-pass' : 'reject',
    renderMayReject: analyzeTimePasses,
    missingFonts: fonts.missing,
    unsupportedPdfOperators: [],
    pdfOperatorScan: 'deferred-until-render',
    // The real object count scales with rendered cells x borders x pages, so this design-time item count is
    // a floor, not a prediction of the maxFixedObjects limit.
    estimatedObjectCount: staticObjectCount,
    estimateIncludesDatasetRows: false,
    deferredRenderChecks: ['pdfOperators', 'objectLimit', 'pageLimit', 'imageLimit', 'textRunLimit', 'mixedPageSizes'],
    limits: {
      pages: config.maxFixedPages,
      objects: config.maxFixedObjects,
      images: config.maxFixedImages,
      textRuns: config.maxFixedTextRuns,
    },
  };
}

