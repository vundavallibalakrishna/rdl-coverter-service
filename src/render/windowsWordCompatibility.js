import { ServiceError } from '../errors.js';
import { fontEmbeddingEligibility } from './fonts.js';

const WORD_MAX_PAGE_POINTS = 22 * 72;
const WORD_MAX_TABLE_COLUMNS = 63;
const PRECISION_POINTS = 0.25;

function hasOwn(object, property) {
  return Boolean(object) && Object.prototype.hasOwnProperty.call(object, property);
}

// There is no alternate editable-DOCX pagination mode anymore. Rejecting the former switches is deliberate:
// silently ignoring one would let a caller believe it selected continuous or profile-driven behavior.
export function validateWindowsWordRequest(request = {}) {
  const obsolete = [];
  if (hasOwn(request.docx, 'nativePageFragments')) obsolete.push('docx.nativePageFragments');
  if (hasOwn(request, 'docxNativePageFragments')) obsolete.push('docxNativePageFragments');
  if (hasOwn(request.docx, 'profile')) obsolete.push('docx.profile');
  if (hasOwn(request, 'docxProfile')) obsolete.push('docxProfile');
  if (obsolete.length > 0) {
    throw new ServiceError(
      'RDL_INVALID',
      `Obsolete DOCX options are not supported: ${obsolete.join(', ')}`,
      400,
      {
        obsolete,
        layoutMode: 'windows-paged-editable',
      },
    );
  }
}

function walkItems(items, visitor) {
  for (const item of items || []) {
    visitor(item);
    walkItems(item.items, visitor);
    for (const row of item.rows || []) {
      for (const cell of row.cells || []) walkItems(cell.items, visitor);
    }
  }
}

function snap(value) {
  return Math.round(Number(value || 0) / PRECISION_POINTS) * PRECISION_POINTS;
}

function coordinateBoundaryEstimate(model) {
  const boundaries = new Set([0, snap(model.page.width)]);
  const origins = [
    { x: model.page.marginLeft, items: model.page.header?.items || [] },
    { x: model.page.marginLeft, items: model.body?.items || [] },
    { x: model.page.marginLeft, items: model.page.footer?.items || [] },
  ];
  for (const origin of origins) {
    walkItems(origin.items, (item) => {
      if (!Number.isFinite(item.left) || !Number.isFinite(item.width)) return;
      boundaries.add(snap(origin.x + item.left));
      boundaries.add(snap(origin.x + item.left + item.width));
      if (item.type === 'Tablix') {
        let x = origin.x + item.left;
        for (const column of item.columns || []) {
          x += typeof column === 'number' ? column : column.width || 0;
          boundaries.add(snap(x));
        }
      }
    });
  }
  return boundaries.size - 1;
}

function modelStats(model) {
  const stats = {
    tablixes: 0,
    charts: 0,
    images: 0,
    subreports: model.features?.subreports || 0,
    literalRotations: 0,
    expressionWritingModes: 0,
    largestDeclaredTableGrid: 0,
    dynamicColumnTablixes: 0,
    estimatedPageGridColumns: coordinateBoundaryEstimate(model),
  };
  walkItems([
    ...(model.page.header?.items || []),
    ...(model.body?.items || []),
    ...(model.page.footer?.items || []),
  ], (item) => {
    if (item.type === 'Chart') stats.charts += 1;
    if (item.type === 'Image') stats.images += 1;
    if (item.type === 'Tablix') {
      stats.tablixes += 1;
      stats.largestDeclaredTableGrid = Math.max(stats.largestDeclaredTableGrid, item.columns?.length || 0);
      if (item.hasColumnGroups) stats.dynamicColumnTablixes += 1;
    }
    const writingMode = item.style?.writingMode;
    if (typeof writingMode === 'string' && writingMode.startsWith('=')) stats.expressionWritingModes += 1;
    else if (writingMode && !/^default$/i.test(String(writingMode))) stats.literalRotations += 1;
  });
  return stats;
}

export function analyzeWindowsWordCompatibility(model, config = {}, request = {}) {
  validateWindowsWordRequest(request);
  const stats = modelStats(model);
  const page = {
    widthPt: model.page.width,
    heightPt: model.page.height,
    maximumPt: WORD_MAX_PAGE_POINTS,
    eligible: model.page.width <= WORD_MAX_PAGE_POINTS && model.page.height <= WORD_MAX_PAGE_POINTS,
  };
  const tableGrid = {
    maximumColumns: WORD_MAX_TABLE_COLUMNS,
    largestDeclaredColumns: stats.largestDeclaredTableGrid,
    estimatedPageGridColumns: stats.estimatedPageGridColumns,
    eligibleAtAnalysis: stats.largestDeclaredTableGrid <= WORD_MAX_TABLE_COLUMNS
      && stats.estimatedPageGridColumns <= WORD_MAX_TABLE_COLUMNS,
    runtimeTraceRequired: stats.dynamicColumnTablixes > 0,
  };
  const fontEmbedding = fontEmbeddingEligibility(config, model.fonts || []);
  const fontsEligible = fontEmbedding.every((font) => font.eligible);
  const unsupported = [];
  if (!page.eligible) unsupported.push({
    code: 'WORD_PAGE_SIZE_LIMIT',
    message: 'The declared page exceeds Microsoft Word’s 22-by-22-inch limit.',
  });
  if (!tableGrid.eligibleAtAnalysis) unsupported.push({
    code: 'WORD_TABLE_COLUMN_LIMIT',
    message: 'The declared or estimated fixed page grid exceeds Microsoft Word’s 63-column table limit.',
  });
  if (stats.literalRotations > 0) unsupported.push({
    code: 'WORD_EDITABLE_ROTATION',
    message: 'The report declares rotated/vertical text that is not supported by the page-locked editable contract.',
  });
  for (const font of fontEmbedding.filter((entry) => entry.blocksWindowsPagedEditable)) {
    unsupported.push({
      code: 'FONT_EMBEDDING_FORBIDDEN',
      message: `Font '${font.family}' is missing a required variant or cannot be embedded for editing.`,
    });
  }
  return {
    renderer: 'DOCX_EDITABLE',
    platform: 'Microsoft Word for Windows',
    layoutMode: 'windows-paged-editable',
    pdfLayoutAuthority: true,
    compatibleAtAnalysis: unsupported.length === 0 && fontsEligible,
    runtimeTraceValidationRequired: true,
    page,
    tableGrid,
    fontEmbedding,
    stats,
    unsupported,
    limits: {
      pageInches: 22,
      tableColumns: WORD_MAX_TABLE_COLUMNS,
      geometryPrecisionPt: PRECISION_POINTS,
    },
  };
}
