import fs from 'node:fs';
import { ServiceError } from '../errors.js';

const PROFILE_TOP_LEVEL_KEYS = new Set([
  'id',
  'description',
  'certified',
  'match',
  'docx',
  'note',
  'source',
  'renderer',
  'bestVariant',
]);
const PROFILE_MATCH_KEYS = new Set(['definitionSha256', 'name', 'namespace']);
const PROFILE_DOCX_KEYS = new Set(['nativePageFragments']);
const PROFILE_ID_PATTERN = /^[A-Za-z0-9._:-]{1,128}$/;

function asArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function failProfile(reason, index) {
  const suffix = Number.isInteger(index) ? ` at index ${index}` : '';
  throw new Error(`Invalid DOCX profile${suffix}: ${reason}`);
}

function validateString(value, field, index, { required = false, pattern = null } = {}) {
  if (value === undefined || value === null || value === '') {
    if (required) failProfile(`${field} is required`, index);
    return;
  }
  if (typeof value !== 'string') failProfile(`${field} must be a string`, index);
  if (pattern && !pattern.test(value)) failProfile(`${field} has an invalid format`, index);
}

function validateProfile(profile, index, seenIds) {
  if (!profile || typeof profile !== 'object' || Array.isArray(profile)) failProfile('entry must be an object', index);
  for (const key of Object.keys(profile)) {
    if (!PROFILE_TOP_LEVEL_KEYS.has(key)) failProfile(`unknown property '${key}'`, index);
  }
  validateString(profile.id, 'id', index, { required: true, pattern: PROFILE_ID_PATTERN });
  if (seenIds.has(profile.id)) failProfile(`duplicate id '${profile.id}'`, index);
  seenIds.add(profile.id);

  validateString(profile.description, 'description', index);
  validateString(profile.note, 'note', index);
  validateString(profile.source, 'source', index);
  validateString(profile.renderer, 'renderer', index);
  validateString(profile.bestVariant, 'bestVariant', index);
  if (profile.certified !== undefined && typeof profile.certified !== 'boolean') failProfile('certified must be a boolean', index);

  if (!profile.match || typeof profile.match !== 'object' || Array.isArray(profile.match)) failProfile('match must be an object', index);
  for (const key of Object.keys(profile.match)) {
    if (!PROFILE_MATCH_KEYS.has(key)) failProfile(`unknown match property '${key}'`, index);
  }
  validateString(profile.match.definitionSha256, 'match.definitionSha256', index, { pattern: /^[a-f0-9]{64}$/i });
  validateString(profile.match.name, 'match.name', index);
  validateString(profile.match.namespace, 'match.namespace', index);
  if (!profile.match.definitionSha256 && !profile.match.name && !profile.match.namespace) {
    failProfile('match must include definitionSha256, name, or namespace', index);
  }

  if (profile.docx !== undefined) {
    if (!profile.docx || typeof profile.docx !== 'object' || Array.isArray(profile.docx)) failProfile('docx must be an object', index);
    for (const key of Object.keys(profile.docx)) {
      if (!PROFILE_DOCX_KEYS.has(key)) failProfile(`unknown docx property '${key}'`, index);
    }
    if (profile.docx.nativePageFragments !== undefined && typeof profile.docx.nativePageFragments !== 'boolean') {
      failProfile('docx.nativePageFragments must be a boolean', index);
    }
  }
}

function configuredProfiles(config = {}) {
  if (!config.docxProfilePath) return { profiles: [], error: null };
  try {
    const parsed = JSON.parse(fs.readFileSync(config.docxProfilePath, 'utf8'));
    if (parsed === null || typeof parsed !== 'object') failProfile('root must be an object or array');
    if (!Array.isArray(parsed) && parsed.profiles !== undefined && !Array.isArray(parsed.profiles)) {
      failProfile('profiles must be an array');
    }
    const profiles = asArray(parsed.profiles ?? parsed);
    const seenIds = new Set();
    profiles.forEach((profile, index) => validateProfile(profile, index, seenIds));
    return { profiles, error: null };
  } catch (error) {
    return { profiles: [], error: error.message || 'Unable to load DOCX profile configuration' };
  }
}

function profileMatches(profile, model) {
  const match = profile.match || {};
  if (match.definitionSha256 && match.definitionSha256 !== model.identity?.definitionSha256) return false;
  if (match.name && match.name !== model.identity?.name && match.name !== model.name) return false;
  if (match.namespace && match.namespace !== model.namespace) return false;
  return Boolean(match.definitionSha256 || match.name || match.namespace);
}

function profileSummary(profile) {
  if (!profile) return null;
  return {
    id: profile.id || null,
    description: profile.description || null,
    certified: profile.certified === true,
    match: {
      definitionSha256: profile.match?.definitionSha256 || null,
      name: profile.match?.name || null,
      namespace: profile.match?.namespace || null,
    },
    docx: {
      nativePageFragments: profile.docx?.nativePageFragments === true,
    },
  };
}

export function matchStructuredDocxProfile(model, request = {}, config = {}, { strictRequested = false } = {}) {
  const requestedId = request.docx?.profile || request.docxProfile || null;
  const { profiles, error } = configuredProfiles(config);
  if (error && (strictRequested || requestedId || config.docxProfileAuto)) {
    throw new ServiceError('CONFIG_INVALID', 'DOCX profile configuration is invalid', 500, { profilePath: config.docxProfilePath });
  }
  const matches = profiles.filter((profile) => profileMatches(profile, model));
  const requested = requestedId
    ? profiles.find((profile) => profile.id === requestedId)
    : null;
  if (requestedId && !requested) {
    throw new ServiceError('PARAMETER_INVALID', `DOCX profile '${requestedId}' was not found`, 400);
  }
  if (requested && !profileMatches(requested, model)) {
    throw new ServiceError('PARAMETER_INVALID', `DOCX profile '${requestedId}' does not match this RDL`, 400);
  }
  const autoSelected = config.docxProfileAuto ? matches.find((profile) => profile.certified === true) : null;
  const selected = requested || autoSelected || null;
  return {
    available: profiles.length,
    error,
    requested: requestedId,
    autoApply: config.docxProfileAuto === true,
    matches: matches.map(profileSummary),
    selected: profileSummary(selected),
    selectedProfile: selected || null,
  };
}

export function resolveStructuredDocxOptions(model, request = {}, config = {}) {
  const profile = matchStructuredDocxProfile(model, request, config, { strictRequested: true });
  const explicit = request.docx?.nativePageFragments;
  const legacyExplicit = request.docxNativePageFragments;
  const nativePageFragments = explicit === true || explicit === false
    ? explicit
    : legacyExplicit === true || legacyExplicit === false
      ? legacyExplicit
      : profile.selectedProfile?.docx?.nativePageFragments === true
        ? true
        : config?.docxNativePageFragments === true;
  return {
    nativePageFragments,
    profile: {
      available: profile.available,
      requested: profile.requested,
      autoApply: profile.autoApply,
      selected: profile.selected,
    },
  };
}

function walkItems(items = [], visitor) {
  for (const item of items || []) {
    visitor(item);
    if (item.items) walkItems(item.items, visitor);
    if (item.rows) {
      for (const row of item.rows || []) {
        for (const cell of row.cells || []) walkItems(cell.items || [], visitor);
      }
    }
  }
}

function walkMembers(members = [], visitor, depth = 0) {
  for (const member of members || []) {
    visitor(member, depth);
    walkMembers(member.children || [], visitor, depth + 1);
  }
}

function memberStats(members = []) {
  const stats = {
    count: 0,
    grouped: 0,
    maxDepth: 0,
    repeatOnNewPage: 0,
    keepTogether: 0,
    keepWithGroup: 0,
    fixedData: 0,
    pageBreakGroups: 0,
    parentGroups: 0,
  };
  walkMembers(members, (member, depth) => {
    stats.count += 1;
    stats.maxDepth = Math.max(stats.maxDepth, depth + 1);
    if (member.group?.expressions?.length) stats.grouped += 1;
    if (member.repeatOnNewPage) stats.repeatOnNewPage += 1;
    if (member.keepTogether) stats.keepTogether += 1;
    if (member.keepWithGroup && member.keepWithGroup !== 'None') stats.keepWithGroup += 1;
    if (member.fixedData) stats.fixedData += 1;
    if (member.group?.parent) stats.parentGroups += 1;
    if (/^(Start|End|Between|StartAndEnd)$/i.test(String(member.group?.pageBreak || 'None'))) stats.pageBreakGroups += 1;
  });
  return stats;
}

function tablixTemplateStats(tablix) {
  const stats = {
    rows: tablix.rows?.length || 0,
    columns: tablix.columns?.length || 0,
    rowSpanCells: 0,
    columnSpanCells: 0,
    maxRowSpan: 1,
    maxColumnSpan: 1,
    headerRows: 0,
  };
  for (const row of tablix.rows || []) {
    if (row.isHeader) stats.headerRows += 1;
    for (const cell of row.cells || []) {
      const rowSpan = cell.rowSpan || 1;
      const colSpan = cell.colSpan || 1;
      if (rowSpan > 1) stats.rowSpanCells += 1;
      if (colSpan > 1) stats.columnSpanCells += 1;
      stats.maxRowSpan = Math.max(stats.maxRowSpan, rowSpan);
      stats.maxColumnSpan = Math.max(stats.maxColumnSpan, colSpan);
    }
  }
  return stats;
}

function addRisk(risks, code, severity, message) {
  risks.push({ code, severity, message });
}

function highestSeverity(risks) {
  if (risks.some((risk) => risk.severity === 'high')) return 'high';
  if (risks.some((risk) => risk.severity === 'medium')) return 'medium';
  if (risks.some((risk) => risk.severity === 'low')) return 'low';
  return 'none';
}

export function shouldUseNativePageFragments(model, request = {}, config = {}) {
  return resolveStructuredDocxOptions(model, request, config).nativePageFragments;
}

export function analyzeStructuredEditableCompatibility(model, config = {}, request = {}) {
  const risks = [];
  const stats = {
    tablixes: 0,
    charts: 0,
    images: 0,
    rowSpanCells: 0,
    columnSpanCells: 0,
    maxRowSpan: 1,
    maxColumnSpan: 1,
    repeatedHeaderMembers: 0,
    groupedMembers: 0,
    nestedGroupDepth: 0,
    pageBreakGroups: 0,
    keepTogetherMembers: 0,
    matrixTablixes: 0,
  };

  walkItems([
    ...(model.page.header?.items || []),
    ...(model.body.items || []),
    ...(model.page.footer?.items || []),
  ], (item) => {
    if (item.type === 'Chart') stats.charts += 1;
    if (item.type === 'Image') stats.images += 1;
    if (item.type !== 'Tablix') return;

    stats.tablixes += 1;
    if (item.hasColumnGroups) stats.matrixTablixes += 1;
    const template = tablixTemplateStats(item);
    stats.rowSpanCells += template.rowSpanCells;
    stats.columnSpanCells += template.columnSpanCells;
    stats.maxRowSpan = Math.max(stats.maxRowSpan, template.maxRowSpan);
    stats.maxColumnSpan = Math.max(stats.maxColumnSpan, template.maxColumnSpan);

    const rowMembers = memberStats(item.rowMembers);
    const columnMembers = memberStats(item.columnMembers);
    stats.repeatedHeaderMembers += rowMembers.repeatOnNewPage + columnMembers.repeatOnNewPage;
    stats.groupedMembers += rowMembers.grouped + columnMembers.grouped;
    stats.nestedGroupDepth = Math.max(stats.nestedGroupDepth, rowMembers.maxDepth, columnMembers.maxDepth);
    stats.pageBreakGroups += rowMembers.pageBreakGroups + columnMembers.pageBreakGroups;
    stats.keepTogetherMembers += rowMembers.keepTogether + columnMembers.keepTogether;
  });

  if (stats.tablixes > 0) {
    addRisk(risks, 'WORD_REPAGINATES_NATIVE_TABLES', 'medium',
      'DOCX_EDITABLE uses real Word tables, so Microsoft Word owns final pagination and can differ from the PDF.');
  }
  if (stats.rowSpanCells > 0 || stats.groupedMembers > 0 || stats.nestedGroupDepth > 1) {
    addRisk(risks, 'ROW_GROUP_SPANS_CAN_DRIFT', 'high',
      'Nested row groups and merged row-span blocks are editable, but Word may move/split them differently from SSRS/PDF.');
  }
  if (stats.repeatedHeaderMembers > 0) {
    addRisk(risks, 'REPEATED_HEADERS_VIEWER_DEPENDENT', 'medium',
      'Repeated tablix headers are emitted as Word table headers, but their exact repeat points follow Word pagination.');
  }
  if (stats.pageBreakGroups > 0 || (model.features?.pageBreaks || 0) > 0) {
    addRisk(risks, 'EXPLICIT_PAGE_BREAKS_CAN_SHIFT', 'medium',
      'RDL page breaks are represented in Word, but surrounding native table layout can shift the resulting page boundary.');
  }
  if (stats.matrixTablixes > 0) {
    addRisk(risks, 'DYNAMIC_COLUMNS_CAN_REFLOW', 'medium',
      'Matrix/dynamic-column tablixes are expanded into native table grids; Word may redistribute tight column content.');
  }
  if (stats.charts > 0) {
    addRisk(risks, 'CHARTS_ARE_RASTER_IMAGES', 'low',
      'Charts remain visible but are embedded as images in DOCX_EDITABLE, so chart labels are not normal editable text.');
  }
  if ((model.page.header?.items?.length || 0) > 0 || (model.page.footer?.items?.length || 0) > 0) {
    addRisk(risks, 'RDL_PAGE_BANDS_ARE_POSITIONED', 'low',
      'RDL page headers and footers are coordinate-based; the structured DOCX renderer positions those bands rather than converting them into body tables.');
  }

  const riskLevel = highestSeverity(risks);
  const profile = matchStructuredDocxProfile(model, request, config);
  const fragmentRecommendation = (() => {
    if (stats.tablixes === 0) return 'not_needed';
    if (stats.rowSpanCells > 0 || stats.groupedMembers > 0 || stats.nestedGroupDepth > 1) return 'avoid_unless_certified';
    if (stats.repeatedHeaderMembers > 0 || (model.features?.pageBreaks || 0) > 0) return 'try_and_certify';
    return 'off_by_default';
  })();

  return {
    compatible: model.unsupported.length === 0,
    layoutMode: 'structured',
    nativeBodyTables: true,
    exactPageParity: false,
    riskLevel,
    risks,
    stats,
    nativePageFragments: {
      supported: true,
      enabledByDefault: config?.docxNativePageFragments === true,
      requestOptions: ['docx.nativePageFragments', 'docxNativePageFragments', 'docx.profile'],
      recommendation: fragmentRecommendation,
      note: 'This keeps real Word tables. It is a per-report certification knob, not a universal PDF-matching guarantee.',
    },
    profiles: {
      available: profile.available,
      error: profile.error ? 'DOCX profile configuration is invalid' : null,
      autoApply: profile.autoApply,
      matches: profile.matches,
      selected: profile.selected,
    },
    recommendation: riskLevel === 'high'
      ? 'Use DOCX_EDITABLE for maximum normal Word editability, then certify page-by-page. If PDF resemblance is more important than table editing, compare DOCX_FIXED_EDITABLE.'
      : 'Use DOCX_EDITABLE for the client-facing editable document and certify with a Word-exported PDF comparison.',
  };
}
