import { ServiceError } from '../errors.js';
import { resolveExcelLayoutMode } from '../excelLayoutMode.js';
import { parseRdl } from './parser.js';
import { parameterSignature, resolveParameterValues } from './validation.js';

const MAX_SUBREPORT_DEFINITIONS = 32;
const MAX_SUBREPORT_DEPTH = 8;

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function canonicalReportName(value) {
  const normalized = String(value || '')
    .trim()
    .replace(/\\/g, '/')
    .replace(/\/+/g, '/')
    .replace(/\.rdl$/i, '');
  return normalized.toLowerCase();
}

function decodeRdlBase64(value, config) {
  if (typeof value !== 'string' || value.length === 0) {
    throw new ServiceError('RDL_INVALID', 'Bundled subreport rdlBase64 is required');
  }
  const normalized = value.replace(/^data:[^;]+;base64,/, '').replace(/\s+/g, '');
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(normalized)) {
    throw new ServiceError('RDL_INVALID', 'Bundled subreport rdlBase64 is invalid');
  }
  const buffer = Buffer.from(normalized, 'base64');
  if (buffer.length === 0) throw new ServiceError('RDL_INVALID', 'Bundled subreport RDL is empty');
  if (buffer.length > config.maxRdlBytes) {
    throw new ServiceError('RDL_INVALID', 'Bundled subreport RDL exceeds the configured size limit', 413);
  }
  return buffer;
}

function visitItems(items, visitor) {
  const visitMembers = (members) => {
    for (const member of members || []) {
      if (member.header) visitItems(member.header.cell.items, visitor);
      visitMembers(member.children);
    }
  };
  for (const item of items || []) {
    visitor(item);
    if (item.type === 'Tablix') {
      visitMembers(item.rowMembers);
      visitMembers(item.columnMembers);
      for (const cornerRow of item.tablixCorner || []) {
        for (const cornerCell of cornerRow) visitItems(cornerCell.items, visitor);
      }
    }
    if (item.items) visitItems(item.items, visitor);
    if (item.rows) {
      for (const row of item.rows) {
        for (const cell of row.cells) visitItems(cell.items, visitor);
      }
    }
  }
}

function stripResolvedSubreportBlockers(model) {
  const unresolved = [];
  visitItems([
    ...model.body.items,
    ...(model.page.header?.items || []),
    ...(model.page.footer?.items || []),
  ], (item) => {
    if (item.type === 'Subreport' && !item.resolvedSubreport && !item.resolveBundledSubreport) unresolved.push(item);
  });
  if (unresolved.length > 0) return;
  model.unsupported = model.unsupported.filter((feature) => (
    feature !== 'Subreport'
    && feature !== 'TablixCellContent:Subreport'
    && !(feature.startsWith('RdlPath:') && /(?:^|\.)Subreport(?:\.|$)/.test(feature))
  ));
}

function bundleMap(raw, config) {
  if (!isPlainObject(raw)) return new Map();
  const entries = Object.entries(raw);
  if (entries.length > MAX_SUBREPORT_DEFINITIONS) {
    throw new ServiceError(
      'RDL_INVALID',
      `Subreport bundle exceeds the supported limit of ${MAX_SUBREPORT_DEFINITIONS} definitions`,
      413,
    );
  }
  const result = new Map();
  let totalBytes = 0;
  for (const [reportName, definition] of entries) {
    if (!isPlainObject(definition)) {
      throw new ServiceError('RDL_INVALID', `Bundled subreport ${reportName} must be an object`);
    }
    const canonical = canonicalReportName(reportName);
    if (!canonical) throw new ServiceError('RDL_INVALID', 'Bundled subreport name is required');
    if (result.has(canonical)) {
      throw new ServiceError('RDL_INVALID', `Duplicate bundled subreport name: ${reportName}`);
    }
    const rdl = decodeRdlBase64(definition.rdlBase64, config);
    totalBytes += rdl.length;
    if (totalBytes > config.maxRequestBytes) {
      throw new ServiceError('RDL_INVALID', 'Bundled subreport definitions exceed the configured request size limit', 413);
    }
    if (!Array.isArray(definition.instances) || definition.instances.length === 0) {
      throw new ServiceError('RDL_INVALID', `Bundled subreport ${reportName} requires at least one invocation instance`);
    }
    result.set(canonical, { reportName, rdl, instances: definition.instances, used: false, model: null });
  }
  return result;
}

function bundledRenderingRowCount(model, datasets) {
  return model.renderingDatasets.reduce((total, datasetName) => (
    total + (Array.isArray(datasets?.[datasetName]) ? datasets[datasetName].length : 0)
  ), 0);
}

/**
 * Resolves every Subreport item against caller-supplied definitions without executing report-server paths
 * or child queries. Each child invocation carries its own concrete parameters and exact DataField rows.
 * Parsed child models are cached by canonical ReportName and attached to the normalized Subreport item.
 */
export function resolveBundledSubreports(model, request, config) {
  const bundles = bundleMap(request.subreports, config);
  let bundledRows = 0;
  const fonts = new Set(model.fonts);
  let hasSubreports = false;
  visitItems([
    ...model.body.items,
    ...(model.page.header?.items || []),
    ...(model.page.footer?.items || []),
  ], (item) => {
    if (item.type === 'Subreport') hasSubreports = true;
  });
  const output = String(request.output || '').toUpperCase();
  const supportsBundledSubreports = ['PDF', 'DOCX_VISUAL', 'DOCX_EDITABLE'].includes(output)
    || (output === 'XLSX' && resolveExcelLayoutMode(request) === 'REPORT');
  if (hasSubreports && !supportsBundledSubreports) {
    throw new ServiceError(
      'UNSUPPORTED_FEATURE',
      output === 'XLSX'
        ? 'Bundled subreports require Excel REPORT mode'
        : 'Bundled subreports are supported only for PDF, DOCX, and Excel REPORT output',
      400,
      { features: ['Subreport'] },
    );
  }

  const resolveDefinition = (canonical, lineage) => {
    const definition = bundles.get(canonical);
    if (!definition) {
      throw new ServiceError(
        'UNSUPPORTED_FEATURE',
        `Bundled definition is missing for subreport: ${canonical}`,
        400,
        { features: ['Subreport'] },
      );
    }
    if (lineage.includes(canonical)) {
      throw new ServiceError('RDL_INVALID', `Subreport cycle detected: ${[...lineage, canonical].join(' -> ')}`);
    }
    if (lineage.length >= MAX_SUBREPORT_DEPTH) {
      throw new ServiceError(
        'UNSUPPORTED_FEATURE',
        `Subreport nesting depth exceeds the supported limit of ${MAX_SUBREPORT_DEPTH}`,
      );
    }
    definition.used = true;
    if (definition.model) return definition;

    const child = parseRdl(definition.rdl, {
      maxRdlBytes: config.maxRdlBytes,
      maxXmlNodes: config.maxXmlNodes,
      maxXmlDepth: config.maxXmlDepth,
    });
    definition.model = child;
    definition.limits = config;
    for (const font of child.fonts) fonts.add(font);
    visitItems([
      ...child.body.items,
      ...(child.page.header?.items || []),
      ...(child.page.footer?.items || []),
    ], (item) => {
      if (item.type !== 'Subreport') return;
      const childCanonical = canonicalReportName(item.reportName);
      // A child Subreport is evaluated only after its containing region emits a visible instance. SSRS
      // does not need a nested definition for a dormant call in an empty or hidden parent region, but a
      // call that does materialize must still fail closed when the caller omitted its bundle.
      item.resolveBundledSubreport = () => {
        const nested = resolveDefinition(childCanonical, [...lineage, canonical]);
        item.resolvedSubreport = nested;
        return nested;
      };
    });
    stripResolvedSubreportBlockers(child);

    const signatures = new Set();
    definition.instances = definition.instances.map((instance, index) => {
      if (!isPlainObject(instance)) {
        throw new ServiceError('RDL_INVALID', `Subreport ${definition.reportName} instance ${index} must be an object`);
      }
      const childRequest = {
        parameters: instance.parameters || {},
        datasets: instance.datasets || {},
      };
      // SSRS does not evaluate a subreport's parameters or datasets when the invoking item is hidden.
      // Defer full input validation until materialization, after the parent item's Visibility expression
      // has been evaluated. This also lets callers carry a harmless placeholder instance for a hidden
      // invocation. Keep counting declared rendering rows here so a hidden bundle cannot bypass limits.
      bundledRows += bundledRenderingRowCount(child, childRequest.datasets);
      if (bundledRows > config.maxRows) {
        throw new ServiceError('RDL_INVALID', `Dataset rows exceed the ${config.maxRows} row limit`, 413);
      }
      let signature;
      try {
        // Defaults are part of an invocation identity. Invalid parameters remain representable here so
        // hidden calls stay inert; a visible call is rejected by validateRenderInput during materialization.
        signature = parameterSignature(child.parameters, resolveParameterValues(child.parameters, childRequest.parameters));
      } catch (error) {
        if (!(error instanceof ServiceError) || error.code !== 'PARAMETER_INVALID') throw error;
        signature = parameterSignature(child.parameters, childRequest.parameters);
      }
      if (signatures.has(signature)) {
        throw new ServiceError(
          'RDL_INVALID',
          `Subreport ${definition.reportName} has duplicate invocation parameters`,
        );
      }
      signatures.add(signature);
      return { ...childRequest, signature };
    });
    definition.instancesBySignature = new Map(definition.instances.map((instance) => [instance.signature, instance]));
    return definition;
  };

  visitItems([
    ...model.body.items,
    ...(model.page.header?.items || []),
    ...(model.page.footer?.items || []),
  ], (item) => {
    if (item.type !== 'Subreport') return;
    const canonical = canonicalReportName(item.reportName);
    item.resolvedSubreport = resolveDefinition(canonical, []);
  });
  stripResolvedSubreportBlockers(model);

  const unused = [...bundles.values()].filter((definition) => !definition.used);
  if (unused.length > 0) {
    throw new ServiceError(
      'RDL_INVALID',
      `Unused bundled subreport definition: ${unused[0].reportName}`,
    );
  }
  return { model, bundledRows, fonts: [...fonts] };
}
