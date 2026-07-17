import {
  AlignmentType, BorderStyle, Document, Footer, Header, HorizontalPositionRelativeFrom, ImageRun, LineRuleType,
  PageOrientation, Packer, Paragraph, ShadingType, Table, TableCell, TableLayoutType, TableRow, TextRun,
  TextWrappingType, VerticalAlignTable, VerticalPositionRelativeFrom, WidthType,
} from 'docx';
import { pointsToDisplayPixels, pointsToTwips } from '../units.js';
import { cellText, cellTextbox, color, isHidden, normalizeDatasets, styleColor, styleValue, tablixRows, textForItem } from './common.js';
import { cellGridWidth, computeDocxTableGeometry } from './docxTableLayout.js';
import { computeCellPlacements } from './tableGrid.js';
import { materializeChart } from './chartData.js';
import { renderChartPng } from './chartImage.js';

function alignment(value) {
  const normalized = String(value || '').toLowerCase();
  if (normalized === 'center') return AlignmentType.CENTER;
  if (normalized === 'right') return AlignmentType.RIGHT;
  if (normalized === 'justify') return AlignmentType.JUSTIFIED;
  return AlignmentType.LEFT;
}

function textRuns(text, runProps) {
  // Word ignores literal "\n" inside a run, so each source line becomes its own run and every
  // line after the first is preceded by a break. This preserves the multi-paragraph cell text
  // that previously collapsed onto a single line in editable DOCX.
  const lines = String(text ?? '').split('\n');
  return lines.map((line, index) => new TextRun(
    index === 0 ? { ...runProps, text: line } : { ...runProps, text: line, break: 1 },
  ));
}

function paragraphForTextbox(item, context, overrideText, options = {}) {
  const style = item.style;
  const backgroundColor = styleColor(style.backgroundColor, context, null);
  const runProps = {
    font: styleValue(style.fontFamily, context, 'Arial'),
    size: Math.round((style.fontSize || 10) * 2),
    bold: /bold|600|700|800|900/i.test(String(styleValue(style.fontWeight, context, 'Normal'))),
    italics: /italic/i.test(String(styleValue(style.fontStyle, context, 'Normal'))),
    underline: /underline/i.test(String(styleValue(style.textDecoration, context, 'None'))) ? {} : undefined,
    strike: /line.?through/i.test(String(styleValue(style.textDecoration, context, 'None'))),
    color: styleColor(style.color, context, '#000000').replace('#', ''),
  };
  return new Paragraph({
    alignment: alignment(style.textAlign),
    shading: backgroundColor ? { type: ShadingType.CLEAR, fill: backgroundColor.replace('#', '') } : undefined,
    spacing: { before: 0, after: 0 },
    keepLines: Boolean(item.keepTogether),
    keepNext: Boolean(options.keepNext),
    widowControl: false,
    wordWrap: true,
    overflowPunctuation: false,
    // A group page break is expressed by breaking before the first paragraph of the group's first row.
    pageBreakBefore: options.pageBreakBefore || undefined,
    // Recursive (parent/child) groups indent the first cell by recursion depth.
    indent: options.indentLeft ? { left: options.indentLeft } : undefined,
    children: textRuns(overrideText ?? textForItem(item, context), runProps),
  });
}

function verticalAlignment(value) {
  const normalized = String(value || '').toLowerCase();
  if (normalized === 'middle' || normalized === 'center') return VerticalAlignTable.CENTER;
  if (normalized === 'bottom') return VerticalAlignTable.BOTTOM;
  return VerticalAlignTable.TOP;
}

function borderFor(style, context) {
  const sides = style?.borders || (style?.border ? { top: style.border, right: style.border, bottom: style.border, left: style.border } : null);
  if (!sides) return undefined;
  const convert = (border) => {
    const configuredStyle = String(styleValue(border?.style, context, 'None'));
    const configuredColor = styleColor(border?.color, context, null);
    if (/^none$/i.test(configuredStyle) || !configuredColor) return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const docxStyle = /dash/i.test(configuredStyle) ? BorderStyle.DASHED : /dot/i.test(configuredStyle) ? BorderStyle.DOTTED : BorderStyle.SINGLE;
    return { style: docxStyle, size: Math.max(1, Math.round((border.width || 1) * 8)), color: configuredColor.replace('#', '') };
  };
  return { top: convert(sides.top), right: convert(sides.right), bottom: convert(sides.bottom), left: convert(sides.left) };
}

// Native pixel dimensions of a PNG (IHDR) or JPEG (SOFn) buffer, or null when unreadable.
function naturalImageSize(buffer) {
  if (buffer.length > 24 && buffer.toString('ascii', 1, 4) === 'PNG') {
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
  }
  if (buffer.length > 4 && buffer[0] === 0xFF && buffer[1] === 0xD8) {
    let i = 2;
    while (i < buffer.length - 8) {
      if (buffer[i] !== 0xFF) { i += 1; continue; }
      const marker = buffer[i + 1];
      if (marker >= 0xC0 && marker <= 0xCF && ![0xC4, 0xC8, 0xCC].includes(marker)) {
        return { height: buffer.readUInt16BE(i + 5), width: buffer.readUInt16BE(i + 7) };
      }
      i += 2 + buffer.readUInt16BE(i + 2);
    }
  }
  return null;
}

function imageParagraph(model, item, floating = false) {
  const image = model.embeddedImages[item.value];
  if (!image?.data) return null;
  const buffer = Buffer.from(image.data.replace(/\s+/g, ''), 'base64');
  // Match the PDF renderer's Sizing handling. Fit (SSRS behaviour) stretches to fill the box; the box,
  // not the source aspect ratio, wins. FitProportional/Clip/AutoSize keep the source aspect ratio,
  // contained within the box, so the image is never distorted.
  const sizing = String(item.sizing || 'FitProportional');
  let width = item.width;
  let height = item.height;
  if (!/^Fit$/i.test(sizing)) {
    const natural = naturalImageSize(buffer);
    if (natural) {
      const scale = Math.min(item.width / natural.width, item.height / natural.height);
      width = natural.width * scale;
      height = natural.height * scale;
    }
  }
  // Header/footer images (e.g. the top-right logo) are positioned absolutely at their RDL Left/Top, like
  // the PDF, instead of flowing inline — otherwise the logo lands wherever the paragraph flow puts it.
  const emu = (pt) => Math.round(pt * 12700);
  const floatingOpts = floating ? {
    floating: {
      horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: emu(model.page.marginLeft + item.left) },
      verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: emu(model.page.marginTop + item.top) },
      wrap: { type: TextWrappingType.NONE },
      allowOverlap: true,
    },
  } : {};
  return new Paragraph({ children: [new ImageRun({
    data: buffer,
    type: image.mimeType.includes('jpeg') ? 'jpg' : 'png',
    transformation: { width: pointsToDisplayPixels(width), height: pointsToDisplayPixels(height) },
    ...floatingOpts,
  })] });
}

function childrenForItems(model, items, context, floating = false) {
  const children = [];
  for (const item of [...items].sort((left, right) => left.zIndex - right.zIndex || left.top - right.top || left.left - right.left)) {
    if (isHidden(item.hidden, context)) continue;
    if (item.type === 'Textbox') children.push(paragraphForTextbox(item, context));
    else if (item.type === 'Image') {
      const paragraph = imageParagraph(model, item, floating);
      if (paragraph) children.push(paragraph);
    } else if (item.type === 'Rectangle') children.push(...childrenForItems(model, item.items || [], context, floating));
    else if (item.type === 'Line') children.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: color(item.style.border.color).replace('#', '') } } }));
  }
  return children;
}

function tableForTablix(model, item, request) {
  const globals = { PageNumber: 1, TotalPages: 1, ReportName: request.outputFileName || model.name, ExecutionTime: new Date(), variables: model.variables || {} };
  const datasets = normalizeDatasets(model, request);
  const { rows, columns } = tablixRows(item, request, globals, model);
  // Matrix tablixes expand to a data-dependent column grid; clamp uses the expanded natural width.
  const layoutItem = item.hasColumnGroups ? { ...item, columns, width: columns.reduce((sum, width) => sum + width, 0) } : item;
  const geometry = computeDocxTableGeometry(model, layoutItem);
  const placements = computeCellPlacements(rows, geometry.gridTwips.length);
  // Keep-together: Word has no group-level "keep rows together", but a merged (row-span) cell means those
  // rows form one block. Mark every row a merge spans (except its last) to keepNext, so Word holds the
  // block on one page — mirroring the PDF, where the row-spanning Risk-No/rating cells keep each risk whole.
  const keepWithNext = new Array(rows.length).fill(false);
  rows.forEach((row, rowIndex) => {
    for (const cell of row.cells) {
      const span = cell.rowSpan || 1;
      if (span > 1) for (let r = rowIndex; r < Math.min(rowIndex + span - 1, rows.length - 1); r += 1) keepWithNext[r] = true;
    }
  });
  const tableContext = { fields: {}, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets };
  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: geometry.tableTwips, type: WidthType.DXA },
    indent: geometry.indentTwips > 0 ? { size: geometry.indentTwips, type: WidthType.DXA } : undefined,
    columnWidths: geometry.gridTwips,
    borders: borderFor(item.style, tableContext),
    rows: rows.map((row, rowIndex) => new TableRow({
      tableHeader: row.isHeader || undefined,
      cantSplit: row.isHeader || row.keepTogether || undefined,
      children: row.cells.map((cell, index, cells) => {
        const textbox = cellTextbox(cell);
        const style = textbox?.style || item.style;
        // Container-only cells (content wrapped in a Rectangle) have no border-bearing item of their own,
        // so their edges come from the tablix style — matching the PDF renderer.
        const borderStyle = cell.containerWrapped ? item.style : style;
        const columnIndex = placements[rowIndex][index];
        const width = cellGridWidth(geometry.gridTwips, columnIndex, cell.colSpan || 1);
        const context = { fields: row.fields, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets };
        const backgroundColor = styleColor(style.backgroundColor, context, null);
        return new TableCell({
          columnSpan: cell.colSpan || 1,
          rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
          width: { size: width, type: WidthType.DXA },
          verticalAlign: verticalAlignment(style.verticalAlign),
          borders: borderFor(borderStyle, context),
          shading: backgroundColor ? { type: ShadingType.CLEAR, fill: backgroundColor.replace('#', '') } : undefined,
          margins: { top: pointsToTwips(style.paddingTop), right: pointsToTwips(style.paddingRight), bottom: pointsToTwips(style.paddingBottom), left: pointsToTwips(style.paddingLeft) },
          children: textbox && !cell.hidden ? [paragraphForTextbox(textbox, context, cellText(cell), { pageBreakBefore: index === 0 && row.pageBreakBefore, keepNext: keepWithNext[rowIndex], indentLeft: index === 0 && row.indentLevel ? pointsToTwips(row.indentLevel * 12) : undefined })] : [new Paragraph('')],
        });
      }),
    })),
  });
}

function headerFooter(model, part, request, Kind) {
  if (!part) return undefined;
  const context = { parameters: request.parameters || {}, globals: { PageNumber: 1, TotalPages: 1, ExecutionTime: new Date(), variables: model.variables || {} }, datasets: normalizeDatasets(model, request), dataset: [] };
  return { default: new Kind({ children: childrenForItems(model, part.items, context, true) }) };
}

async function chartParagraph(model, item, request, config, tempDir, context, index) {
  // Charts are inherently graphical, so they embed as a rasterized image at chart size while the rest
  // of the document stays native OpenXML. Requires Poppler + a tempDir; without them the chart is
  // skipped rather than blocking the whole document.
  if (!config || !tempDir) return null;
  const data = materializeChart(item, context.datasets, context.parameters, context.globals);
  const image = await renderChartPng(item, data, config, tempDir, context, index);
  if (!image) return null;
  // Scale the chart image (preserving aspect ratio) so it can never exceed the usable page body and
  // spill onto an extra page. The reserve leaves room for an accompanying title/legend on the page.
  const usableWidth = Math.max(1, model.page.width - model.page.marginLeft - model.page.marginRight);
  const usableHeight = Math.max(1, model.page.height - model.page.marginTop - (model.page.header?.height || 0) - model.page.marginBottom - (model.page.footer?.height || 0));
  const reserve = 0.94;
  const fit = Math.min(1, usableWidth / item.width, (usableHeight * reserve) / item.height);
  return new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [new ImageRun({
      data: image.data,
      type: 'png',
      transformation: { width: pointsToDisplayPixels(item.width * fit), height: pointsToDisplayPixels(item.height * fit) },
    })],
  });
}

export async function renderEditableDocx(model, request, config, tempDir) {
  const landscape = model.page.width > model.page.height;
  const children = [];
  const context = { parameters: request.parameters || {}, globals: { PageNumber: 1, TotalPages: 1, ExecutionTime: new Date(), variables: model.variables || {} }, fields: {}, dataset: [], datasets: normalizeDatasets(model, request) };
  let forcePageBreak = false;
  let chartIndex = 0;
  for (const item of [...model.body.items].sort((a, b) => a.top - b.top || a.left - b.left || a.zIndex - b.zIndex)) {
    if (isHidden(item.hidden, context)) continue;
    const pageBreak = item.pageBreak && !isHidden(item.pageBreak.disabled, context) ? String(item.pageBreak.location) : 'None';
    // Start the next content on a fresh page with a minimal-height pageBreakBefore paragraph rather
    // than a standalone page-break run. A break run placed after page-filling content is pushed onto
    // the next page and then advances again, leaving a blank page; pageBreakBefore does not.
    if ((forcePageBreak || /^(Start|StartAndEnd)$/i.test(pageBreak)) && children.length > 0) {
      children.push(new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT } }));
    }
    forcePageBreak = false;
    if (item.type === 'Tablix') children.push(tableForTablix(model, item, request));
    else if (item.type === 'Chart') {
      const paragraph = await chartParagraph(model, item, request, config, tempDir, context, chartIndex);
      chartIndex += 1;
      if (paragraph) children.push(paragraph);
    } else children.push(...childrenForItems(model, [item], context));
    if (/^(End|StartAndEnd)$/i.test(pageBreak)) forcePageBreak = true;
  }
  const document = new Document({
    creator: 'RDL Converter Service',
    title: request.outputFileName || model.name,
    sections: [{
      properties: {
        page: {
          // A landscape-sized page must be declared landscape or Word reflows it and can spill pages.
          // The docx library swaps width/height for landscape, so pass the pre-rotation (portrait)
          // dimensions and let it produce the standard w>h landscape pgSz.
          size: landscape
            ? { width: pointsToTwips(model.page.height), height: pointsToTwips(model.page.width), orientation: PageOrientation.LANDSCAPE }
            : { width: pointsToTwips(model.page.width), height: pointsToTwips(model.page.height), orientation: PageOrientation.PORTRAIT },
          margin: {
            top: pointsToTwips(model.page.marginTop + (model.page.header?.height || 0)), right: pointsToTwips(model.page.marginRight),
            bottom: pointsToTwips(model.page.marginBottom + (model.page.footer?.height || 0)), left: pointsToTwips(model.page.marginLeft),
            header: pointsToTwips(model.page.marginTop), footer: pointsToTwips(model.page.marginBottom),
          },
        },
      },
      headers: headerFooter(model, model.page.header, request, Header),
      footers: headerFooter(model, model.page.footer, request, Footer),
      children,
    }],
  });
  const buffer = await Packer.toBuffer(document);
  return { buffer, pageCount: null, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', extension: 'docx' };
}
