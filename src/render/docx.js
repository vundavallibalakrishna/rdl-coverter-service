import {
  AlignmentType, BorderStyle, Document, Footer, Header, HeightRule, HorizontalPositionRelativeFrom, ImageRun, LineRuleType,
  PageOrientation, Packer, Paragraph, ShadingType, SimpleField, Table, TableCell, TableLayoutType, TableRow, TextRun,
  TextWrappingType, VerticalAlignTable, VerticalAnchor, VerticalPositionRelativeFrom, WidthType, WpsShapeRun,
} from 'docx';
import PDFDocument from 'pdfkit';
import { ServiceError } from '../errors.js';
import { pointsToDisplayPixels, pointsToTwips } from '../units.js';
import { evaluateExpression } from '../rdl/expression.js';
import { cellText, cellTextbox, color, isHidden, normalizeDatasets, styleColor, styleSize, styleValue, tablixRows, textForItem } from './common.js';
import { cellGridWidth, computeDocxTableGeometry } from './docxTableLayout.js';
import { computeCellPlacements } from './tableGrid.js';
import { materializeChart } from './chartData.js';
import { renderChartPng } from './chartImage.js';
import { pdfFont } from './fonts.js';
import { resolveStructuredDocxOptions } from './structuredCompatibility.js';

function alignment(value, context = {}) {
  // TextAlign can be an expression; resolve before matching so an `=IIF(...)` is not compared literally.
  const normalized = String(styleValue(value, context, '') || '').toLowerCase();
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

function splitConcatenation(expression) {
  const source = String(expression || '').replace(/^=/, '');
  const parts = [];
  let start = 0;
  let quote = null;
  let depth = 0;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (quote) {
      if (character === quote && source[index + 1] === quote) index += 1;
      else if (character === quote) quote = null;
      continue;
    }
    if (character === '"' || character === "'") quote = character;
    else if (character === '(') depth += 1;
    else if (character === ')') depth = Math.max(0, depth - 1);
    else if (character === '&' && depth === 0) {
      parts.push(source.slice(start, index).trim());
      start = index + 1;
    }
  }
  if (quote || depth !== 0) return null;
  parts.push(source.slice(start).trim());
  return parts.filter(Boolean);
}

// True when a textbox references PageNumber/TotalPages, which must become live Word fields rather than a
// value frozen at render time. Used both for band textboxes and for tablix cells (a footer table cell can
// hold "=Page N of M" and previously printed a literal "1" because the cell pre-flattened its text).
function hasPageFieldExpression(item) {
  return (item.paragraphs || []).some((paragraph) => paragraph.some((run) => (
    typeof (run?.value ?? run) === 'string' && /Globals!(PageNumber|TotalPages)/i.test(run?.value ?? run)
  )));
}

function fieldAwareChildren(item, context, runProps, overrideText) {
  if (overrideText !== undefined) return textRuns(overrideText, runProps);
  if (!hasPageFieldExpression(item)) return textRuns(textForItem(item, context), runProps);
  const values = (item.paragraphs || []).flatMap((paragraph, paragraphIndex) => paragraph.flatMap((run, runIndex) => {
    const source = run?.value ?? run;
    const children = [];
    if (paragraphIndex > 0 && runIndex === 0) children.push(new TextRun({ ...runProps, text: '', break: 1 }));
    if (typeof source !== 'string' || !/^=/.test(source) || !/Globals!(PageNumber|TotalPages)/i.test(source)) {
      const value = typeof source === 'string' && source.startsWith('=') ? evaluateExpression(source, context) : source;
      if (value !== null && value !== undefined) children.push(...textRuns(String(value), runProps));
      return children;
    }
    const parts = splitConcatenation(source);
    if (!parts) return textRuns(textForItem(item, context), runProps);
    for (const part of parts) {
      if (/^Globals!PageNumber(?:\.Value)?$/i.test(part)) children.push(new SimpleField('PAGE', '1'));
      else if (/^Globals!TotalPages(?:\.Value)?$/i.test(part)) children.push(new SimpleField('NUMPAGES', '1'));
      else {
        const value = evaluateExpression(`=${part}`, context);
        if (value !== null && value !== undefined) children.push(...textRuns(String(value), runProps));
      }
    }
    return children;
  }));
  return values.length ? values : textRuns(textForItem(item, context), runProps);
}

function paragraphForTextbox(item, context, overrideText, options = {}) {
  const style = item.style;
  const backgroundColor = styleColor(style.backgroundColor, context, null);
  let effectiveText = overrideText;
  if (effectiveText === undefined && options.clipToBox) {
    const lineHeight = (styleSize(style.fontSize, context, 10) || 10) * 1.2;
    const availableHeight = Math.max(1, item.height - styleSize(style.paddingTop, context, 2) - styleSize(style.paddingBottom, context, 2));
    const maxLines = Math.max(1, Math.floor(availableHeight / Math.max(1, lineHeight)));
    const lines = String(textForItem(item, context) ?? '').split('\n');
    if (lines.length > maxLines) effectiveText = lines.slice(0, maxLines).join('\n');
  }
  const runProps = {
    font: styleValue(style.fontFamily, context, 'Arial'),
    // fontSize may be an expression (e.g. =IIF(Sev="High",14,10)); resolve via styleSize (a literal passes
    // through). Using it raw produced NaN half-points and threw inside the docx library.
    size: Math.max(2, Math.round((styleSize(style.fontSize, context, 10) || 10) * 2)),
    bold: /bold|600|700|800|900/i.test(String(styleValue(style.fontWeight, context, 'Normal'))),
    italics: /italic/i.test(String(styleValue(style.fontStyle, context, 'Normal'))),
    underline: /underline/i.test(String(styleValue(style.textDecoration, context, 'None'))) ? {} : undefined,
    strike: /line.?through/i.test(String(styleValue(style.textDecoration, context, 'None'))),
    color: styleColor(style.color, context, '#000000').replace('#', ''),
  };
  return new Paragraph({
    alignment: alignment(style.textAlign, context),
    shading: !options.suppressShading && backgroundColor ? { type: ShadingType.CLEAR, fill: backgroundColor.replace('#', '') } : undefined,
    spacing: {
      before: 0,
      after: 0,
      line: options.lineHeight ? pointsToTwips(options.lineHeight) : undefined,
      lineRule: options.lineHeight ? LineRuleType.EXACT : undefined,
    },
    keepLines: Boolean(item.keepTogether),
    keepNext: Boolean(options.keepNext),
    widowControl: false,
    wordWrap: options.noWrap ? true : undefined,
    overflowPunctuation: false,
    // A group page break is expressed by breaking before the first paragraph of the group's first row.
    pageBreakBefore: options.pageBreakBefore || undefined,
    // Recursive (parent/child) groups indent the first cell by recursion depth.
    indent: options.indentLeft ? { left: options.indentLeft } : undefined,
    children: fieldAwareChildren(item, context, runProps, effectiveText),
  });
}

function verticalAlignment(value, context = {}) {
  // VerticalAlign can be an expression; resolve before matching.
  const normalized = String(styleValue(value, context, '') || '').toLowerCase();
  if (normalized === 'middle' || normalized === 'center') return VerticalAlignTable.CENTER;
  if (normalized === 'bottom') return VerticalAlignTable.BOTTOM;
  return VerticalAlignTable.TOP;
}

function borderFor(style, context) {
  const sides = style?.borders || (style?.border ? { top: style.border, right: style.border, bottom: style.border, left: style.border } : null);
  if (!sides) return undefined;
  const convert = (border) => {
    const configuredStyle = String(styleValue(border?.style, context, 'None'));
    if (!configuredStyle || /^none$/i.test(configuredStyle)) return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    // SSRS defaults an omitted BorderColor to Black, so a Solid border with no explicit colour must still
    // be drawn — resolving to null previously dropped the whole side. Width may be an expression too.
    const configuredColor = styleColor(border?.color, context, '#000000') || '#000000';
    const width = styleSize(border?.width, context, 1);
    // A conditional width of 0 (=IIF(rn=1,"1pt","0pt")) means the side is intentionally absent.
    if (width <= 0) return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const docxStyle = /dash/i.test(configuredStyle) ? BorderStyle.DASHED
      : /dot/i.test(configuredStyle) ? BorderStyle.DOTTED
      : /double/i.test(configuredStyle) ? BorderStyle.DOUBLE
      : BorderStyle.SINGLE;
    return { style: docxStyle, size: Math.max(1, Math.round(width * 8)), color: configuredColor.replace('#', '') };
  };
  return { top: convert(sides.top), right: convert(sides.right), bottom: convert(sides.bottom), left: convert(sides.left) };
}

// Word accepts png/jpg/gif/bmp raster data. Detect the real format from magic bytes rather than trusting
// the RDL-declared MIMEType (which can be missing, "image/jpg" without the "e", or simply wrong); a
// mislabelled type makes Word show a broken image. Falls back to png for anything unrecognized.
function detectImageType(buffer) {
  if (buffer.length > 8 && buffer[0] === 0x89 && buffer.toString('ascii', 1, 4) === 'PNG') return 'png';
  if (buffer.length > 3 && buffer[0] === 0xFF && buffer[1] === 0xD8) return 'jpg';
  if (buffer.length > 6 && buffer.toString('ascii', 0, 3) === 'GIF') return 'gif';
  if (buffer.length > 2 && buffer[0] === 0x42 && buffer[1] === 0x4D) return 'bmp';
  return 'png';
}

// Native pixel dimensions of a PNG/JPEG/GIF/BMP buffer, or null when unreadable.
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
  if (buffer.length > 10 && buffer.toString('ascii', 0, 3) === 'GIF') {
    return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) };
  }
  if (buffer.length > 26 && buffer[0] === 0x42 && buffer[1] === 0x4D) {
    return { width: buffer.readInt32LE(18), height: Math.abs(buffer.readInt32LE(22)) };
  }
  return null;
}

function imageParagraph(model, item, floating = false, origin = null, context = {}, align = undefined) {
  // Image Value and Sizing can be expressions; resolve before use or the raw expression misses the map.
  const image = model.embeddedImages[styleValue(item.value, context, item.value)];
  if (!image?.data) return null;
  const buffer = Buffer.from(image.data.replace(/\s+/g, ''), 'base64');
  // Match the PDF renderer's Sizing handling. Fit (SSRS behaviour) stretches to fill the box; the box,
  // not the source aspect ratio, wins. FitProportional/Clip/AutoSize keep the source aspect ratio,
  // contained within the box, so the image is never distorted.
  const sizing = String(styleValue(item.sizing, context, 'FitProportional') || 'FitProportional');
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
      horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: emu((origin?.x ?? model.page.marginLeft) + item.left) },
      verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: emu((origin?.y ?? model.page.marginTop) + item.top) },
      wrap: { type: TextWrappingType.NONE },
      allowOverlap: true,
    },
  } : {};
  return new Paragraph({
    // Absolutely-positioned (floating) images carry their own coordinates; a flowed image is placed
    // horizontally by paragraph alignment derived from its box position.
    alignment: floating ? undefined : align,
    children: [new ImageRun({
      data: buffer,
      type: detectImageType(buffer),
      transformation: { width: pointsToDisplayPixels(width), height: pointsToDisplayPixels(height) },
      ...floatingOpts,
    })],
  });
}

function positionedShapeParagraph(item, context, origin, children = [], wordZIndex = null) {
  const pointToEmu = (point) => Math.round(point * 12700);
  const backgroundColor = styleColor(item.style?.backgroundColor, context, null);
  const border = item.style?.border;
  const borderStyle = String(styleValue(border?.style, context, 'None'));
  const borderColor = styleColor(border?.color, context, null);
  const resolvedVerticalAlign = String(styleValue(item.style?.verticalAlign, context, '') || '');
  const verticalAnchor = /bottom/i.test(resolvedVerticalAlign) ? VerticalAnchor.BOTTOM
    : /middle|center/i.test(resolvedVerticalAlign) ? VerticalAnchor.CENTER : VerticalAnchor.TOP;
  const hasBorder = !/^none$/i.test(borderStyle) && Boolean(borderColor);
  const makeShape = ({ fill, outline, shapeChildren, zOffset = 0 }) => new WpsShapeRun({
    type: 'wps',
    transformation: { width: Math.max(1, pointsToDisplayPixels(item.width)), height: Math.max(1, pointsToDisplayPixels(item.height)) },
    floating: {
      horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: pointToEmu(origin.x + item.left) },
      verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: pointToEmu(origin.y + item.top) },
      wrap: { type: TextWrappingType.NONE },
      allowOverlap: true,
      layoutInCell: false,
      zIndex: Math.max(1, (wordZIndex ?? (item.zIndex + 1)) + zOffset),
    },
    solidFill: fill ? { type: 'rgb', value: fill.replace('#', '') } : undefined,
    outline,
    bodyProperties: {
      margins: {
        top: pointToEmu(styleSize(item.style?.paddingTop, context, 0)), right: pointToEmu(styleSize(item.style?.paddingRight, context, 0)),
        bottom: pointToEmu(styleSize(item.style?.paddingBottom, context, 0)), left: pointToEmu(styleSize(item.style?.paddingLeft, context, 0)),
      },
      verticalAnchor,
      // Page-section items are positioned in a fixed RDL band. Allowing Word to auto-grow them changes
      // z-order coverage and can hide adjacent repeated-header labels; keep the declared box exactly.
      noAutoFit: true,
    },
    children: shapeChildren?.length ? shapeChildren : [new Paragraph({ spacing: { before: 0, after: 0 }, children: [] })],
  });
  const borderOutline = hasBorder
    ? { type: 'solidFill', solidFillType: 'rgb', value: borderColor.replace('#', ''), width: pointToEmu(styleSize(border?.width, context, 1) || 1) }
    : backgroundColor ? undefined : { type: 'noFill' };
  // `docx` emits a shape that has both fill and outline in an invalid DrawingML property order
  // (`noFill`, line, then `solidFill`). Word keeps the line but ignores the late fill. Keep the visual
  // pieces editable, but stack them explicitly so text is never drawn beneath a fill or border shape.
  const runs = backgroundColor && hasBorder
    ? [
      makeShape({ fill: backgroundColor, outline: undefined, shapeChildren: [], zOffset: 0 }),
      makeShape({ fill: null, outline: { type: 'noFill' }, shapeChildren: children, zOffset: 1 }),
      makeShape({ fill: null, outline: borderOutline, shapeChildren: [], zOffset: 2 }),
    ]
    : [makeShape({ fill: backgroundColor, outline: borderOutline, shapeChildren: children })];
  return new Paragraph({
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
    children: runs,
  });
}

// An empty, exact-height paragraph that reproduces vertical whitespace from RDL Top coordinates when flowing
// free-form (coordinate-positioned) content, so a coordinate-designed page is not crammed against the top.
// Stays editable — it is just a blank line the user can adjust.
function spacerParagraph(points) {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: Math.max(1, pointsToTwips(points)), lineRule: LineRuleType.EXACT },
    children: [],
  });
}

// Horizontal placement for a flowed free-form item, derived from where its box sits in the container: a box
// with roughly equal left/right margins is centred, one flush to the right edge is right-aligned, else left.
// Generic — driven by the item's Left/Width, never by which item it is. Textboxes keep their own TextAlign;
// this is for items (images) that have no text-alignment of their own.
function flowAlignment(item, containerWidth) {
  const width = item.width || 0;
  if (!containerWidth || !width) return undefined;
  const left = item.left || 0;
  const right = containerWidth - (left + width);
  const tolerance = Math.max(6, containerWidth * 0.03);
  if (Math.abs(left - right) <= tolerance) return AlignmentType.CENTER;
  if (right <= tolerance && left > tolerance) return AlignmentType.RIGHT;
  return undefined;
}

// A flowed free-form Textbox whose fill/border should NOT span the whole text column — a narrower-than-
// container shaded or bordered box (e.g. a small date chip on a cover) — becomes a fixed-width single-cell
// table so the shading/border is confined to the RDL Width and the box is horizontally placed from its
// Left. Returns null when the box has no visible fill/border or is effectively full width, so a plain
// paragraph (or a full-width shaded bar) keeps the simpler path unchanged. Generic: width, position, fill,
// and border all come from the model, nothing keyed to a specific report. A minimal separator paragraph
// trails the table so an adjacent flowed box is not merged into it by Word.
function shadedBoxTable(child, context, containerWidth) {
  const width = child.width || 0;
  if (!width || !containerWidth) return null;
  const tolerance = Math.max(6, containerWidth * 0.03);
  if (width >= containerWidth - tolerance) return null; // full-width bar: paragraph shading is fine
  const backgroundColor = styleColor(child.style?.backgroundColor, context, null);
  const resolved = borderFor(child.style, context);
  const hasBorder = resolved && Object.values(resolved).some((side) => side.style !== BorderStyle.NONE);
  if (!backgroundColor && !hasBorder) return null; // nothing to confine; alignment/text handles placement
  // Explicit no-border set: borderFor returns undefined when the RDL declares no border, and the docx
  // library then draws its OWN default single-line grid — a border the report never asked for.
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const borders = {
    top: resolved?.top || none, bottom: resolved?.bottom || none, left: resolved?.left || none,
    right: resolved?.right || none, insideHorizontal: none, insideVertical: none,
  };
  const align = flowAlignment(child, containerWidth);
  const indentTwips = align ? 0 : pointsToTwips(child.left || 0); // left-aligned box sits at its Left offset
  const table = new Table({
    layout: TableLayoutType.FIXED,
    width: { size: pointsToTwips(width), type: WidthType.DXA },
    columnWidths: [pointsToTwips(width)],
    alignment: align === AlignmentType.CENTER ? AlignmentType.CENTER : align === AlignmentType.RIGHT ? AlignmentType.RIGHT : undefined,
    indent: indentTwips > 0 ? { size: indentTwips, type: WidthType.DXA } : undefined,
    borders,
    rows: [new TableRow({
      height: { value: Math.max(1, pointsToTwips(child.height || 0)), rule: HeightRule.ATLEAST },
      children: [new TableCell({
        width: { size: pointsToTwips(width), type: WidthType.DXA },
        verticalAlign: verticalAlignment(child.style?.verticalAlign, context),
        shading: backgroundColor ? { type: ShadingType.CLEAR, fill: backgroundColor.replace('#', '') } : undefined,
        margins: {
          top: pointsToTwips(styleSize(child.style?.paddingTop, context, 2)),
          right: pointsToTwips(styleSize(child.style?.paddingRight, context, 2)),
          bottom: pointsToTwips(styleSize(child.style?.paddingBottom, context, 2)),
          left: pointsToTwips(styleSize(child.style?.paddingLeft, context, 2)),
        },
        children: [paragraphForTextbox(child, context, undefined, { suppressShading: true })],
      })],
    })],
  });
  return [table, spacerParagraph(0)];
}

// Flows a free-form container's children top-to-bottom. Each child reserves its full RDL box HEIGHT (not
// just its text height) with its vertical alignment applied as top/bottom padding — otherwise Word collapses
// an oversized or vertically-centred textbox to its text and the whitespace below it disappears, making the
// gaps uneven. Images are horizontally aligned from their box position. Gaps between boxes and the remaining
// container height are emitted as editable spacer lines.
function flowChildrenWithSpacing(model, items, context, containerHeight, containerWidth, zOrder) {
  const out = [];
  let flowBottom = 0;
  for (const child of [...items].sort((a, b) => (a.top || 0) - (b.top || 0) || (a.left || 0) - (b.left || 0))) {
    if (isHidden(child.hidden, context)) continue;
    const gap = (child.top || 0) - flowBottom;
    if (gap > 0.5) out.push(spacerParagraph(gap));
    const boxHeight = child.height || 0;
    const boxed = child.type === 'Textbox' && boxHeight > 0 ? shadedBoxTable(child, context, containerWidth) : null;
    if (boxed) {
      // A narrower-than-container shaded/bordered box: fixed-width single-cell table. Its row height already
      // reserves the full box height, so no valign-padding spacers are needed.
      out.push(...boxed);
    } else if (child.type === 'Textbox' && boxHeight > 0) {
      // Estimate the rendered text height, then split the leftover box height per VerticalAlign so the box
      // keeps the internal padding it has in the PDF.
      const fontSize = styleSize(child.style?.fontSize, context, 10) || 10;
      const lineCount = Math.max(1, String(textForItem(child, context) ?? '').split('\n').length);
      const textHeight = Math.min(boxHeight, lineCount * fontSize * 1.25);
      const slack = Math.max(0, boxHeight - textHeight);
      const verticalAlign = String(styleValue(child.style?.verticalAlign, context, 'top')).toLowerCase();
      const topPad = /middle|center/.test(verticalAlign) ? slack / 2 : /bottom/.test(verticalAlign) ? slack : 0;
      const bottomPad = /middle|center/.test(verticalAlign) ? slack / 2 : /bottom/.test(verticalAlign) ? 0 : slack;
      if (topPad > 0.5) out.push(spacerParagraph(topPad));
      out.push(...childrenForItems(model, [child], context, false, null, zOrder));
      if (bottomPad > 0.5) out.push(spacerParagraph(bottomPad));
    } else if (child.type === 'Image') {
      const paragraph = imageParagraph(model, child, false, null, context, flowAlignment(child, containerWidth));
      if (paragraph) out.push(paragraph);
    } else {
      out.push(...childrenForItems(model, [child], context, false, null, zOrder));
    }
    flowBottom = Math.max(flowBottom, (child.top || 0) + boxHeight);
  }
  const trailing = (containerHeight || 0) - flowBottom;
  if (trailing > 0.5) out.push(spacerParagraph(trailing));
  return out;
}

function childrenForItems(model, items, context, floating = false, origin = null, zOrder = { value: 1 }) {
  const children = [];
  const nextZIndex = () => {
    const value = zOrder.value;
    zOrder.value += 3;
    return value;
  };
  for (const item of [...items].sort((left, right) => left.zIndex - right.zIndex || left.top - right.top || left.left - right.left)) {
    if (isHidden(item.hidden, context)) continue;
    if (item.type === 'Textbox') {
      if (floating) children.push(positionedShapeParagraph(item, context, origin, [paragraphForTextbox(item, context, undefined, { suppressShading: true, noWrap: true, clipToBox: true })], nextZIndex()));
      else children.push(paragraphForTextbox(item, context));
    }
    else if (item.type === 'Image') {
      const paragraph = imageParagraph(model, item, floating, origin, context);
      if (paragraph) children.push(paragraph);
    } else if (item.type === 'Rectangle') {
      // Resolve BackgroundColor and Border/Style before gating — an expression string is always truthy and
      // always fails `^none$`, so a raw test would draw/omit the shape opposite to its evaluated result.
      const rectBg = styleColor(item.style?.backgroundColor, context, null);
      const rectBorderStyle = String(styleValue(item.style?.border?.style, context, 'None'));
      if (floating && (rectBg || !/^none$/i.test(rectBorderStyle))) children.push(positionedShapeParagraph(item, context, origin, [], nextZIndex()));
      if (floating) {
        const childOrigin = { x: origin.x + item.left, y: origin.y + item.top };
        children.push(...childrenForItems(model, item.items || [], context, true, childOrigin, zOrder));
      } else {
        // Body (flow) Rectangle: preserve its internal vertical layout with spacers, and place its children
        // horizontally from their box positions — so a coordinate-designed cover keeps its layout.
        children.push(...flowChildrenWithSpacing(model, item.items || [], context, item.height, item.width, zOrder));
      }
    } else if (item.type === 'Line') {
      const lineWidth = styleSize(item.style?.border?.width, context, 1) || 1;
      const lineColor = styleColor(item.style?.border?.color, context, '#000000');
      if (floating) {
        const line = {
          ...item,
          width: Math.max(item.width, lineWidth),
          height: Math.max(item.height, lineWidth),
          style: { ...item.style, backgroundColor: lineColor, border: { style: 'None' }, paddingTop: 0, paddingRight: 0, paddingBottom: 0, paddingLeft: 0 },
        };
        children.push(positionedShapeParagraph(line, context, origin, [], nextZIndex()));
      } else children.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: lineColor.replace('#', '') } } }));
    }
  }
  return children;
}

function applyMeasurementFont(doc, config, style, context) {
  const bold = /bold|600|700|800|900/i.test(String(styleValue(style.fontWeight, context, 'Normal')));
  const italic = /italic/i.test(String(styleValue(style.fontStyle, context, 'Normal')));
  const family = styleValue(style.fontFamily, context, 'Arial');
  doc.font(pdfFont(config, family, bold, italic)).fontSize(styleSize(style.fontSize, context, 10) || 10);
}

function measureRows(rows, geometry, request, globals, datasets, datasetName, config, measurementDoc) {
  const placements = computeCellPlacements(rows, geometry.gridTwips.length);
  const metrics = rows.map((row, rowIndex) => {
    const contexts = row.cells.map((cell) => ({
      fields: row.fields,
      parameters: request.parameters || {},
      globals,
      dataset: datasets[datasetName] || [],
      datasets,
      cell,
    }));
    const heights = row.cells.map((cell, index) => {
      const textbox = cellTextbox(cell);
      if (!textbox || cell.hidden) return 0;
      const columnIndex = placements[rowIndex][index];
      const width = cellGridWidth(geometry.gridTwips, columnIndex, cell.colSpan || 1) / 20;
      const context = contexts[index];
      applyMeasurementFont(measurementDoc, config, textbox.style, context);
      const innerWidth = Math.max(1, width - styleSize(textbox.style.paddingLeft, context, 2) - styleSize(textbox.style.paddingRight, context, 2));
      const text = cellText(cell);
      const textHeight = text ? measurementDoc.heightOfString(text, { width: innerWidth, lineGap: 0 }) : 0;
      return textHeight + styleSize(textbox.style.paddingTop, context, 2) + styleSize(textbox.style.paddingBottom, context, 2);
    });
    return {
      height: Math.max(row.height, ...heights),
      lineHeights: row.cells.map((cell, index) => {
        const textbox = cellTextbox(cell);
        if (!textbox) return null;
        applyMeasurementFont(measurementDoc, config, textbox.style, contexts[index]);
        return measurementDoc.currentLineHeight(true);
      }),
    };
  });
  return { placements, metrics };
}

function fragmentRowsForWord(model, rows, rowMetrics) {
  const headerRows = rows.filter((row) => row.isHeader);
  const bodyRows = rows.filter((row) => !row.isHeader);
  if (bodyRows.length === 0) return [rows];

  const metricByRow = new Map(rows.map((row, index) => [row, rowMetrics[index]]));
  const headerHeight = headerRows.reduce((sum, row) => sum + (metricByRow.get(row)?.height || row.height || 0), 0);
  const pageBodyHeight = Math.max(1,
    model.page.height - model.page.marginTop - (model.page.header?.height || 0)
      - model.page.marginBottom - (model.page.footer?.height || 0));
  const capacity = Math.max(1, pageBodyHeight - headerHeight);

  const fragments = [];
  let current = [];
  let currentHeight = 0;
  for (let index = 0; index < bodyRows.length;) {
    const row = bodyRows[index];
    let blockEnd = index;
    for (let cursor = index; cursor <= blockEnd && cursor < bodyRows.length; cursor += 1) {
      const sourceIndex = rows.indexOf(bodyRows[cursor]);
      for (const cell of bodyRows[cursor].cells) {
        const span = cell.rowSpan || 1;
        if (span > 1) {
          const endSourceIndex = sourceIndex + span - 1;
          const endBodyIndex = bodyRows.findIndex((candidate) => rows.indexOf(candidate) === endSourceIndex);
          if (endBodyIndex > blockEnd) blockEnd = endBodyIndex;
        }
      }
    }
    if (blockEnd < index) blockEnd = index;
    const block = bodyRows.slice(index, blockEnd + 1);
    const blockHeight = block.reduce((sum, candidate) => sum + (metricByRow.get(candidate)?.height || candidate.height || 0), 0);
    if ((row.pageBreakBefore || (current.length > 0 && currentHeight + blockHeight > capacity)) && current.length > 0) {
      fragments.push([...headerRows, ...current]);
      current = [];
      currentHeight = 0;
    }
    current.push(...block);
    currentHeight += blockHeight;
    index = blockEnd + 1;
  }
  if (current.length > 0) fragments.push([...headerRows, ...current]);
  return fragments.length > 0 ? fragments : [rows];
}

function tableForTablix(model, item, request, config, measurementDoc, structuredOptions) {
  const globals = { PageNumber: 1, TotalPages: 1, ReportName: request.outputFileName || model.name, ExecutionTime: new Date(), variables: model.variables || {} };
  const datasets = normalizeDatasets(model, request);
  const { rows, columns } = tablixRows(item, request, globals, model);
  // Matrix tablixes expand to a data-dependent column grid; clamp uses the expanded natural width.
  const layoutItem = item.hasColumnGroups ? { ...item, columns, width: columns.reduce((sum, width) => sum + width, 0) } : item;
  const geometry = computeDocxTableGeometry(model, layoutItem);
  const { placements, metrics: rowMetrics } = measureRows(rows, geometry, request, globals, datasets, item.datasetName, config, measurementDoc);
  // Keep-together: Word has no group-level "keep rows together", but a merged (row-span) cell means those
  // rows form one block. Only link a span when the whole block can fit on a fresh content page. The PDF
  // renderer makes the same distinction; chaining an oversized span makes Word move content repeatedly
  // and produces the large blank areas that SSRS avoids by splitting the group.
  const keepWithNext = new Array(rows.length).fill(false);
  const repeatedHeaderHeight = rows.reduce((sum, row, index) => sum + (row.isHeader ? rowMetrics[index].height : 0), 0);
  const freshPageCapacity = Math.max(1,
    model.page.height - model.page.marginTop - (model.page.header?.height || 0)
      - model.page.marginBottom - (model.page.footer?.height || 0) - repeatedHeaderHeight);
  rows.forEach((row, rowIndex) => {
    for (const cell of row.cells) {
      const span = cell.rowSpan || 1;
      const end = Math.min(rowIndex + span, rows.length);
      const spanHeight = rowMetrics.slice(rowIndex, end).reduce((sum, metric) => sum + metric.height, 0);
      if (span > 1 && spanHeight <= freshPageCapacity) {
        for (let r = rowIndex; r < Math.min(end - 1, rows.length - 1); r += 1) keepWithNext[r] = true;
      }
    }
  });
  const useNativePageFragments = structuredOptions?.nativePageFragments === true;
  const fragments = useNativePageFragments ? fragmentRowsForWord(model, rows, rowMetrics) : [rows];
  const metricByRow = new Map(rows.map((row, index) => [row, rowMetrics[index]]));
  const keepWithNextByRow = new Map(rows.map((row, index) => [row, keepWithNext[index]]));
  const tableContext = { fields: {}, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets };
  const tableForRows = (fragmentRows, fragmentIndex) => {
    const fragmentPlacements = computeCellPlacements(fragmentRows, geometry.gridTwips.length);
    return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: geometry.tableTwips, type: WidthType.DXA },
    indent: geometry.indentTwips > 0 ? { size: geometry.indentTwips, type: WidthType.DXA } : undefined,
    columnWidths: geometry.gridTwips,
    borders: borderFor(item.style, tableContext),
    rows: fragmentRows.map((row, rowIndex) => {
      const rowMetric = metricByRow.get(row);
      return new TableRow({
      tableHeader: row.isHeader || undefined,
      cantSplit: row.isHeader || row.keepTogether || undefined,
      height: { value: pointsToTwips(rowMetric?.height || row.height), rule: HeightRule.ATLEAST },
      children: row.cells.map((cell, index, cells) => {
        const textbox = cellTextbox(cell);
        const style = textbox?.style || item.style;
        // Container-only cells (content wrapped in a Rectangle) have no border-bearing item of their own,
        // so their edges come from the tablix style — matching the PDF renderer.
        const borderStyle = cell.containerWrapped ? item.style : style;
        const columnIndex = fragmentPlacements[rowIndex][index];
        const width = cellGridWidth(geometry.gridTwips, columnIndex, cell.colSpan || 1);
        const context = { fields: row.fields, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets };
        const backgroundColor = styleColor(style.backgroundColor, context, null);
        return new TableCell({
          columnSpan: cell.colSpan || 1,
          rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
          width: { size: width, type: WidthType.DXA },
          verticalAlign: verticalAlignment(style.verticalAlign, context),
          borders: borderFor(borderStyle, context),
          shading: backgroundColor ? { type: ShadingType.CLEAR, fill: backgroundColor.replace('#', '') } : undefined,
          margins: { top: pointsToTwips(styleSize(style.paddingTop, context, 2)), right: pointsToTwips(styleSize(style.paddingRight, context, 2)), bottom: pointsToTwips(styleSize(style.paddingBottom, context, 2)), left: pointsToTwips(styleSize(style.paddingLeft, context, 2)) },
          children: textbox && !cell.hidden ? [paragraphForTextbox(textbox, context,
            hasPageFieldExpression(textbox) ? undefined : cellText(cell), {
            pageBreakBefore: index === 0 && rowIndex === 0 && fragmentIndex > 0,
            keepNext: keepWithNextByRow.get(row),
            indentLeft: index === 0 && row.indentLevel ? pointsToTwips(row.indentLevel * 12) : undefined,
            lineHeight: rowMetric?.lineHeights[index],
          })] : [new Paragraph('')],
        });
      }),
    });
    }),
  });
  };
  return fragments.map((fragment, index) => tableForRows(fragment, index));
}

function headerFooter(model, part, request, Kind, location) {
  if (!part) return undefined;
  // A Word header/footer is materialized once and reused. Evaluating page-dependent RDL visibility as
  // page 1 of 1 incorrectly hides constructs intended for every non-final page (a common SSRS repeated-
  // header pattern). Use a representative middle page here; PAGE/NUMPAGES text itself is still emitted
  // as live Word fields by fieldAwareChildren.
  const context = { parameters: request.parameters || {}, globals: { PageNumber: 2, TotalPages: 3, ExecutionTime: new Date(), variables: model.variables || {} }, datasets: normalizeDatasets(model, request), dataset: [] };
  const origin = {
    x: model.page.marginLeft,
    y: location === 'footer' ? model.page.height - model.page.marginBottom - part.height : model.page.marginTop,
  };
  return { default: new Kind({ children: childrenForItems(model, part.items, context, true, origin) }) };
}

async function chartParagraph(model, item, request, config, tempDir, context, index) {
  // Charts are inherently graphical, so they embed as a rasterized image at chart size while the rest of
  // the document stays native OpenXML. This needs Poppler + a tempDir. A visible chart that cannot be
  // rendered must FAIL CLOSED — a silently missing chart in a compliance document is the exact silent-drop
  // class the fail-closed design exists to prevent. Hidden charts are already skipped before reaching here.
  if (!config || !tempDir) {
    throw new ServiceError('RENDER_FAILED', 'Chart rendering requires a configured render environment (config and temporary directory)', 500, { chart: item.name });
  }
  const data = materializeChart(item, context.datasets, context.parameters, context.globals);
  const image = await renderChartPng(item, data, config, tempDir, context, index);
  if (!image) {
    throw new ServiceError('RENDER_FAILED', `Chart '${item.name}' could not be rendered`, 500, { chart: item.name });
  }
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
  const measurementConfig = config || { strictFonts: false, fontDir: process.cwd() };
  const structuredOptions = resolveStructuredDocxOptions(model, request, measurementConfig);
  const measurementDoc = new PDFDocument({ autoFirstPage: false });
  measurementDoc.on('data', () => {});
  measurementDoc.addPage({ size: [model.page.width, model.page.height], margins: { top: 0, right: 0, bottom: 0, left: 0 } });
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
    if (item.type === 'Tablix') children.push(...tableForTablix(model, item, request, measurementConfig, measurementDoc, structuredOptions));
    else if (item.type === 'Chart') {
      children.push(await chartParagraph(model, item, request, config, tempDir, context, chartIndex));
      chartIndex += 1;
    } else children.push(...childrenForItems(model, [item], context));
    if (/^(End|StartAndEnd)$/i.test(pageBreak)) forcePageBreak = true;
  }
  measurementDoc.end();
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
      headers: headerFooter(model, model.page.header, request, Header, 'header'),
      footers: headerFooter(model, model.page.footer, request, Footer, 'footer'),
      children,
    }],
  });
  const buffer = await Packer.toBuffer(document);
  return {
    buffer,
    pageCount: null,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    extension: 'docx',
    layoutMode: 'structured',
    editableTextRatio: 1,
    docxProfile: structuredOptions.profile?.selected || null,
    docxNativePageFragments: structuredOptions.nativePageFragments === true,
  };
}
