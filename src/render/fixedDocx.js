import {
  Document, HorizontalPositionRelativeFrom, ImageRun, LineRuleType, PageBreak, PageOrientation, Packer,
  Paragraph, TextRun, TextWrappingType, VerticalAnchor, VerticalPositionRelativeFrom, WpsShapeRun,
} from 'docx';
import JSZip from 'jszip';
import { ServiceError } from '../errors.js';
import { pointsToDisplayPixels, pointsToTwips } from '../units.js';
import { renderPdf } from './pdf.js';
import { extractPdfScene } from './pdfScene.js';

const pointToEmu = (value) => Math.round(value * 12700);
const colorValue = (value, fallback = '000000') => String(value || fallback).replace(/^#/, '').toUpperCase();

function floating(left, top, z) {
  return {
    horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: pointToEmu(left) },
    verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: pointToEmu(top) },
    wrap: { type: TextWrappingType.NONE },
    allowOverlap: true,
    behindDocument: false,
    layoutInCell: false,
    zIndex: Math.max(1, z + 1),
  };
}

function shapeRun({ left, top, width, height, fill, stroke, strokeWidth = 0, z, children, rotation }) {
  const safeWidth = Math.max(width, 0.15);
  const safeHeight = Math.max(height, 0.15);
  return new WpsShapeRun({
    type: 'wps',
    transformation: {
      width: Math.max(1, pointsToDisplayPixels(safeWidth)),
      height: Math.max(1, pointsToDisplayPixels(safeHeight)),
      rotation: Number(rotation) || undefined,
    },
    floating: floating(left, top, z),
    solidFill: fill ? { type: 'rgb', value: colorValue(fill) } : undefined,
    // `docx` serializes a no-fill outline before the shape fill. Supplying both can produce an
    // invalid DrawingML property order that Word-compatible viewers interpret as a black fill.
    // Filled shapes therefore omit the outline entirely; fill-and-stroke objects are split below.
    outline: stroke && strokeWidth > 0
      ? { type: 'solidFill', solidFillType: 'rgb', value: colorValue(stroke), width: pointToEmu(strokeWidth) }
      : fill ? undefined : { type: 'noFill' },
    bodyProperties: {
      margins: { top: 0, right: 0, bottom: 0, left: 0 },
      verticalAnchor: VerticalAnchor.TOP,
      noAutoFit: true,
    },
    children: children || [new Paragraph({ spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT }, children: [] })],
  });
}

function textRun(object) {
  const paragraph = new Paragraph({
    spacing: { before: 0, after: 0, line: Math.max(1, pointsToTwips(object.fontSize)), lineRule: LineRuleType.EXACT },
    widowControl: false,
    children: [new TextRun({
      text: object.text,
      font: object.fontFamily,
      size: Math.max(2, Math.round(object.fontSize * 2)),
      bold: object.bold || undefined,
      italics: object.italic || undefined,
      color: colorValue(object.color),
    })],
  });
  // The PDF item width is its exact glyph advance, while Word applies its own font rounding and may wrap
  // the final word at that exact boundary. A bounded reserve prevents wrapping without creating hundreds
  // of page-wide overlapping text boxes (which some Word-compatible viewers render incorrectly).
  const noWrapWidth = object.width + Math.max(4, object.width * 0.2, object.fontSize * 0.5);
  return shapeRun({ ...object, width: noWrapWidth, children: [paragraph] });
}

function solidLineRuns(object) {
  const thickness = Math.max(object.strokeWidth || 0.5, 0.25);
  if (object.diagonal) {
    const length = Math.hypot(object.x2 - object.x1, object.y2 - object.y1);
    const centerX = (object.x1 + object.x2) / 2;
    const centerY = (object.y1 + object.y2) / 2;
    return [shapeRun({
      ...object,
      left: centerX - length / 2,
      top: centerY - thickness / 2,
      width: length,
      height: thickness,
      rotation: Math.atan2(object.y2 - object.y1, object.x2 - object.x1) * (180 / Math.PI),
      fill: object.fill,
      stroke: null,
    })];
  }
  const horizontal = object.width >= object.height;
  const length = horizontal ? object.width : object.height;
  const dash = Array.isArray(object.dash) && object.dash.length ? object.dash : null;
  if (!dash) {
    return [shapeRun({
      ...object,
      width: horizontal ? Math.max(length, 0.25) : thickness,
      height: horizontal ? thickness : Math.max(length, 0.25),
      fill: object.fill,
      stroke: null,
    })];
  }
  const runs = [];
  let cursor = 0;
  let index = 0;
  let paint = true;
  while (cursor < length) {
    const segment = Math.max(0.25, dash[index % dash.length]);
    if (paint) {
      const visible = Math.min(segment, length - cursor);
      runs.push(shapeRun({
        ...object,
        left: object.left + (horizontal ? cursor : 0),
        top: object.top + (horizontal ? 0 : cursor),
        width: horizontal ? visible : thickness,
        height: horizontal ? thickness : visible,
        fill: object.fill,
        stroke: null,
        z: object.z + runs.length,
      }));
    }
    cursor += segment;
    index += 1;
    paint = !paint;
  }
  return runs;
}

function objectRuns(object) {
  if (object.type === 'text') return [textRun(object)];
  if (object.type === 'image') {
    return [new ImageRun({
      data: object.data,
      type: 'png',
      transformation: { width: Math.max(1, pointsToDisplayPixels(object.width)), height: Math.max(1, pointsToDisplayPixels(object.height)) },
      floating: floating(object.left, object.top, object.z),
    })];
  }
  if (object.type === 'shape' && object.shape === 'line') return solidLineRuns(object);
  if (object.type === 'shape') {
    // Keep the OOXML shape properties in schema order. A fill and its border are visually
    // equivalent as two coincident shapes and avoid the serializer's fill/outline ordering bug.
    if (object.fill && object.stroke && object.strokeWidth > 0) {
      return [
        shapeRun({ ...object, stroke: null }),
        shapeRun({ ...object, fill: null, z: object.z + 1 }),
      ];
    }
    return [shapeRun(object)];
  }
  return [];
}

function pageSize(page) {
  const landscape = page.width > page.height;
  return landscape
    ? { width: pointsToTwips(page.height), height: pointsToTwips(page.width), orientation: PageOrientation.LANDSCAPE }
    : { width: pointsToTwips(page.width), height: pointsToTwips(page.height), orientation: PageOrientation.PORTRAIT };
}

async function finalizeOpenXml(buffer, expectObjects) {
  const zip = await JSZip.loadAsync(buffer);
  const documentPart = zip.file('word/document.xml');
  const settingsPart = zip.file('word/settings.xml');
  if (!documentPart) throw new ServiceError('RENDER_FAILED', 'Fixed editable Word package is missing document.xml');
  if (!settingsPart) throw new ServiceError('RENDER_FAILED', 'Fixed editable Word package is missing settings.xml');
  const xml = await documentPart.async('string');
  const settingsXml = await settingsPart.async('string');
  if (/<w:(?:documentProtection|writeProtection)\b/.test(settingsXml)) {
    throw new ServiceError('RENDER_FAILED', 'Fixed editable Word package must not be protected or read-only');
  }
  let drawingId = 0;
  const uniqueIds = xml.replace(/<wp:docPr\b([^>]*?)\bid="[^"]*"([^>]*)\/>/g, (_match, before, after) => {
    drawingId += 1;
    return `<wp:docPr${before}id="${drawingId}"${after}/>`;
  });
  // Zero ids is only a defect when the scene actually had objects to place — that means the serializer's
  // docPr shape changed and the id-uniquifier stopped matching. A genuinely empty scene (e.g. an all-blank
  // page) is a valid document, not a failure.
  if (expectObjects && drawingId === 0) throw new ServiceError('RENDER_FAILED', 'Fixed editable Word package contains no positioned objects');
  // LibreOffice intermittently paints WPS text boxes with `<a:noFill/>` black on high-object-count pages.
  // Express the same DrawingML semantics as an explicit fully transparent solid colour. Word and
  // LibreOffice both preserve it, while the line's own no-fill remains untouched.
  const transparentFill = '<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill>';
  const normalized = uniqueIds.replace(
    /(<wps:spPr\b(?:(?!<\/wps:spPr>)[\s\S])*?<\/a:prstGeom>)<a:noFill\/>/g,
    `$1${transparentFill}`,
  );
  if (/\blocked="1"|\bnoTextEdit\b/.test(normalized)) {
    throw new ServiceError('RENDER_FAILED', 'Fixed editable Word package contains a locked drawing object');
  }
  zip.file('word/document.xml', normalized);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export async function renderFixedEditableDocx(model, request, config) {
  const pdf = await renderPdf(model, request, config);
  const scene = await extractPdfScene(pdf.buffer, config, { defaultFontFamily: model.defaultFontFamily });
  if (scene.pageCount !== pdf.pageCount) throw new ServiceError('RENDER_FAILED', 'PDF scene extraction produced an unexpected page count');
  if (scene.unsupportedPdfOperators.length) {
    throw new ServiceError('UNSUPPORTED_FEATURE', 'Canonical PDF contains constructs that fixed editable Word cannot preserve', 400, {
      unsupportedPdfOperators: scene.unsupportedPdfOperators,
    });
  }
  const first = scene.pages[0];
  if (!scene.pages.every((page) => Math.abs(page.width - first.width) < 0.01 && Math.abs(page.height - first.height) < 0.01)) {
    throw new ServiceError('UNSUPPORTED_FEATURE', 'Mixed PDF page sizes are not supported by fixed editable Word', 400, { feature: 'MIXED_PAGE_SIZES' });
  }
  const children = [];
  let positionedObjectCount = 0;
  for (const [index, page] of scene.pages.entries()) {
    const runs = page.objects.flatMap((object) => objectRuns({ ...object, pageWidth: page.width }));
    positionedObjectCount += runs.length;
    if (positionedObjectCount > config.maxFixedObjects) {
      throw new ServiceError('UNSUPPORTED_FEATURE', 'Fixed editable Word positioned-object limit exceeded', 413, {
        limit: config.maxFixedObjects,
        actual: positionedObjectCount,
      });
    }
    children.push(new Paragraph({
      spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
      children: runs,
    }));
    if (index < scene.pages.length - 1) {
      children.push(new Paragraph({ spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT }, children: [new PageBreak()] }));
    }
  }
  const document = new Document({
    creator: 'RDL Converter Service',
    title: request.outputFileName || model.name,
    sections: [{
      properties: { page: { size: pageSize(first), margin: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0 } } },
      children,
    }],
  });
  const buffer = await finalizeOpenXml(await Packer.toBuffer(document), positionedObjectCount > 0);
  return {
    buffer,
    pageCount: scene.pageCount,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    extension: 'docx',
    layoutMode: 'fixed-editable',
    editableTextRatio: scene.editableTextRatio,
    sceneStats: {
      objectCount: positionedObjectCount,
      pdfSceneObjects: scene.objectCount,
      textRuns: scene.textRuns,
      images: scene.images,
    },
  };
}
