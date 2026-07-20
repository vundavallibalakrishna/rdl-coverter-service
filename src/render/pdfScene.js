import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { PNG } from 'pngjs';
import { ServiceError } from '../errors.js';

const STANDARD_FONT_DATA_URL = `${fileURLToPath(new URL('../../node_modules/pdfjs-dist/standard_fonts/', import.meta.url)).replace(/\/+$/, '')}/`;
const nodeRequire = createRequire(import.meta.url);

// PDF.js uses the Node 22.3+ process.getBuiltinModule API to load its optional canvas polyfills and
// standard fonts. The service supports all Node 22 releases, so provide the equivalent read-only
// builtin lookup on early Node 22 rather than emitting warnings or losing font metrics.
if (typeof process.getBuiltinModule !== 'function') {
  Object.defineProperty(process, 'getBuiltinModule', {
    value: (name) => nodeRequire(name),
    configurable: true,
  });
}

function channel(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 0;
  return numeric <= 1 ? Math.round(Math.max(0, Math.min(1, numeric)) * 255) : Math.round(Math.max(0, Math.min(255, numeric)));
}

function rgbHex(values) {
  const entries = Array.from(values || []).slice(0, 3).map(channel);
  while (entries.length < 3) entries.push(0);
  return entries.map((entry) => entry.toString(16).padStart(2, '0')).join('').toUpperCase();
}

function grayHex(value) {
  const entry = channel(value);
  return [entry, entry, entry].map((part) => part.toString(16).padStart(2, '0')).join('').toUpperCase();
}

function cmykHex(values) {
  const [c = 0, m = 0, y = 0, k = 0] = Array.from(values || []).map((value) => Math.max(0, Math.min(1, Number(value))));
  return rgbHex([
    255 * (1 - c) * (1 - k),
    255 * (1 - m) * (1 - k),
    255 * (1 - y) * (1 - k),
  ]);
}

function decodeGlyphs(args) {
  const glyphs = Array.isArray(args?.[0]) ? args[0] : [];
  return glyphs
    .filter((entry) => typeof entry !== 'number')
    .map((entry) => entry?.unicode || '')
    .join('');
}

function normalizeText(value) {
  return String(value || '').normalize('NFKC').trim().replace(/\s+/g, ' ');
}

function mapTextStyles(items, sources, unsupported) {
  const mapped = [];
  let sourceIndex = 0;
  let remaining = normalizeText(sources[0]?.text);
  for (const item of items) {
    const target = normalizeText(item.str);
    while (sourceIndex < sources.length && !remaining) {
      sourceIndex += 1;
      remaining = normalizeText(sources[sourceIndex]?.text);
    }
    const source = sources[sourceIndex];
    if (!source) {
      unsupported.add('PDF_TEXT_MAPPING');
      mapped.push({});
      continue;
    }
    if (remaining === target) {
      mapped.push(source);
      remaining = '';
      continue;
    }
    if (remaining.startsWith(`${target} `)) {
      mapped.push(source);
      remaining = remaining.slice(target.length).trimStart();
      continue;
    }
    // PDF.js can also combine adjacent showText operations into one item. This is safe only when the
    // operations use the same visual style; otherwise one editable Word run could not represent them.
    if (target.startsWith(remaining)) {
      let combined = remaining;
      let end = sourceIndex;
      while (combined.length < target.length && end + 1 < sources.length) {
        const next = sources[end + 1];
        if (next.color !== source.color || next.fontName !== source.fontName) break;
        combined = `${combined} ${normalizeText(next.text)}`.trim();
        end += 1;
      }
      if (combined === target) {
        mapped.push(source);
        sourceIndex = end;
        remaining = '';
        continue;
      }
    }
    unsupported.add('PDF_TEXT_MAPPING');
    mapped.push(source);
    remaining = '';
  }
  while (sourceIndex < sources.length && !remaining) {
    sourceIndex += 1;
    remaining = normalizeText(sources[sourceIndex]?.text);
  }
  if (sourceIndex < sources.length || remaining) unsupported.add('PDF_TEXT_MAPPING');
  return mapped;
}

function fontMetadata(font, fallbackFamily = 'Arial') {
  const raw = String(font?.name || font?.fallbackName || fallbackFamily)
    .replace(/^[A-Z]{6}\+/, '');
  const bold = /bold|black|semibold|demibold/i.test(raw);
  const italic = /italic|oblique/i.test(raw);
  let family;
  if (/arial/i.test(raw)) family = 'Arial';
  else if (/timesnewroman|times new roman|times-roman/i.test(raw)) family = 'Times New Roman';
  else if (/segoe.*symbol/i.test(raw)) family = 'Segoe UI Symbol';
  else if (/segoe/i.test(raw)) family = 'Segoe UI';
  else if (/helvetica/i.test(raw)) family = 'Arial';
  else if (/times/i.test(raw)) family = 'Times New Roman';
  else {
    family = raw
      .replace(/-?(Bold|Italic|Oblique|Regular|Roman|Medium|Light|Black|SemiBold|DemiBold).*$/i, '')
      .replace(/PSMT$|MT$/i, '')
      .trim() || fallbackFamily;
  }
  return { family, bold, italic };
}

function getObject(store, id) {
  return new Promise((resolve, reject) => {
    try {
      store.get(id, resolve);
    } catch (error) {
      reject(error);
    }
  });
}

function imageToPng(image) {
  if (!image?.data || !Number.isInteger(image.width) || !Number.isInteger(image.height)) {
    throw new ServiceError('UNSUPPORTED_FEATURE', 'PDF image data cannot be converted to editable Word', 400, { feature: 'PDF_IMAGE_DATA' });
  }
  const png = new PNG({ width: image.width, height: image.height });
  const source = image.data;
  if (image.kind === 3 || source.length === image.width * image.height * 4) {
    png.data.set(source.subarray(0, png.data.length));
  } else if (image.kind === 2 || source.length === image.width * image.height * 3) {
    for (let input = 0, output = 0; output < png.data.length; input += 3, output += 4) {
      png.data[output] = source[input];
      png.data[output + 1] = source[input + 1];
      png.data[output + 2] = source[input + 2];
      png.data[output + 3] = 255;
    }
  } else if (image.kind === 1) {
    const rowBytes = Math.ceil(image.width / 8);
    for (let y = 0; y < image.height; y += 1) {
      for (let x = 0; x < image.width; x += 1) {
        const bit = (source[y * rowBytes + Math.floor(x / 8)] >> (7 - (x % 8))) & 1;
        const value = bit ? 255 : 0;
        const output = (y * image.width + x) * 4;
        png.data[output] = value;
        png.data[output + 1] = value;
        png.data[output + 2] = value;
        png.data[output + 3] = 255;
      }
    }
  } else {
    throw new ServiceError('UNSUPPORTED_FEATURE', 'Unsupported PDF image pixel format', 400, { feature: 'PDF_IMAGE_FORMAT' });
  }
  return PNG.sync.write(png, { colorType: 6 });
}

function parsePath(operators, coordinates, OPS, unsupported) {
  const shapes = [];
  let coordinateIndex = 0;
  let currentPoint = null;
  let subpathStart = null;
  for (const operator of operators || []) {
    if (operator === OPS.rectangle) {
      const [x, y, width, height] = coordinates.slice(coordinateIndex, coordinateIndex + 4);
      coordinateIndex += 4;
      shapes.push({ type: 'rectangle', x, y, width, height });
      currentPoint = { x, y };
      subpathStart = currentPoint;
    } else if (operator === OPS.moveTo) {
      const [x, y] = coordinates.slice(coordinateIndex, coordinateIndex + 2);
      coordinateIndex += 2;
      currentPoint = { x, y };
      subpathStart = currentPoint;
      shapes.push({ type: 'move', x, y });
    } else if (operator === OPS.lineTo) {
      const [x, y] = coordinates.slice(coordinateIndex, coordinateIndex + 2);
      coordinateIndex += 2;
      if (currentPoint) shapes.push({ type: 'line', x1: currentPoint.x, y1: currentPoint.y, x2: x, y2: y });
      currentPoint = { x, y };
    } else if ([OPS.curveTo, OPS.curveTo2, OPS.curveTo3].includes(operator)) {
      if (!currentPoint) {
        unsupported.add('constructPath:curveWithoutMove');
        continue;
      }
      let control1;
      let control2;
      let end;
      if (operator === OPS.curveTo) {
        const values = coordinates.slice(coordinateIndex, coordinateIndex + 6);
        coordinateIndex += 6;
        control1 = { x: values[0], y: values[1] };
        control2 = { x: values[2], y: values[3] };
        end = { x: values[4], y: values[5] };
      } else {
        const values = coordinates.slice(coordinateIndex, coordinateIndex + 4);
        coordinateIndex += 4;
        end = { x: values[2], y: values[3] };
        control1 = operator === OPS.curveTo2 ? currentPoint : { x: values[0], y: values[1] };
        control2 = operator === OPS.curveTo2 ? { x: values[0], y: values[1] } : end;
      }
      const start = currentPoint;
      for (let step = 1; step <= 16; step += 1) {
        const t = step / 16;
        const inverse = 1 - t;
        const point = {
          x: (inverse ** 3) * start.x + 3 * (inverse ** 2) * t * control1.x + 3 * inverse * (t ** 2) * control2.x + (t ** 3) * end.x,
          y: (inverse ** 3) * start.y + 3 * (inverse ** 2) * t * control1.y + 3 * inverse * (t ** 2) * control2.y + (t ** 3) * end.y,
        };
        shapes.push({ type: 'line', x1: currentPoint.x, y1: currentPoint.y, x2: point.x, y2: point.y });
        currentPoint = point;
      }
    } else if (operator === OPS.closePath) {
      if (currentPoint && subpathStart && (currentPoint.x !== subpathStart.x || currentPoint.y !== subpathStart.y)) {
        shapes.push({ type: 'line', x1: currentPoint.x, y1: currentPoint.y, x2: subpathStart.x, y2: subpathStart.y });
      }
      currentPoint = subpathStart;
    } else {
      const name = Object.entries(OPS).find(([, value]) => value === operator)?.[0] || String(operator);
      unsupported.add(`constructPath:${name}`);
      return [];
    }
  }
  return shapes;
}

function pathCanFill(path) {
  return path.length > 0 && path.every((shape) => shape.type === 'rectangle');
}

function polygonFillRectangles(path, fill, z) {
  const subpaths = [];
  let current = [];
  for (const shape of path) {
    if (shape.type === 'move') {
      if (current.length >= 3) subpaths.push(current);
      current = [{ x: shape.x, y: shape.y }];
    } else if (shape.type === 'line') {
      if (!current.length) current.push({ x: shape.x1, y: shape.y1 });
      current.push({ x: shape.x2, y: shape.y2 });
    }
  }
  if (current.length >= 3) subpaths.push(current);
  if (!subpaths.length) return [];
  const minY = Math.min(...subpaths.flat().map((point) => point.y));
  const maxY = Math.max(...subpaths.flat().map((point) => point.y));
  const step = 0.75; // one Word display pixel; keeps the DrawingML package bounded and visually smooth
  const rectangles = [];
  for (let top = Math.floor(minY / step) * step; top < maxY; top += step) {
    const scanY = top + step / 2;
    const intersections = [];
    for (const points of subpaths) {
      for (let index = 0; index < points.length; index += 1) {
        const first = points[index];
        const second = points[(index + 1) % points.length];
        if ((first.y <= scanY && second.y > scanY) || (second.y <= scanY && first.y > scanY)) {
          intersections.push(first.x + ((scanY - first.y) * (second.x - first.x)) / (second.y - first.y));
        }
      }
    }
    intersections.sort((left, right) => left - right);
    for (let index = 0; index + 1 < intersections.length; index += 2) {
      const left = intersections[index];
      const right = intersections[index + 1];
      if (right - left > 0.05) rectangles.push({
        type: 'shape', shape: 'rectangle', left, top, width: right - left, height: step,
        fill, stroke: null, strokeWidth: 0, dash: null, z,
      });
    }
  }
  return rectangles;
}

function addPathObjects(target, path, state, operation, z, unsupported) {
  const fills = operation === 'fill' || operation === 'fillStroke';
  const strokes = operation === 'stroke' || operation === 'fillStroke';
  if (fills && !pathCanFill(path)) {
    const rectangles = polygonFillRectangles(path, state.fillColor, z);
    if (!rectangles.length) unsupported.add('fill:nonRectangularPath');
    else target.push(...rectangles);
  }
  for (const shape of path) {
    if (shape.type === 'rectangle') {
      target.push({
        type: 'shape', shape: 'rectangle',
        left: Math.min(shape.x, shape.x + shape.width), top: Math.min(shape.y, shape.y + shape.height),
        width: Math.abs(shape.width), height: Math.abs(shape.height),
        fill: fills ? state.fillColor : null,
        stroke: strokes ? state.strokeColor : null,
        strokeWidth: strokes ? state.lineWidth : 0,
        dash: strokes ? state.dash : null,
        z,
      });
    } else if (shape.type === 'line' && strokes) {
      const horizontal = Math.abs(shape.y2 - shape.y1) <= 0.01;
      const vertical = Math.abs(shape.x2 - shape.x1) <= 0.01;
      target.push({
        type: 'shape', shape: 'line',
        left: Math.min(shape.x1, shape.x2), top: Math.min(shape.y1, shape.y2),
        width: Math.abs(shape.x2 - shape.x1), height: Math.abs(shape.y2 - shape.y1),
        x1: shape.x1, y1: shape.y1, x2: shape.x2, y2: shape.y2,
        diagonal: !horizontal && !vertical,
        fill: state.strokeColor,
        stroke: null,
        strokeWidth: state.lineWidth,
        dash: state.dash,
        z,
      });
    }
  }
}

function supportedOperatorNames(OPS) {
  return new Set([
    'dependency', 'save', 'restore', 'transform',
    'setLineWidth', 'setLineCap', 'setLineJoin', 'setMiterLimit', 'setDash',
    'setFillRGBColor', 'setStrokeRGBColor', 'setFillGray', 'setStrokeGray', 'setFillCMYKColor', 'setStrokeCMYKColor',
    'constructPath', 'stroke', 'closeStroke', 'fill', 'eoFill', 'fillStroke', 'eoFillStroke', 'closeFillStroke', 'closeEOFillStroke',
    'endPath', 'clip', 'eoClip',
    'beginText', 'endText', 'setCharSpacing', 'setWordSpacing', 'setHScale', 'setLeading', 'setFont',
    'setTextRenderingMode', 'setTextRise', 'moveText', 'setLeadingMoveText', 'setTextMatrix', 'nextLine',
    'showText', 'showSpacedText', 'nextLineShowText', 'nextLineSetSpacingShowText',
    'paintImageXObject', 'markPoint', 'markPointProps', 'beginMarkedContent', 'beginMarkedContentProps', 'endMarkedContent',
    'beginCompat', 'endCompat',
  ].filter((name) => OPS[name] !== undefined));
}

async function extractPage(page, pageNumber, pdfjs, config, defaultFontFamily) {
  const { OPS } = pdfjs;
  const [textContent, operatorList] = await Promise.all([
    page.getTextContent({ disableNormalization: false }),
    page.getOperatorList(),
  ]);
  const viewport = page.getViewport({ scale: 1 });
  const operatorNames = Object.fromEntries(Object.entries(OPS).map(([name, value]) => [value, name]));
  const supportedNames = supportedOperatorNames(OPS);
  const unsupported = new Set();
  const objects = [];
  const textStyles = [];
  const state = { fillColor: '000000', strokeColor: '000000', lineWidth: 1, dash: null, fontName: null, fontSize: 10 };
  const stack = [];
  let path = [];
  let lastTransform = null;

  for (let index = 0; index < operatorList.fnArray.length; index += 1) {
    const operation = operatorList.fnArray[index];
    const args = operatorList.argsArray[index];
    const name = operatorNames[operation] || `operator-${operation}`;
    if (!supportedNames.has(name)) unsupported.add(name);
    if (operation === OPS.save) stack.push({ ...state, dash: state.dash ? [...state.dash] : null, lastTransform });
    else if (operation === OPS.restore) {
      const previous = stack.pop();
      if (previous) {
        Object.assign(state, previous);
        lastTransform = previous.lastTransform;
      }
    } else if (operation === OPS.transform) lastTransform = Array.from(args || []);
    else if (operation === OPS.setLineWidth) state.lineWidth = Math.max(0.05, Number(args?.[0]) || 1);
    else if (operation === OPS.setDash) state.dash = Array.from(args?.[0] || []).map(Number).filter((value) => value > 0);
    else if (operation === OPS.setFillRGBColor) state.fillColor = rgbHex(args);
    else if (operation === OPS.setStrokeRGBColor) state.strokeColor = rgbHex(args);
    else if (operation === OPS.setFillGray) state.fillColor = grayHex(args?.[0]);
    else if (operation === OPS.setStrokeGray) state.strokeColor = grayHex(args?.[0]);
    else if (operation === OPS.setFillCMYKColor) state.fillColor = cmykHex(args);
    else if (operation === OPS.setStrokeCMYKColor) state.strokeColor = cmykHex(args);
    else if (operation === OPS.setFont) {
      state.fontName = args?.[0] || null;
      state.fontSize = Number(args?.[1]) || state.fontSize;
    } else if (operation === OPS.showText) {
      const text = decodeGlyphs(args);
      if (text.trim()) textStyles.push({ text, color: state.fillColor, fontName: state.fontName, z: index });
    } else if (operation === OPS.constructPath) {
      path = parsePath(args?.[0], Array.from(args?.[1] || []), OPS, unsupported);
    } else if ([OPS.fill, OPS.eoFill].includes(operation)) {
      addPathObjects(objects, path, state, 'fill', index, unsupported);
      path = [];
    } else if ([OPS.stroke, OPS.closeStroke].includes(operation)) {
      addPathObjects(objects, path, state, 'stroke', index, unsupported);
      path = [];
    } else if ([OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke].includes(operation)) {
      addPathObjects(objects, path, state, 'fillStroke', index, unsupported);
      path = [];
    } else if (operation === OPS.endPath) path = [];
    else if (operation === OPS.paintImageXObject) {
      const id = args?.[0];
      const transform = lastTransform;
      if (!id || !transform || transform.length !== 6 || Math.abs(transform[1]) > 0.01 || Math.abs(transform[2]) > 0.01) {
        unsupported.add('paintImageXObject:transformed');
        continue;
      }
      const image = await getObject(page.objs, id);
      const data = imageToPng(image);
      const [a, , , d, e, f] = transform;
      objects.push({
        type: 'image', data,
        left: Math.min(e, e + a), top: Math.min(f, f + d),
        width: Math.abs(a), height: Math.abs(d), z: index,
      });
    }
  }

  const visibleText = textContent.items.filter((item) => typeof item.str === 'string' && item.str.trim().length > 0);
  const mappedTextStyles = mapTextStyles(visibleText, textStyles, unsupported);
  const fonts = new Map();
  for (const name of new Set(visibleText.map((item) => item.fontName))) {
    try {
      const font = await getObject(page.commonObjs, name);
      fonts.set(name, fontMetadata(font, defaultFontFamily));
    } catch {
      fonts.set(name, fontMetadata(null, defaultFontFamily));
    }
  }
  for (let index = 0; index < visibleText.length; index += 1) {
    const item = visibleText[index];
    const sourceStyle = mappedTextStyles[index] || {};
    const style = fonts.get(item.fontName) || fontMetadata(null, defaultFontFamily);
    const [a, b, c, d, e, f] = item.transform;
    if (Math.abs(b) > 0.01 || Math.abs(c) > 0.01) {
      unsupported.add('PDF_TEXT_ROTATION');
      continue;
    }
    const fontSize = Math.max(1, Math.abs(d) || Math.abs(a) || Number(item.height) || 10);
    const metrics = textContent.styles[item.fontName] || {};
    const ascent = Number.isFinite(metrics.ascent) ? metrics.ascent : 0.82;
    const descent = Number.isFinite(metrics.descent) ? metrics.descent : -0.22;
    const lineHeight = Math.max(fontSize, fontSize * (ascent - descent));
    objects.push({
      type: 'text', text: item.str,
      left: e, top: viewport.height - f - (fontSize * ascent),
      width: Math.max(Number(item.width) || 0, 0.5), height: lineHeight,
      fontFamily: style.family, fontSize, bold: style.bold, italic: style.italic,
      color: sourceStyle.color || '000000', z: sourceStyle.z ?? (operatorList.fnArray.length + index),
    });
  }

  const textRuns = objects.filter((object) => object.type === 'text').length;
  const images = objects.filter((object) => object.type === 'image').length;
  if (objects.length > config.maxFixedObjects) throw new ServiceError('UNSUPPORTED_FEATURE', 'Fixed editable Word object limit exceeded', 413, { limit: config.maxFixedObjects, actual: objects.length });
  if (textRuns > config.maxFixedTextRuns) throw new ServiceError('UNSUPPORTED_FEATURE', 'Fixed editable Word text-run limit exceeded', 413, { limit: config.maxFixedTextRuns, actual: textRuns });
  if (images > config.maxFixedImages) throw new ServiceError('UNSUPPORTED_FEATURE', 'Fixed editable Word image limit exceeded', 413, { limit: config.maxFixedImages, actual: images });

  return {
    number: pageNumber, width: viewport.width, height: viewport.height,
    objects: objects.sort((left, right) => left.z - right.z),
    unsupportedPdfOperators: [...unsupported].sort(),
    textRuns,
    images,
  };
}

export async function extractPdfScene(pdfBuffer, config, options = {}) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(pdfBuffer),
    disableFontFace: true,
    isEvalSupported: false,
    useSystemFonts: false,
    standardFontDataUrl: STANDARD_FONT_DATA_URL,
    stopEvent: true,
  });
  try {
    const document = await loadingTask.promise;
    if (document.numPages > config.maxFixedPages) {
      throw new ServiceError('UNSUPPORTED_FEATURE', 'Fixed editable Word page limit exceeded', 413, { limit: config.maxFixedPages, actual: document.numPages });
    }
    const pages = [];
    for (let number = 1; number <= document.numPages; number += 1) {
      const page = await document.getPage(number);
      pages.push(await extractPage(page, number, pdfjs, config, options.defaultFontFamily || 'Arial'));
      page.cleanup();
    }
    const unsupportedPdfOperators = [...new Set(pages.flatMap((page) => page.unsupportedPdfOperators))].sort();
    const textRuns = pages.reduce((sum, page) => sum + page.textRuns, 0);
    const images = pages.reduce((sum, page) => sum + page.images, 0);
    const objectCount = pages.reduce((sum, page) => sum + page.objects.length, 0);
    if (objectCount > config.maxFixedObjects) {
      throw new ServiceError('UNSUPPORTED_FEATURE', 'Fixed editable Word object limit exceeded', 413, { limit: config.maxFixedObjects, actual: objectCount });
    }
    return {
      pageCount: pages.length,
      pages,
      objectCount,
      textRuns,
      images,
      unsupportedPdfOperators,
      editableTextRatio: unsupportedPdfOperators.some((operator) => operator.startsWith('PDF_TEXT')) ? 0 : 1,
    };
  } finally {
    await loadingTask.destroy().catch(() => {});
  }
}
