import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import {
  AlignmentType,
  BorderStyle,
  BuilderElement,
  Document,
  Footer,
  HeightRule,
  HorizontalPositionRelativeFrom,
  ImageRun,
  LineRuleType,
  PageOrientation,
  Packer,
  Paragraph,
  SectionType,
  ShadingType,
  Table,
  TableBorders,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  TextDirection,
  TextWrappingType,
  UnderlineType,
  VerticalAlignTable,
  VerticalMergeType,
  VerticalPositionRelativeFrom,
  WidthType,
} from 'docx';
import JSZip from 'jszip';
import { loadConfig } from '../config.js';
import { ServiceError } from '../errors.js';
import { pointsToTwips } from '../units.js';
import { normalizeDatasets } from './common.js';
import { materializeChart } from './chartData.js';
import { renderChartPng } from './chartImage.js';
import { editableFontEmbeddingPermission, resolveFontFile } from './fonts.js';
import { renderPdf } from './pdf.js';
import { buildGridBoundaries } from './gridBoundaries.js';
import { validateLayoutTrace } from './layoutTrace.js';
import { validateWindowsWordRequest } from './windowsWordCompatibility.js';

const WORD_MAX_PAGE_POINTS = 22 * 72;
const WORD_MAX_TABLE_COLUMNS = 63;
const WORD_MAX_TABLE_ROWS = 32_767;
const GRID_PRECISION_POINTS = 0.25;
// Word requires a terminal section paragraph after the page canvas table. Reserve a deterministic 2pt
// band for that paragraph and table-border rounding; report content normally ends above the declared
// bottom margin/footer. A PDF item that actually occupies this band fails closed below.
const SECTION_ANCHOR_POINTS = 2;
const GEOMETRY_EPSILON = 0.13;
const CERTIFIED_GEOMETRY_TOLERANCE_POINTS = 0.5;
const NONE_BORDER = Object.freeze({ style: BorderStyle.NONE, size: 0, color: 'auto' });
const VARIANTS = Object.freeze([
  { key: 'regular', bold: false, italic: false, element: 'embedRegular' },
  { key: 'bold', bold: true, italic: false, element: 'embedBold' },
  { key: 'italic', bold: false, italic: true, element: 'embedItalic' },
  { key: 'boldItalic', bold: true, italic: true, element: 'embedBoldItalic' },
]);

function unsupported(message, details) {
  throw new ServiceError('UNSUPPORTED_FEATURE', message, 422, details);
}

function snap(value) {
  return Math.round(Number(value || 0) / GRID_PRECISION_POINTS) * GRID_PRECISION_POINTS;
}

function pointsToDrawingPixels(points) {
  // docx accepts fractional CSS-pixel dimensions and converts them to integer DrawingML EMUs. Rounding
  // here first can make an image larger than its exact-height Word row: for example, 42.5pt rounds from
  // 56.667px to 57px (42.75pt), so Word clips it until the row is manually enlarged. Preserve the point
  // measurement through the EMU conversion instead.
  return (Number(points || 0) / 72) * 96;
}

function pointsToDrawingEmus(points) {
  return Math.round(Number(points || 0) * 12_700);
}

function pointsToInches(points) {
  return Math.round((Number(points || 0) / 72) * 1000) / 1000;
}

function cleanColor(value, fallback = '000000') {
  const normalized = String(value || fallback).replace(/^#/, '').trim();
  return /^[0-9a-f]{6}$/i.test(normalized) ? normalized.toUpperCase() : fallback;
}

// The grid must address every traced edge, but two edges closer than the certification tolerance are the
// same Word grid line: Word cannot render a table band that narrow and separates the two cell borders
// instead, turning one canonical rule into a double line. See gridBoundaries.js for the full reasoning.
function pageGridAxis(values, protectedSpans) {
  return buildGridBoundaries(values.map(snap), { protectedSpans });
}

function boundaryIndex(axis, value) {
  const index = axis.indexOf(snap(value));
  if (index < 0) unsupported('PDF layout geometry cannot be represented by a stable Word table grid', { value });
  return index;
}

function positiveOverlap(left, right) {
  const width = Math.min(left.x + left.width, right.x + right.width) - Math.max(left.x, right.x);
  const height = Math.min(left.y + left.height, right.y + right.height) - Math.max(left.y, right.y);
  return width > GEOMETRY_EPSILON && height > GEOMETRY_EPSILON;
}

function contains(outer, inner) {
  return inner.x >= outer.x - GEOMETRY_EPSILON
    && inner.y >= outer.y - GEOMETRY_EPSILON
    && inner.x + inner.width <= outer.x + outer.width + GEOMETRY_EPSILON
    && inner.y + inner.height <= outer.y + outer.height + GEOMETRY_EPSILON;
}

function isEmptyCell(item) {
  return item.kind === 'tablixCell' && !String(item.text || '') && (item.lines || []).length === 0;
}

function borderWidth(item, side) {
  const border = item?.borders?.[side];
  if (!border || /^none$/i.test(String(border.style || 'None'))) return 0;
  return Math.max(0, Number(border.width || 0));
}

function textPaintBounds(item) {
  const lines = item?.lines || [];
  const runs = lines.flatMap((line) => line.runs || []);
  if (runs.length === 0 && lines.length === 0) return null;
  const left = Math.min(...lines.map((line) => Number(line.x ?? item.x)));
  const top = Math.min(...lines.map((line) => Number(line.y ?? item.y)));
  const right = Math.max(...lines.map((line) => (
    Number(line.x ?? item.x) + Math.max(0, Number(line.width || 0))
  )));
  const bottom = Math.max(...lines.map((line) => (
    Number(line.y ?? item.y) + Math.max(0, Number(line.height || line.contentHeight || 0))
  )));
  // PDF clips every textbox to its declared box. A glyph's measured advance can extend beyond that box,
  // but the clipped pixels do not participate in a visible overlap and must not block a native Word grid.
  return {
    left: Math.max(Number(item.x || 0), left),
    top: Math.max(Number(item.y || 0), top),
    right: Math.min(Number(item.x || 0) + Number(item.width || 0), right),
    bottom: Math.min(Number(item.y || 0) + Number(item.height || 0), bottom),
  };
}

function isUnpaintedOwner(item) {
  return !item.backgroundColor && !Object.values(item.borders || {}).some(Boolean);
}

function primaryTextAlignment(item) {
  return String(item?.lines?.find((line) => line.runs?.length > 0)?.alignment || 'left').toLowerCase();
}

function horizontalTrimPreservesFlow(item, side, reduction) {
  if (reduction <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS) return true;
  const alignmentValue = primaryTextAlignment(item);
  const remainingTrim = reduction - Number(item.padding?.[side] || 0);
  // Reducing the padding on the same side by the trim amount keeps the physical text-content
  // rectangle unchanged. That preserves left, center, and right alignment alike; the outer native
  // cell merely discards an unpainted edge strip. Without this check a centered icon adjacent to a
  // label is falsely rejected even when its declared side padding fully absorbs the overlap.
  if (remainingTrim <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS + GEOMETRY_EPSILON) return true;
  // Centered text moves by half of the remaining outer-edge trim. Permit that only while the
  // resulting position remains inside the 0.5pt certification geometry tolerance.
  if (alignmentValue === 'center'
    && remainingTrim <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS * 2 + GEOMETRY_EPSILON) return true;
  if (side === 'right') return alignmentValue === 'left';
  if (alignmentValue === 'right') return true;
  return false;
}

function verticalTrimPreservesFlow(item, side, reduction) {
  if (reduction <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS) return true;
  const vertical = String(item.verticalAlign || 'top').toLowerCase();
  const remainingTrim = reduction - Number(item.padding?.[side] || 0);
  // See horizontalTrimPreservesFlow: same-side padding compensation preserves the complete content
  // rectangle, including middle-aligned text.
  if (remainingTrim <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS + GEOMETRY_EPSILON) return true;
  if (/middle|center/.test(vertical)
    && remainingTrim <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS * 2 + GEOMETRY_EPSILON) return true;
  if (side === 'bottom') {
    return /top/.test(vertical)
      || false;
  }
  return /bottom/.test(vertical);
}

function adjustPadding(item, side, reduction) {
  if (!item.padding || reduction <= 0) return;
  item.padding = {
    ...item.padding,
    [side]: Math.max(0, Number(item.padding[side] || 0) - reduction),
  };
}

function coalesceVerticalEdge(first, second) {
  const upper = first.y <= second.y ? first : second;
  const lower = upper === first ? second : first;
  if (contains(upper, lower) || contains(lower, upper)) return null;
  const overlap = upper.y + upper.height - lower.y;
  if (overlap <= GEOMETRY_EPSILON) return null;
  const upperPaint = textPaintBounds(upper);
  const lowerPaint = textPaintBounds(lower);
  const paintSafe = isUnpaintedOwner(upper)
    && isUnpaintedOwner(lower)
    && (!upperPaint || !lowerPaint
      || upperPaint.bottom <= lowerPaint.top + GEOMETRY_EPSILON);
  const borderAllowance = Math.max(
    borderWidth(upper, 'bottom'),
    borderWidth(lower, 'top'),
    GRID_PRECISION_POINTS,
  ) + GRID_PRECISION_POINTS;
  const maximumOverlap = Math.min(
    CERTIFIED_GEOMETRY_TOLERANCE_POINTS * 2,
    borderAllowance,
  );
  if (!paintSafe && overlap > maximumOverlap + GEOMETRY_EPSILON) return null;

  const sharedEdge = snap(paintSafe && upperPaint && lowerPaint
    ? (upperPaint.bottom + lowerPaint.top) / 2
    : ((upper.y + upper.height) + lower.y) / 2);
  if ((upperPaint && upperPaint.bottom > sharedEdge + GEOMETRY_EPSILON)
    || (lowerPaint && lowerPaint.top < sharedEdge - GEOMETRY_EPSILON)) return null;

  const upperBottom = upper.y + upper.height;
  const lowerBottom = lower.y + lower.height;
  const upperReduction = Math.max(0, upperBottom - sharedEdge);
  const lowerReduction = Math.max(0, sharedEdge - lower.y);
  if (!verticalTrimPreservesFlow(upper, 'bottom', upperReduction)
    || !verticalTrimPreservesFlow(lower, 'top', lowerReduction)) return null;
  upper.height = Math.max(0, sharedEdge - upper.y);
  lower.y = sharedEdge;
  lower.height = Math.max(0, lowerBottom - sharedEdge);
  adjustPadding(upper, 'bottom', upperReduction);
  adjustPadding(lower, 'top', lowerReduction);
  return {
    axis: 'vertical',
    first: upper.itemName,
    second: lower.itemName,
    originalOverlap: overlap,
    sharedEdge,
    sourceEdges: [
      {
        side: 'bottom',
        from: upperBottom,
        to: sharedEdge,
        start: Math.max(upper.x, lower.x),
        end: Math.min(upper.x + upper.width, lower.x + lower.width),
      },
      {
        side: 'top',
        from: lower.y - lowerReduction,
        to: sharedEdge,
        start: Math.max(upper.x, lower.x),
        end: Math.min(upper.x + upper.width, lower.x + lower.width),
      },
    ],
  };
}

function coalesceHorizontalEdge(first, second) {
  const left = first.x <= second.x ? first : second;
  const right = left === first ? second : first;
  if (contains(left, right) || contains(right, left)) return null;
  const overlap = left.x + left.width - right.x;
  if (overlap <= GEOMETRY_EPSILON) return null;
  const leftPaint = textPaintBounds(left);
  const rightPaint = textPaintBounds(right);
  const paintSafe = isUnpaintedOwner(left)
    && isUnpaintedOwner(right)
    && (!leftPaint || !rightPaint
      || leftPaint.right <= rightPaint.left + GEOMETRY_EPSILON);
  const borderAllowance = Math.max(
    borderWidth(left, 'right'),
    borderWidth(right, 'left'),
    GRID_PRECISION_POINTS,
  ) + GRID_PRECISION_POINTS;
  const maximumOverlap = Math.min(
    CERTIFIED_GEOMETRY_TOLERANCE_POINTS * 2,
    borderAllowance,
  );
  if (!paintSafe && overlap > maximumOverlap + GEOMETRY_EPSILON) return null;

  const sharedEdge = snap(paintSafe && leftPaint && rightPaint
    ? (leftPaint.right + rightPaint.left) / 2
    : ((left.x + left.width) + right.x) / 2);
  if ((leftPaint && leftPaint.right > sharedEdge + GEOMETRY_EPSILON)
    || (rightPaint && rightPaint.left < sharedEdge - GEOMETRY_EPSILON)) return null;

  const leftRight = left.x + left.width;
  const rightBoundary = right.x + right.width;
  const leftReduction = Math.max(0, leftRight - sharedEdge);
  const rightReduction = Math.max(0, sharedEdge - right.x);
  if (!horizontalTrimPreservesFlow(left, 'right', leftReduction)
    || !horizontalTrimPreservesFlow(right, 'left', rightReduction)) return null;
  left.width = Math.max(0, sharedEdge - left.x);
  right.x = sharedEdge;
  right.width = Math.max(0, rightBoundary - sharedEdge);
  adjustPadding(left, 'right', leftReduction);
  adjustPadding(right, 'left', rightReduction);
  return {
    axis: 'horizontal',
    first: left.itemName,
    second: right.itemName,
    originalOverlap: overlap,
    sharedEdge,
    sourceEdges: [
      {
        side: 'right',
        from: leftRight,
        to: sharedEdge,
        start: Math.max(left.y, right.y),
        end: Math.min(left.y + left.height, right.y + right.height),
      },
      {
        side: 'left',
        from: right.x - rightReduction,
        to: sharedEdge,
        start: Math.max(left.y, right.y),
        end: Math.min(left.y + left.height, right.y + right.height),
      },
    ],
  };
}

function coalesceShallowEdgeOverlaps(items) {
  const adjustments = [];
  for (let leftIndex = 0; leftIndex < items.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < items.length; rightIndex += 1) {
      const left = items[leftIndex];
      const right = items[rightIndex];
      if (!positiveOverlap(left, right)) continue;
      if (!['textbox', 'tablixCell'].includes(left.kind)
        || !['textbox', 'tablixCell'].includes(right.kind)) continue;
      if (!/^(default|horizontal)?$/i.test(String(left.writingMode || 'default'))
        || !/^(default|horizontal)?$/i.test(String(right.writingMode || 'default'))) continue;

      const overlapWidth = Math.min(left.x + left.width, right.x + right.width)
        - Math.max(left.x, right.x);
      const overlapHeight = Math.min(left.y + left.height, right.y + right.height)
        - Math.max(left.y, right.y);
      // The smaller intersection is only a heuristic, not a representation rule. Two boxes can
      // have a smaller vertical intersection while still be horizontal neighbours (for example,
      // a short label overlapping the full height of an adjacent value cell). In that case their
      // text necessarily crosses a horizontal split, but their unpainted horizontal edge can be
      // coalesced without moving either painted run. Try the likely axis first, then the other
      // axis; each helper independently proves that its trim preserves the canonical paint.
      const preferVertical = overlapHeight <= overlapWidth;
      const adjustment = preferVertical
        ? coalesceVerticalEdge(left, right) || coalesceHorizontalEdge(left, right)
        : coalesceHorizontalEdge(left, right) || coalesceVerticalEdge(left, right);
      if (adjustment) adjustments.push(adjustment);
    }
  }
  return adjustments;
}

function moveCoalescedBorderLines(lines, adjustments) {
  for (const line of lines) {
    for (const adjustment of adjustments) {
      for (const edge of adjustment.sourceEdges || []) {
        if (adjustment.axis === 'vertical') {
          const horizontal = Math.abs(line.height) <= GEOMETRY_EPSILON;
          const coversEdge = line.x <= edge.start + GEOMETRY_EPSILON
            && line.x + line.width >= edge.end - GEOMETRY_EPSILON;
          if (horizontal && coversEdge && Math.abs(line.y - edge.from) <= GEOMETRY_EPSILON) {
            line.y = edge.to;
          }
          const vertical = Math.abs(line.width) <= GEOMETRY_EPSILON;
          const onPerpendicularBoundary = Math.abs(line.x - edge.start) <= GEOMETRY_EPSILON
            || Math.abs(line.x - edge.end) <= GEOMETRY_EPSILON;
          if (vertical && onPerpendicularBoundary && edge.side === 'top'
            && Math.abs(line.y - edge.from) <= GEOMETRY_EPSILON) {
            const bottom = line.y + line.height;
            line.y = edge.to;
            line.height = Math.max(0, bottom - edge.to);
          } else if (vertical && onPerpendicularBoundary && edge.side === 'bottom'
            && Math.abs(line.y + line.height - edge.from) <= GEOMETRY_EPSILON) {
            line.height = Math.max(0, edge.to - line.y);
          }
        } else {
          const vertical = Math.abs(line.width) <= GEOMETRY_EPSILON;
          const coversEdge = line.y <= edge.start + GEOMETRY_EPSILON
            && line.y + line.height >= edge.end - GEOMETRY_EPSILON;
          if (vertical && coversEdge && Math.abs(line.x - edge.from) <= GEOMETRY_EPSILON) {
            line.x = edge.to;
          }
          const horizontal = Math.abs(line.height) <= GEOMETRY_EPSILON;
          const onPerpendicularBoundary = Math.abs(line.y - edge.start) <= GEOMETRY_EPSILON
            || Math.abs(line.y - edge.end) <= GEOMETRY_EPSILON;
          if (horizontal && onPerpendicularBoundary && edge.side === 'left'
            && Math.abs(line.x - edge.from) <= GEOMETRY_EPSILON) {
            const right = line.x + line.width;
            line.x = edge.to;
            line.width = Math.max(0, right - edge.to);
          } else if (horizontal && onPerpendicularBoundary && edge.side === 'right'
            && Math.abs(line.x + line.width - edge.from) <= GEOMETRY_EPSILON) {
            line.width = Math.max(0, edge.to - line.x);
          }
        }
      }
    }
  }
}

function snapFooterDividersToContentEdges(lines, candidates) {
  for (const line of lines) {
    if (line.region !== 'footer' || Math.abs(line.height) > GEOMETRY_EPSILON) continue;
    const nextContentEdge = candidates
      .map((candidate) => candidate.y)
      .filter((edge) => edge > line.y + GEOMETRY_EPSILON
        && edge - line.y <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS + GEOMETRY_EPSILON)
      .sort((left, right) => left - right)[0];
    if (nextContentEdge !== undefined) line.y = nextContentEdge;
  }
}

function alignment(value) {
  const normalized = String(value || '').toLowerCase();
  if (normalized === 'center') return AlignmentType.CENTER;
  if (normalized === 'right') return AlignmentType.RIGHT;
  if (normalized === 'justify') return AlignmentType.JUSTIFIED;
  return AlignmentType.LEFT;
}

function verticalAlignment(value) {
  const normalized = String(value || '').toLowerCase();
  if (/middle|center/.test(normalized)) return VerticalAlignTable.CENTER;
  if (/bottom/.test(normalized)) return VerticalAlignTable.BOTTOM;
  return VerticalAlignTable.TOP;
}

function wordTextDirection(value) {
  const normalized = String(value || 'default').replace(/[\s_-]/g, '').toLowerCase();
  if (normalized === 'default' || normalized === 'horizontal') return undefined;
  if (normalized === 'rotate270') return TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT;
  if (normalized === 'vertical') return TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT;
  return null;
}

function borderStyle(style) {
  const normalized = String(style || 'Solid').replace(/[\s_-]/g, '').toLowerCase();
  if (normalized === 'double') return BorderStyle.DOUBLE;
  if (normalized === 'dotted') return BorderStyle.DOTTED;
  if (normalized === 'dashed') return BorderStyle.DASHED;
  if (normalized === 'dashdot') return BorderStyle.DOT_DASH;
  if (normalized === 'dashdotdot') return BorderStyle.DOT_DOT_DASH;
  return BorderStyle.SINGLE;
}

function wordBorder(border) {
  if (!border) return NONE_BORDER;
  return {
    style: borderStyle(border.style),
    size: Math.max(1, Math.round(Number(border.width || 1) * 8)),
    color: cleanColor(border.color),
  };
}

function strongerBorder(left, right) {
  if (!left) return right || null;
  if (!right) return left;
  const leftWidth = Number(left.width || 0);
  const rightWidth = Number(right.width || 0);
  return rightWidth >= leftWidth ? right : left;
}

function linesForParagraphs(item, bottomPaddingTwips = 0, fitTextCounter = null, topPaddingTwips = 0) {
  const source = item.lines || [];
  if (source.length === 0) return [new Paragraph({
    spacing: {
      before: Math.max(0, topPaddingTwips),
      after: Math.max(0, bottomPaddingTwips),
      line: 1,
      lineRule: LineRuleType.EXACT,
    },
    children: [new TextRun({ text: '' })],
  })];

  const paragraphGroups = [];
  let current = [];
  for (const line of source) {
    current.push(line);
    if (line.paragraphEnd) {
      paragraphGroups.push(current);
      current = [];
    }
  }
  if (current.length > 0) paragraphGroups.push(current);

  const linePitchTwips = (line) => Math.max(1, pointsToTwips(Number(
    line.contentHeight
      ?? Math.max(0, Number(line.height || 0) - Number(line.before || 0) - Number(line.after || 0)),
  ) || 0.05));

  // Word applies one line pitch to every hard-broken line in a paragraph. A traced PDF paragraph can
  // legitimately contain different physical line heights, for example a 10pt bold label on its first
  // line followed by wrapped 9pt parameter text. Using the tallest traced line for the complete Word
  // paragraph makes the fixed page-grid row too short even though the canonical lines fit. Split only
  // when the physical pitch changes; equal-pitch lines remain one editable paragraph with native breaks.
  const groups = [];
  for (const paragraphGroup of paragraphGroups) {
    let pitch = null;
    let segment = [];
    for (const line of paragraphGroup) {
      const nextPitch = linePitchTwips(line);
      if (segment.length > 0 && nextPitch !== pitch) {
        groups.push({ lines: segment, linePitchTwips: pitch });
        segment = [];
      }
      pitch = nextPitch;
      segment.push(line);
    }
    if (segment.length > 0) groups.push({ lines: segment, linePitchTwips: pitch });
  }

  return groups.map((group, groupIndex) => {
    const first = group.lines[0];
    const last = group.lines[group.lines.length - 1];
    const runs = [];
    group.lines.forEach((line, lineIndex) => {
      const lineRuns = line.runs?.length ? line.runs : [{ text: '', font: {} }];
      // PDFKit and Microsoft Word can produce slightly different glyph advances for the same embedded
      // font. The canonical trace has already selected the physical line and measured its exact width;
      // leaving Word to measure that text again can make a nearly-full line wrap a few words early and
      // then clip inside the trace-locked row height. WordprocessingML fitText is the native mechanism for
      // assigning a manual width to one or more contiguous runs. Give every run on the physical PDF line
      // the same id and width so mixed formatting remains editable while Word cannot choose a new wrap.
      const tracedLineWidthTwips = Number.isFinite(Number(line.width)) && Number(line.width) > 0
        ? Math.max(1, pointsToTwips(Number(line.width)))
        : null;
      const fitTextId = tracedLineWidthTwips !== null && fitTextCounter
        ? fitTextCounter.value++
        : null;
      lineRuns.forEach((run, runIndex) => {
        const font = run.font || {};
        const textRun = new TextRun({
          text: String(run.text ?? ''),
          break: lineIndex > 0 && runIndex === 0 ? 1 : undefined,
          font: font.family || 'Arial',
          size: Math.max(1, Math.round(Number(font.size || 10) * 2)),
          bold: Boolean(font.bold),
          italics: Boolean(font.italic),
          underline: font.underline ? { type: UnderlineType.SINGLE } : undefined,
          strike: Boolean(font.strike),
          color: cleanColor(font.color),
          characterSpacing: 0,
        });
        if (fitTextId !== null) {
          textRun.root[0].root.push(new BuilderElement({
            name: 'w:fitText',
            attributes: {
              id: { key: 'w:id', value: fitTextId },
              val: { key: 'w:val', value: tracedLineWidthTwips },
            },
          }));
        }
        runs.push(textRun);
      });
    });
    return new Paragraph({
      alignment: alignment(first.alignment),
      spacing: {
        before: Math.max(
          0,
          pointsToTwips(first.before || 0) + (groupIndex === 0 ? topPaddingTwips : 0),
        ),
        after: Math.max(
          0,
          pointsToTwips(last.after || 0)
            + (groupIndex === groups.length - 1 ? bottomPaddingTwips : 0),
        ),
        line: group.linePitchTwips,
        lineRule: LineRuleType.EXACT,
      },
      children: runs.length > 0 ? runs : [new TextRun({ text: '' })],
    });
  });
}

function detectImageType(buffer) {
  if (buffer.length > 8 && buffer[0] === 0x89 && buffer.toString('ascii', 1, 4) === 'PNG') return 'png';
  if (buffer.length > 3 && buffer[0] === 0xFF && buffer[1] === 0xD8) return 'jpg';
  if (buffer.length > 6 && buffer.toString('ascii', 0, 3) === 'GIF') return 'gif';
  if (buffer.length > 2 && buffer[0] === 0x42 && buffer[1] === 0x4D) return 'bmp';
  return null;
}

function naturalImageSize(buffer) {
  if (buffer.length > 24 && buffer.toString('ascii', 1, 4) === 'PNG') {
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
  }
  if (buffer.length > 4 && buffer[0] === 0xFF && buffer[1] === 0xD8) {
    let index = 2;
    while (index < buffer.length - 8) {
      if (buffer[index] !== 0xFF) {
        index += 1;
        continue;
      }
      const marker = buffer[index + 1];
      if (marker >= 0xC0 && marker <= 0xCF && ![0xC4, 0xC8, 0xCC].includes(marker)) {
        return { height: buffer.readUInt16BE(index + 5), width: buffer.readUInt16BE(index + 7) };
      }
      index += 2 + buffer.readUInt16BE(index + 2);
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

function collectModels(model, result = []) {
  result.push(model);
  for (const child of Object.values(model.subreports || {})) collectModels(child, result);
  return result;
}

function collectItems(items, map) {
  for (const item of items || []) {
    if (item.name && !map.has(item.name)) map.set(item.name, item);
    collectItems(item.items, map);
    for (const row of item.rows || []) {
      for (const cell of row.cells || []) collectItems(cell.items, map);
    }
  }
}

function modelResources(model) {
  const embeddedImages = {};
  const items = new Map();
  for (const current of collectModels(model)) {
    Object.assign(embeddedImages, current.embeddedImages || {});
    collectItems(current.body?.items, items);
    collectItems(current.page?.header?.items, items);
    collectItems(current.page?.footer?.items, items);
  }
  return { embeddedImages, items };
}

async function pictureForItem(
  item,
  resources,
  model,
  request,
  config,
  tempDir,
  chartIndex,
  bottomPaddingTwips = 0,
  topPaddingTwips = 0,
) {
  let data;
  let type;
  if (item.kind === 'image') {
    const image = resources.embeddedImages[item.embeddedImage];
    if (!image?.data) unsupported('A PDF-traced embedded image is unavailable to the Word renderer', {
      item: item.itemName,
      embeddedImage: item.embeddedImage,
    });
    data = Buffer.from(image.data.replace(/\s+/g, ''), 'base64');
    type = detectImageType(data);
    if (!type) unsupported('The embedded image format cannot be represented safely in native Word', {
      item: item.itemName,
    });
  } else {
    const chart = resources.items.get(item.itemName);
    if (!chart || !config || !tempDir) {
      unsupported('A PDF-traced chart cannot be materialized as a native Word picture', { item: item.itemName });
    }
    const datasets = normalizeDatasets(model, request);
    const globals = {
      PageNumber: item.pageNumber || 1,
      TotalPages: item.totalPages || 1,
      ExecutionTime: new Date(),
      variables: model.variables || {},
    };
    const chartData = materializeChart(chart, datasets, request.parameters || {}, globals);
    const rendered = await renderChartPng(
      chart,
      chartData,
      config,
      tempDir,
      { datasets, parameters: request.parameters || {}, globals, fields: {}, dataset: [] },
      chartIndex,
    );
    if (!rendered?.data) unsupported('A PDF-traced chart could not be rendered for native Word', {
      item: item.itemName,
    });
    data = rendered.data;
    type = 'png';
  }

  let width = item.width;
  let height = item.height;
  const sizing = String(item.sizing || 'FitProportional');
  if (/^(Clip|AutoSize)$/i.test(sizing)) {
    unsupported(`RDL image sizing '${sizing}' is not safely representable in the page-locked editable Word contract`, {
      item: item.itemName,
    });
  }
  if (!/^Fit$/i.test(sizing)) {
    const natural = naturalImageSize(data);
    if (natural) {
      const scale = Math.min(item.width / natural.width, item.height / natural.height);
      width = natural.width * scale;
      height = natural.height * scale;
    }
  }
  return new Paragraph({
    // Inline drawings are aligned to a text baseline. In Microsoft Word an exact line as tall as a large
    // chart places that baseline inside the line box, so the picture can protrude upward into preceding
    // PDF regions even though its enclosing row is exact. Images and charts are the only report items that
    // the page-locked contract permits as drawings, so float them at the origin of their canonical owner
    // cell and keep their anchor paragraph physically negligible. Cell-relative positioning also remains
    // stable in Word footer stories and avoids page offsets being applied twice by alternate OOXML viewers.
    spacing: {
      // The drawing is anchored to this paragraph, so it must absorb any top margin the cell gave up.
      before: Math.max(0, topPaddingTwips),
      after: Math.max(0, bottomPaddingTwips),
      line: 1,
      lineRule: LineRuleType.EXACT,
    },
    children: [new ImageRun({
      data,
      type,
      transformation: {
        width: Math.max(1, pointsToDrawingPixels(width)),
        height: Math.max(1, pointsToDrawingPixels(height)),
      },
      floating: {
        horizontalPosition: {
          relative: HorizontalPositionRelativeFrom.CHARACTER,
          offset: pointsToDrawingEmus(Math.max(0, (item.width - width) / 2)),
        },
        verticalPosition: {
          relative: VerticalPositionRelativeFrom.PARAGRAPH,
          offset: pointsToDrawingEmus(Math.max(0, (item.height - height) / 2)),
        },
        allowOverlap: true,
        lockAnchor: true,
        behindDocument: false,
        layoutInCell: true,
        margins: { top: 0, right: 0, bottom: 0, left: 0 },
        wrap: { type: TextWrappingType.NONE },
        zIndex: Math.max(1, Math.round(Number(item.zIndex || 0)) + 1),
      },
    })],
  });
}

function lineBorder(line) {
  if (!line?.line) return null;
  // A trace line is a visible primitive that the canonical PDF actually stroked. Older traces did not
  // record the stroke style, so absence means the PDF's solid stroke rather than BorderStyle=None.
  const style = line.line.style || 'Solid';
  if (/^none$/i.test(String(style))) return null;
  return {
    style,
    width: Number(line.line?.width || 1),
    color: line.line?.color || '#000000',
  };
}

function edgeMatches(item, side, box) {
  if (side === 'top' || side === 'bottom') {
    const itemY = side === 'top' ? item.y : item.y + item.height;
    const boxY = side === 'top' ? box.y : box.y + box.height;
    return Math.abs(itemY - boxY) <= GEOMETRY_EPSILON
      && item.x <= box.x + GEOMETRY_EPSILON
      && item.x + item.width >= box.x + box.width - GEOMETRY_EPSILON;
  }
  const itemX = side === 'left' ? item.x : item.x + item.width;
  const boxX = side === 'left' ? box.x : box.x + box.width;
  return Math.abs(itemX - boxX) <= GEOMETRY_EPSILON
    && item.y <= box.y + GEOMETRY_EPSILON
    && item.y + item.height >= box.y + box.height - GEOMETRY_EPSILON;
}

function lineMatches(line, side, box) {
  const horizontal = Math.abs(line.height) <= GEOMETRY_EPSILON;
  const vertical = Math.abs(line.width) <= GEOMETRY_EPSILON;
  if (!horizontal && !vertical) return false;
  if ((side === 'top' || side === 'bottom') && horizontal) {
    const y = side === 'top' ? box.y : box.y + box.height;
    return Math.abs(line.y - y) <= GEOMETRY_EPSILON
      && line.x <= box.x + GEOMETRY_EPSILON
      && line.x + line.width >= box.x + box.width - GEOMETRY_EPSILON;
  }
  if ((side === 'left' || side === 'right') && vertical) {
    const x = side === 'left' ? box.x : box.x + box.width;
    return Math.abs(line.x - x) <= GEOMETRY_EPSILON
      && line.y <= box.y + GEOMETRY_EPSILON
      && line.y + line.height >= box.y + box.height - GEOMETRY_EPSILON;
  }
  return false;
}

function lineOwnsCellSide(line, side) {
  const horizontal = Math.abs(line.height) <= GEOMETRY_EPSILON;
  const vertical = Math.abs(line.width) <= GEOMETRY_EPSILON;
  if (horizontal) {
    // A standalone horizontal line normally belongs to the cell above (or to the first row's top
    // edge at the canvas origin), avoiding competing borders in ordinary page content.
    // Word can suppress the bottom border of a thin, otherwise-empty leading footer row. A footer
    // divider is a shared native table edge, so materialize the same border on both touching cells.
    // Word resolves identical adjacent borders as one rule; this retains the canonical line at
    // fractional footer coordinates without rasterizing or changing the PDF layout.
    return side === 'bottom'
      || (side === 'top' && (Math.abs(line.y) <= GEOMETRY_EPSILON || line.region === 'footer'));
  }
  if (vertical) {
    // Apply the equivalent single-owner rule horizontally: the cell to the left owns the line, except at
    // the canvas origin where the first cell must own its left edge.
    return side === 'right' || (side === 'left' && Math.abs(line.x) <= GEOMETRY_EPSILON);
  }
  return false;
}

function lineCoincidesWithEdge(line, box) {
  const horizontal = Math.abs(line.height) <= GEOMETRY_EPSILON;
  const vertical = Math.abs(line.width) <= GEOMETRY_EPSILON;
  if (horizontal) {
    const onHorizontalEdge = Math.abs(line.y - box.y) <= GEOMETRY_EPSILON
      || Math.abs(line.y - (box.y + box.height)) <= GEOMETRY_EPSILON;
    const overlap = Math.min(line.x + line.width, box.x + box.width) - Math.max(line.x, box.x);
    return onHorizontalEdge && overlap > GEOMETRY_EPSILON;
  }
  if (vertical) {
    const onVerticalEdge = Math.abs(line.x - box.x) <= GEOMETRY_EPSILON
      || Math.abs(line.x - (box.x + box.width)) <= GEOMETRY_EPSILON;
    const overlap = Math.min(line.y + line.height, box.y + box.height) - Math.max(line.y, box.y);
    return onVerticalEdge && overlap > GEOMETRY_EPSILON;
  }
  return false;
}

function lineCrossesInterior(line, box) {
  const horizontal = Math.abs(line.height) <= GEOMETRY_EPSILON;
  const vertical = Math.abs(line.width) <= GEOMETRY_EPSILON;
  if (horizontal) {
    const insideVertically = line.y > box.y + GEOMETRY_EPSILON
      && line.y < box.y + box.height - GEOMETRY_EPSILON;
    const overlap = Math.min(line.x + line.width, box.x + box.width) - Math.max(line.x, box.x);
    return insideVertically && overlap > GEOMETRY_EPSILON;
  }
  if (vertical) {
    const insideHorizontally = line.x > box.x + GEOMETRY_EPSILON
      && line.x < box.x + box.width - GEOMETRY_EPSILON;
    const overlap = Math.min(line.y + line.height, box.y + box.height) - Math.max(line.y, box.y);
    return insideHorizontally && overlap > GEOMETRY_EPSILON;
  }
  return false;
}

function resolvedCellBorders(box, owner, decorators, lines) {
  return Object.fromEntries(['top', 'right', 'bottom', 'left'].map((side) => {
    let resolved = owner?.borders?.[side] || null;
    for (const decorator of decorators) {
      if (edgeMatches(decorator, side, box)) resolved = strongerBorder(resolved, decorator.borders?.[side]);
    }
    for (const line of lines) {
      if (lineOwnsCellSide(line, side) && lineMatches(line, side, box)) {
        resolved = strongerBorder(resolved, lineBorder(line));
      }
    }
    return [side, wordBorder(resolved)];
  }));
}

function resolvedBackground(box, owner, decorators) {
  if (owner?.backgroundColor) return owner.backgroundColor;
  const containing = decorators
    .filter((item) => item.backgroundColor && contains(item, box))
    .sort((left, right) => Number(left.zIndex || 0) - Number(right.zIndex || 0));
  return containing.at(-1)?.backgroundColor || null;
}

function isInvisibleRectangle(item) {
  return item.kind === 'rectangle'
    && !item.backgroundColor
    && !Object.values(item.borders || {}).some(Boolean);
}

function preparePageGrid(page, {
  items = page.items,
  originY = 0,
  canvasHeight = page.height,
  reserveSectionAnchor = true,
} = {}) {
  if (page.width > WORD_MAX_PAGE_POINTS + GEOMETRY_EPSILON || page.height > WORD_MAX_PAGE_POINTS + GEOMETRY_EPSILON) {
    unsupported('A PDF page exceeds Microsoft Word’s 22-by-22-inch page-size limit', {
      page: page.number,
      widthPt: page.width,
      heightPt: page.height,
      widthIn: pointsToInches(page.width),
      heightIn: pointsToInches(page.height),
      maximumIn: 22,
      exactPageLockedOutputAvailable: false,
    });
  }
  // Snap physical edges, not origins and dimensions independently. Independent rounding can move the
  // derived right/bottom edge by another quarter point and turn two coincident PDF cells into a false
  // overlap in Word (or leave a false gap). This is the same edge-coalescing semantic the PDF trace uses.
  const normalized = items.map((item) => {
    const x = snap(item.x);
    const y = snap(Number(item.y || 0) - originY);
    const right = snap(Number(item.x || 0) + Number(item.width || 0));
    const bottom = snap(Number(item.y || 0) + Number(item.height || 0) - originY);
    return {
      ...item,
      x,
      y,
      width: Math.max(0, right - x),
      height: Math.max(0, bottom - y),
    };
  }).filter((item) => (
    item.width >= 0
    && item.height >= 0
    // Layout-only RDL rectangles have already served their purpose in the canonical PDF pagination.
    // Their children are independently traced, so retaining an unpainted container would add artificial
    // Word grid boundaries and can reject an otherwise valid page when the design box spans page cuts.
    && !isInvisibleRectangle(item)
  ));
  const maximumCanvasBottom = canvasHeight - (reserveSectionAnchor ? SECTION_ANCHOR_POINTS : 0);
  // The table needs to extend only through the last painted PDF primitive. Empty space below it remains
  // ordinary page space; filling that space with exact-height blank table rows makes Word
  // push the mandatory terminal section paragraph onto a spurious blank page.
  const canvasBottom = Math.max(
    GRID_PRECISION_POINTS,
    normalized.length > 0
      ? Math.min(maximumCanvasBottom, Math.max(...normalized.map((item) => item.y + item.height)))
      : GRID_PRECISION_POINTS,
  );

  for (const item of normalized) {
    if (item.x < -GEOMETRY_EPSILON || item.y < -GEOMETRY_EPSILON
      || item.x + item.width > page.width + GEOMETRY_EPSILON
      || item.y + item.height > maximumCanvasBottom + GEOMETRY_EPSILON) {
      unsupported('A PDF item falls outside the Word page canvas', {
        page: page.number,
        item: item.itemName,
        kind: item.kind,
      });
    }
    if (wordTextDirection(item.writingMode) === null) {
      unsupported('Rotated or vertical editable text is not safely representable by the page-locked Word renderer', {
        page: page.number,
        item: item.itemName,
        writingMode: item.writingMode,
      });
    }
    if (item.kind === 'line'
      && Math.abs(item.width) > GEOMETRY_EPSILON
      && Math.abs(item.height) > GEOMETRY_EPSILON) {
      unsupported('Diagonal RDL lines are not safely representable as native Word cell borders', {
        page: page.number,
        item: item.itemName,
      });
    }
  }

  const decorators = normalized.filter((item) => item.kind === 'rectangle');
  const lines = normalized.filter((item) => item.kind === 'line');
  const candidates = normalized.filter((item) => ['textbox', 'tablixCell', 'image', 'chart'].includes(item.kind));
  // A PDF footer divider may be deliberately offset by a sub-point spacer before the first footer
  // content row. Word can discard a border on that empty spacer row. Move only to the immediately
  // following traced content edge and only inside the 0.5pt certification tolerance, so the border
  // is owned by a material Word row without changing canonical PDF geometry.
  snapFooterDividersToContentEdges(lines, candidates);
  const demoted = new Set();
  for (const candidate of candidates) {
    if (!isEmptyCell(candidate)) continue;
    if (candidates.some((other) => other !== candidate && contains(candidate, other) && positiveOverlap(candidate, other))) {
      demoted.add(candidate);
      decorators.push(candidate);
    }
  }
  const owners = candidates.filter((candidate) => !demoted.has(candidate));
  // Some valid RDL layouts intentionally let adjacent painted boxes overlap by the shared border stroke
  // or by the renderer's quarter-point edge precision. Word table cells cannot overlap, so resolve only
  // those shallow, content-free edge strips to their midpoint. Each source edge moves by no more than the
  // certified 0.5pt geometry tolerance. Full containment and genuine content crossings remain fail-closed.
  const coalescedEdges = coalesceShallowEdgeOverlaps(owners);
  moveCoalescedBorderLines(lines, coalescedEdges);
  for (let leftIndex = 0; leftIndex < owners.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < owners.length; rightIndex += 1) {
      const left = owners[leftIndex];
      const right = owners[rightIndex];
      if (positiveOverlap(left, right)) {
        unsupported('Overlapping editable PDF regions cannot be represented safely as native Word cells', {
          page: page.number,
          first: left.itemName,
          second: right.itemName,
        });
      }
    }
  }

  for (const line of lines) {
    for (const owner of owners) {
      // A line ending exactly where an adjacent item begins is a valid border junction, not an overlap.
      // Reject only a positive-length crossing through the item's interior. Collinear edge segments and
      // endpoint/corner contacts remain representable as independent native Word cell borders.
      if (lineCrossesInterior(line, owner) && !lineCoincidesWithEdge(line, owner)) {
        unsupported('An RDL line crosses editable content instead of coinciding with a cell edge', {
          page: page.number,
          line: line.itemName,
          item: owner.itemName,
          lineBounds: {
            x: line.x,
            y: line.y,
            width: line.width,
            height: line.height,
          },
          itemBounds: {
            x: owner.x,
            y: owner.y,
            width: owner.width,
            height: owner.height,
          },
        });
      }
    }
  }

  // Only owners occupy grid cells, so only their extents have to survive the collapse; a decorator or a
  // line that loses a sub-tolerance band still resolves onto the shared edge it decorates.
  const xAxis = pageGridAxis([
    0,
    page.width,
    ...normalized.flatMap((item) => [item.x, item.x + item.width]),
  ], owners.map((item) => [snap(item.x), snap(item.x + item.width)]));
  const yAxis = pageGridAxis([
    0,
    canvasBottom,
    ...normalized.flatMap((item) => [item.y, item.y + item.height]),
  ], owners.map((item) => [snap(item.y), snap(item.y + item.height)]));
  const xBoundaries = xAxis.boundaries;
  const yBoundaries = yAxis.boundaries;
  if (xBoundaries.length - 1 > WORD_MAX_TABLE_COLUMNS) {
    unsupported('The PDF page requires more than Microsoft Word’s 63 table columns', {
      page: page.number,
      columns: xBoundaries.length - 1,
    });
  }
  if (yBoundaries.length - 1 > WORD_MAX_TABLE_ROWS) {
    unsupported('The PDF page requires more Word table rows than the platform supports', {
      page: page.number,
      rows: yBoundaries.length - 1,
    });
  }

  const placements = owners.map((item) => ({
    item,
    startColumn: boundaryIndex(xAxis, item.x),
    endColumn: boundaryIndex(xAxis, item.x + item.width),
    startRow: boundaryIndex(yAxis, item.y),
    endRow: boundaryIndex(yAxis, item.y + item.height),
  }));
  for (const placement of placements) {
    // Every owner paints text, an image, or a fill, so it must keep at least one grid cell. A protected
    // span guarantees that for the collapse above; anything left here is a genuinely sub-tolerance box
    // that Word cannot show, and dropping it silently would lose report content.
    if (placement.endColumn <= placement.startColumn || placement.endRow <= placement.startRow) {
      unsupported('A PDF region is too small to occupy its own native Word grid cell', {
        page: page.number,
        item: placement.item.itemName,
        kind: placement.item.kind,
        widthPt: placement.item.width,
        heightPt: placement.item.height,
      });
    }
  }
  const coverage = Array.from(
    { length: yBoundaries.length - 1 },
    () => Array(xBoundaries.length - 1).fill(null),
  );
  for (const placement of placements) {
    for (let row = placement.startRow; row < placement.endRow; row += 1) {
      for (let column = placement.startColumn; column < placement.endColumn; column += 1) {
        if (coverage[row][column]) {
          unsupported('Two PDF regions resolve to the same native Word grid cell', {
            page: page.number,
            first: coverage[row][column].item.itemName,
            second: placement.item.itemName,
          });
        }
        coverage[row][column] = placement;
      }
    }
  }
  return {
    page,
    xBoundaries,
    yBoundaries,
    placements,
    coverage,
    decorators,
    lines,
    coalescedEdges,
  };
}

function emptyFooterParagraph() {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text: '' })],
  });
}

function footerLayout(page) {
  const region = page.regions?.footer;
  if (!region || region.height <= GEOMETRY_EPSILON) return null;
  const items = page.items.filter((item) => item.region === 'footer');
  const contentBottom = items.length > 0
    ? Math.max(...items.map((item) => Number(item.y || 0) + Number(item.height || 0)))
    : region.y + region.height;
  const height = Math.max(region.height, contentBottom - region.y);
  return {
    region,
    items,
    height,
    bottomDistance: Math.max(0, page.height - region.y - height),
  };
}

async function nativePageFooter(
  page,
  resources,
  model,
  request,
  config,
  tempDir,
  chartCounter,
  fitTextCounter,
) {
  const layout = footerLayout(page);
  if (!layout) return null;

  if (layout.items.length === 0) {
    // Every canonical page owns its footer relationship. An explicit empty footer prevents Word from
    // inheriting visible content from the previous one-page section when the RDL hides a page footer.
    return new Footer({ children: [emptyFooterParagraph()] });
  }

  const grid = preparePageGrid(page, {
    items: layout.items,
    originY: layout.region.y,
    canvasHeight: layout.height,
    reserveSectionAnchor: false,
  });
  return new Footer({
    // A footer story ending in a table makes Word synthesize a terminal paragraph whose height is
    // renderer-dependent. Materialize the required paragraph with the same one-twip exact geometry used
    // by empty grid cells so it cannot enlarge the footer band and steal space from the page body.
    children: [
      await pageTable(
        grid,
        resources,
        model,
        request,
        config,
        tempDir,
        chartCounter,
        fitTextCounter,
      ),
      emptyFooterParagraph(),
    ],
  });
}

function cellBox(grid, row, column, rowSpan = 1, columnSpan = 1) {
  return {
    x: grid.xBoundaries[column],
    y: grid.yBoundaries[row],
    width: grid.xBoundaries[column + columnSpan] - grid.xBoundaries[column],
    height: grid.yBoundaries[row + rowSpan] - grid.yBoundaries[row],
  };
}

// Word starts a cell's content area - and the background fill behind it - below the cell's top margin.
// When the cell is vertically merged and its first band is shorter than that margin, which is what any
// neighbouring item starting a fraction lower produces, the fill cannot begin in that band at all: it
// resumes in the next one, leaving the top border stranded above a strip of unfilled cell. Word's screen
// renderer draws that as a separate rule above the item with a gap beneath it. Such a margin has to move
// out of `tcMar` and into the paragraph flow, which carries the same offset without displacing the fill.
// A cell whose first band can hold its own padding - every ordinary cell - keeps the margin untouched.
function topMarginFitsFirstBand(item, firstBandTwips) {
  return pointsToTwips(item?.padding?.top || 0) <= firstBandTwips;
}

function cellMargins(item, firstBandTwips = Infinity) {
  const padding = item?.padding || {};
  return {
    top: topMarginFitsFirstBand(item, firstBandTwips)
      ? Math.max(0, pointsToTwips(padding.top || 0))
      : 0,
    right: Math.max(0, pointsToTwips(padding.right || 0)),
    // Microsoft Word adds the largest bottom cell margin to an exact row height. The canonical PDF trace
    // already includes bottom padding inside the physical cell box, so tcMar/bottom would make the Word
    // row taller. Preserve the same inner content box with trailing paragraph space instead; that space
    // participates in top/center/bottom vertical alignment without changing the exact row height.
    bottom: 0,
    left: Math.max(0, pointsToTwips(padding.left || 0)),
    marginUnitType: WidthType.DXA,
  };
}

// One traced report item can span several page-grid rows, because any other item anywhere on the page
// contributes its own edges to the shared grid. WordprocessingML expresses that as a vertical merge, and
// the merged region's rules come from its outer cells: the top from the first band, the bottom from the
// last, the sides from every band. Repeating the item's own top and bottom on the inner bands paints a
// horizontal rule *inside* the cell at each grid row it crosses. That is invisible while the bands are
// tall enough for Word to suppress it, but a band only as tall as the strokes themselves - the common
// case when a neighbouring item starts a point below this one - renders it as a second rule just under
// the real border. Distribute the horizontal rules across the merge instead of repeating them.
function mergeBandBorders(borders, band) {
  if (!band || band.count <= 1) return borders;
  return {
    ...borders,
    top: band.index === 0 ? borders.top : NONE_BORDER,
    bottom: band.index === band.count - 1 ? borders.bottom : NONE_BORDER,
  };
}

async function tableCellFor(
  grid,
  row,
  column,
  placement,
  resources,
  model,
  request,
  config,
  tempDir,
  chartCounter,
  fitTextCounter,
  band = null,
) {
  const rowSpan = placement ? placement.endRow - placement.startRow : 1;
  const columnSpan = placement ? placement.endColumn - placement.startColumn : 1;
  const owner = placement?.item || null;
  const merged = Boolean(band) && band.count > 1;
  const continuation = merged && band.index > 0;
  // Borders, background, and geometry are always resolved against the item's whole traced box, never
  // against the individual band, so a merge cannot change what the canonical PDF painted.
  const box = cellBox(grid, placement ? placement.startRow : row, column, rowSpan, columnSpan);
  if (continuation) {
    return new TableCell({
      width: { size: pointsToTwips(box.width), type: WidthType.DXA },
      columnSpan,
      verticalMerge: VerticalMergeType.CONTINUE,
      borders: mergeBandBorders(resolvedCellBorders(box, owner, grid.decorators, grid.lines), band),
      shading: (() => {
        const fill = resolvedBackground(box, owner, grid.decorators);
        return fill ? { type: ShadingType.CLEAR, fill: cleanColor(fill), color: 'auto' } : undefined;
      })(),
      children: [new Paragraph({
        spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
        children: [new TextRun({ text: '' })],
      })],
    });
  }
  const bottomPaddingTwips = Math.max(
    0,
    pointsToTwips(owner?.padding?.bottom || 0),
  );
  const firstBandTwips = placement
    ? pointsToTwips(grid.yBoundaries[placement.startRow + 1] - grid.yBoundaries[placement.startRow])
    : Infinity;
  const margins = owner ? cellMargins(owner, firstBandTwips) : {
    top: 0, right: 0, bottom: 0, left: 0, marginUnitType: WidthType.DXA,
  };
  // Whatever `cellMargins` refused to put in `tcMar/top` is carried by the content instead, so the item
  // keeps the same inner box it had in the canonical PDF.
  const displacedTopPaddingTwips = owner && !topMarginFitsFirstBand(owner, firstBandTwips)
    ? Math.max(0, pointsToTwips(owner.padding?.top || 0))
    : 0;
  let children;
  if (owner?.kind === 'image' || owner?.kind === 'chart') {
    const withPage = {
      ...owner,
      pageNumber: grid.page.number,
      totalPages: request.__canonicalPageCount,
    };
    children = [await pictureForItem(
      withPage,
      resources,
      model,
      request,
      config,
      tempDir,
      chartCounter.value++,
      bottomPaddingTwips,
      displacedTopPaddingTwips,
    )];
  } else {
    children = owner
      ? linesForParagraphs(owner, bottomPaddingTwips, fitTextCounter, displacedTopPaddingTwips)
      : [new Paragraph({
        spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
        children: [new TextRun({ text: '' })],
      })];
  }
  const background = resolvedBackground(box, owner, grid.decorators);
  return new TableCell({
    width: { size: pointsToTwips(box.width), type: WidthType.DXA },
    columnSpan,
    // The continuation bands are emitted explicitly above, so `rowSpan` must stay off: it would make the
    // docx builder append its own continuations that repeat this cell's borders on every inner band.
    verticalMerge: merged ? VerticalMergeType.RESTART : undefined,
    margins,
    verticalAlign: owner?.kind === 'image' || owner?.kind === 'chart'
      // The floating picture is positioned from this paragraph's top edge. Centering a negligible
      // anchor paragraph inside a tall owner cell moves the entire drawing down by half the cell height.
      ? VerticalAlignTable.TOP
      : verticalAlignment(owner?.verticalAlign),
    // RDL's two orthogonal writing modes have direct native WordprocessingML table-cell equivalents.
    // The canonical PDF trace has already resolved any expression-backed WritingMode, so Word consumes
    // only the physical direction painted by PDF and never re-evaluates report data independently.
    textDirection: owner ? wordTextDirection(owner.writingMode) || undefined : undefined,
    shading: background ? { type: ShadingType.CLEAR, fill: cleanColor(background), color: 'auto' } : undefined,
    borders: mergeBandBorders(resolvedCellBorders(box, owner, grid.decorators, grid.lines), band),
    children,
  });
}

async function pageTable(
  grid,
  resources,
  model,
  request,
  config,
  tempDir,
  chartCounter,
  fitTextCounter,
) {
  const rows = [];
  for (let row = 0; row < grid.yBoundaries.length - 1; row += 1) {
    const children = [];
    let column = 0;
    while (column < grid.xBoundaries.length - 1) {
      const placement = grid.coverage[row][column];
      if (placement) {
        if (placement.startColumn !== column) {
          column += 1;
          continue;
        }
        const cell = await tableCellFor(
          grid,
          row,
          column,
          placement,
          resources,
          model,
          request,
          config,
          tempDir,
          chartCounter,
          fitTextCounter,
          {
            index: row - placement.startRow,
            count: placement.endRow - placement.startRow,
          },
        );
        children.push(cell);
        column = placement.endColumn;
      } else {
        const cell = await tableCellFor(
          grid,
          row,
          column,
          null,
          resources,
          model,
          request,
          config,
          tempDir,
          chartCounter,
          fitTextCounter,
        );
        children.push(cell);
        column += 1;
      }
    }
    const tracedHeightTwips = Math.max(
      1,
      pointsToTwips(grid.yBoundaries[row + 1] - grid.yBoundaries[row]),
    );
    rows.push(new TableRow({
      cantSplit: true,
      height: {
        value: tracedHeightTwips,
        rule: HeightRule.EXACT,
      },
      children,
    }));
  }
  return new Table({
    width: { size: pointsToTwips(grid.page.width), type: WidthType.DXA },
    columnWidths: grid.xBoundaries.slice(1).map((value, index) => (
      pointsToTwips(value - grid.xBoundaries[index])
    )),
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
    indent: { size: 0, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: TableBorders.NONE,
    rows,
  });
}

function pageProperties(page, index) {
  const landscape = page.width > page.height;
  const footerDistance = footerLayout(page)?.bottomDistance || 0;
  return {
    type: index === 0 ? undefined : SectionType.NEXT_PAGE,
    page: {
      size: landscape
        ? {
          width: pointsToTwips(page.height),
          height: pointsToTwips(page.width),
          orientation: PageOrientation.LANDSCAPE,
        }
        : {
          width: pointsToTwips(page.width),
          height: pointsToTwips(page.height),
          orientation: PageOrientation.PORTRAIT,
        },
      margin: {
        top: 0,
        right: 0,
        // The canonical body grid already stops before the PDF footer. A small negative bottom margin
        // gives Word's mandatory end-of-section paragraph a non-visible flow allowance below that fixed
        // grid, instead of forcing a mathematically full final table row onto another physical page.
        // Footer placement is governed independently by the exact w:footer distance below.
        bottom: -pointsToTwips(SECTION_ANCHOR_POINTS),
        left: 0,
        header: 0,
        footer: pointsToTwips(footerDistance),
        gutter: 0,
      },
    },
  };
}

function consumedFamilies(trace) {
  return [...new Set(trace.pages.flatMap((page) => page.items.flatMap((item) => (
    (item.lines || []).flatMap((line) => (line.runs || []).map((run) => run.font?.family).filter(Boolean))
  ))))];
}

async function embeddedFontFamilies(trace, config) {
  const families = consumedFamilies(trace);
  const result = [];
  for (const family of families) {
    const files = {};
    let missing = null;
    for (const variant of VARIANTS) {
      const file = resolveFontFile(config.fontDir, family, variant.bold, variant.italic);
      if (!file) {
        missing = variant.key;
        break;
      }
      files[variant.key] = {
        file,
        data: await fs.readFile(file),
        ...editableFontEmbeddingPermission(file, family, variant.key),
      };
    }
    if (missing) {
      throw new ServiceError('FONT_MISSING', `Required font is unavailable: ${family}:${missing}`, 503, {
        family,
        variant: missing,
      });
    }
    result.push({ family, files });
  }
  return result;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function obfuscateFont(data, fontKey) {
  const guidBytes = fontKey
    .replace(/-/g, '')
    .match(/../g)
    .map((hex) => Number.parseInt(hex, 16))
    .reverse();
  const result = Buffer.from(data);
  for (let index = 0; index < Math.min(32, result.length); index += 1) {
    result[index] ^= guidBytes[index % guidBytes.length];
  }
  return result;
}

async function addFontVariants(buffer, embeddedFonts) {
  const zip = await JSZip.loadAsync(buffer);
  if (embeddedFonts.length > 0) {
    const fontTableFile = zip.file('word/fontTable.xml');
    const relationshipsFile = zip.file('word/_rels/fontTable.xml.rels');
    if (!fontTableFile || !relationshipsFile) {
      throw new ServiceError('RENDER_FAILED', 'Word font table packaging is incomplete', 500);
    }
    let fontTable = await fontTableFile.async('string');
    let relationships = await relationshipsFile.async('string');
    const existingIds = [...relationships.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1]));
    let nextRelationship = Math.max(0, ...existingIds) + 1;
    const existingTargets = [...relationships.matchAll(/Target="fonts\/font(\d+)\.odttf"/g)]
      .map((match) => Number(match[1]));
    let nextFontPart = Math.max(0, ...existingTargets) + 1;

    for (const embedded of embeddedFonts) {
      const name = escapeXml(embedded.family);
      const fontPattern = new RegExp(`(<w:font\\s+w:name="${escapeRegExp(name)}"[^>]*>)([\\s\\S]*?)(</w:font>)`);
      const matched = fontTable.match(fontPattern);
      if (!matched) {
        throw new ServiceError('RENDER_FAILED', `Embedded Word font entry is missing: ${embedded.family}`, 500);
      }
      const existingFontMarkup = matched[2].replace(/<w:embedRegular\b([^>]*)\/>/, (_element, attributes) => {
        const explicitFullFile = /\bw:subsetted=/.test(attributes)
          ? attributes.replace(/\bw:subsetted="[^"]*"/, 'w:subsetted="0"')
          : `${attributes} w:subsetted="0"`;
        return `<w:embedRegular${explicitFullFile}/>`;
      });
      const additions = [];
      for (const variant of VARIANTS.filter((candidate) => candidate.key !== 'regular')) {
        const relationshipId = `rId${nextRelationship++}`;
        const fontPart = `font${nextFontPart++}.odttf`;
        const fontKey = randomUUID().toUpperCase();
        additions.push(`<w:${variant.element} r:id="${relationshipId}" w:fontKey="{${fontKey}}" w:subsetted="0"/>`);
        relationships = relationships.replace(
          '</Relationships>',
          `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/${fontPart}"/></Relationships>`,
        );
        zip.file(`word/fonts/${fontPart}`, obfuscateFont(embedded.files[variant.key].data, fontKey));
      }
      fontTable = fontTable.replace(fontPattern, `${matched[1]}${existingFontMarkup}${additions.join('')}${matched[3]}`);
    }
    zip.file('word/fontTable.xml', fontTable);
    zip.file('word/_rels/fontTable.xml.rels', relationships);
  }

  // docx assigns wp:docPr id="1" to every independently-created ImageRun. Repeated IDs violate the
  // DrawingML non-visual-property identity contract and can make Word repair the package or suppress
  // repeated images. Normalize IDs across all document stories after construction.
  const drawingParts = Object.keys(zip.files)
    .filter((name) => /^word\/(?:document|header\d+|footer\d+)\.xml$/i.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  let nextDrawingId = 1;
  for (const name of drawingParts) {
    const file = zip.file(name);
    if (!file) continue;
    let xml = await file.async('string');
    if (name === 'word/document.xml') {
      // docx emits an 18pt document grid in every section by default. Microsoft Word applies it to the
      // otherwise-empty section paragraph after each page table, turning the intended one-twip anchor
      // into an 18pt line and pushing a near-full final row onto a new page. Page-locked output uses
      // explicit point/twip geometry throughout, so a document grid is both unnecessary and incorrect.
      xml = xml
        .replace(/<w:docGrid\b[^>]*\/>/g, '')
        .replace(
          /<w:p><w:pPr>(?=<w:sectPr\b)/g,
          '<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"/>',
        );
    }
    zip.file(name, xml.replace(
      /(<wp:docPr\b[^>]*\bid=")\d+(")/g,
      (_match, prefix, suffix) => `${prefix}${nextDrawingId++}${suffix}`,
    ));
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function writeInternalArtifacts(tempDir, pdf, trace) {
  if (!tempDir) return null;
  const pdfPath = path.join(tempDir, 'docx-canonical.pdf');
  const tracePath = path.join(tempDir, 'docx-layout-trace.json');
  try {
    await fs.writeFile(pdfPath, pdf, { mode: 0o600 });
    await fs.writeFile(tracePath, JSON.stringify(trace), { mode: 0o600 });
    return { pdfPath, tracePath };
  } catch (error) {
    await Promise.all([pdfPath, tracePath].map((file) => fs.unlink(file).catch(() => {})));
    throw error;
  }
}

async function cleanupInternalArtifacts(files) {
  if (!files) return;
  await Promise.all(Object.values(files).map((file) => fs.unlink(file).catch(() => {})));
}

export async function renderPagedEditableDocx(model, request, config, tempDir, telemetry) {
  config ||= loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
  const reportTelemetry = (phase, metrics = {}) => {
    try { telemetry?.(phase, metrics); } catch { /* Telemetry cannot affect canonical PDF or DOCX output. */ }
  };
  validateWindowsWordRequest(request);
  reportTelemetry('docx.compatibility-validated');
  const canonical = await renderPdf(model, request, config, {
    captureLayoutTrace: true,
    telemetry: (phase, metrics) => reportTelemetry(`docx.canonical-${phase}`, metrics),
  });
  const trace = canonical.layoutTrace;
  const tracedItemCount = trace.pages.reduce((sum, page) => sum + (page.items?.length || 0), 0);
  reportTelemetry('docx.canonical-pdf-completed', {
    pageCount: canonical.pageCount,
    canonicalPdfBytes: canonical.buffer.length,
    tracedItemCount,
  });
  try {
    validateLayoutTrace(trace, canonical.pageCount);
  } catch (error) {
    throw new ServiceError('RENDER_FAILED', 'Canonical PDF layout trace is incomplete', 500, {
      cause: error.message,
    });
  }
  reportTelemetry('docx.layout-trace-validated', { pageCount: canonical.pageCount, tracedItemCount });
  let ownedTempDir = null;
  let workingTempDir = tempDir;
  const requiresChartWorkspace = trace.pages.some((page) => (
    page.items.some((item) => item.kind === 'chart')
  ));
  if (!workingTempDir && requiresChartWorkspace) {
    await fs.mkdir(config.tempRoot, { recursive: true, mode: 0o700 });
    await fs.chmod(config.tempRoot, 0o700);
    ownedTempDir = await fs.mkdtemp(path.join(config.tempRoot, 'docx-chart-'));
    await fs.chmod(ownedTempDir, 0o700);
    workingTempDir = ownedTempDir;
  }
  reportTelemetry('docx.workspace-prepared', { requiresChartWorkspace, ownsWorkspace: Boolean(ownedTempDir) });
  let internalFiles = null;
  try {
    internalFiles = await writeInternalArtifacts(workingTempDir, canonical.buffer, trace);
    reportTelemetry('docx.internal-artifacts-written', { written: Boolean(internalFiles) });
    const embeddedFonts = await embeddedFontFamilies(trace, config);
    const embeddedFontBytes = embeddedFonts.reduce((familySum, embedded) => (
      familySum + Object.values(embedded.files).reduce((variantSum, variant) => variantSum + variant.data.length, 0)
    ), 0);
    reportTelemetry('docx.fonts-loaded', {
      embeddedFontFamilyCount: embeddedFonts.length,
      embeddedFontVariantCount: embeddedFonts.length * VARIANTS.length,
      embeddedFontBytes,
    });
    const resources = modelResources(model);
    const canonicalRequest = { ...request, __canonicalPageCount: canonical.pageCount };
    const chartCounter = { value: 0 };
    const fitTextCounter = { value: 1 };
    const sections = [];
    for (const [index, page] of trace.pages.entries()) {
      // Footer items must not participate in body flow. Word always inserts a terminal paragraph after
      // the section's body table; keeping footer rows in that table lets the terminal paragraph push the
      // final footer row onto a blank page. A native footer is positioned independently from body flow.
      const bodyGrid = preparePageGrid(page, {
        items: page.items.filter((item) => item.region !== 'footer'),
      });
      const footer = await nativePageFooter(
        page,
        resources,
        model,
        canonicalRequest,
        config,
        workingTempDir,
        chartCounter,
        fitTextCounter,
      );
      sections.push({
        properties: pageProperties(page, index),
        footers: footer ? { default: footer } : undefined,
        children: [await pageTable(
          bodyGrid,
          resources,
          model,
          canonicalRequest,
          config,
          workingTempDir,
          chartCounter,
          fitTextCounter,
        )],
      });
      if ((index + 1) % 25 === 0 || index + 1 === trace.pages.length) {
        reportTelemetry('docx.page-construction-progress', {
          pagesConstructed: index + 1,
          pageCount: trace.pages.length,
          chartCount: chartCounter.value,
        });
      }
    }
    reportTelemetry('docx.native-pages-constructed', {
      pageCount: sections.length,
      chartCount: chartCounter.value,
    });
    const document = new Document({
      creator: 'RDL Converter Service',
      title: request.outputFileName || model.name,
      description: 'Windows Word page-locked editable rendering derived from the canonical PDF layout trace',
      compatibilityModeVersion: 15,
      features: { updateFields: false },
      fonts: embeddedFonts.map((embedded) => ({
        name: embedded.family,
        data: embedded.files.regular.data,
      })),
      styles: {
        default: {
          document: {
            run: { font: 'Arial', size: 2 },
            paragraph: {
              spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
            },
          },
        },
      },
      sections,
    });
    reportTelemetry('docx.ooxml-pack-started', { pageCount: sections.length });
    let buffer = await Packer.toBuffer(document);
    reportTelemetry('docx.ooxml-pack-completed', { packageBytes: buffer.length });
    buffer = await addFontVariants(buffer, embeddedFonts);
    reportTelemetry('docx.font-variants-packaged', {
      packageBytes: buffer.length,
      embeddedFontFamilyCount: embeddedFonts.length,
      embeddedFontVariantCount: embeddedFonts.length * VARIANTS.length,
    });
    return {
      buffer,
      pageCount: canonical.pageCount,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      extension: 'docx',
      layoutMode: 'windows-paged-editable',
      editableTextRatio: 1,
      canonicalPdfSha256: createHash('sha256').update(canonical.buffer).digest('hex'),
    };
  } finally {
    await cleanupInternalArtifacts(internalFiles);
    if (ownedTempDir) await fs.rm(ownedTempDir, { recursive: true, force: true });
    reportTelemetry('docx.internal-artifacts-cleaned');
  }
}
