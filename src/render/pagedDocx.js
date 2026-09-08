import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import {
  AlignmentType,
  BorderStyle,
  BuilderElement,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
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
import { resolveReportCulture } from '../rdl/expression.js';
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
// The reflowable profile's Word default run size (half-points) and the single-spacing pitch Word gives a
// line of that size. An empty editable cell receives that pitch, capped by its row budget, so text a user
// types there is visible instead of vanishing inside the one-twip line the page-locked profile uses.
const REFLOWABLE_DEFAULT_FONT_HALF_POINTS = 20;
const WORD_SINGLE_LINE_FACTOR = 1.15;
const REFLOWABLE_EMPTY_LINE_TWIPS = Math.round(
  (REFLOWABLE_DEFAULT_FONT_HALF_POINTS / 2) * WORD_SINGLE_LINE_FACTOR * 20,
);
// Microsoft Word measures a line of embedded-font text slightly differently from the canonical PDF
// renderer. Across the traced lines of real reports the difference ranged from a few points narrower to
// about 1.5pt wider, independent of run boundaries and not proportional to the line length. The
// page-locked profile pins every line with FitText; the reflowable profile cannot, because FitText would
// squeeze text a user adds into the traced width. Instead a reflowable paragraph whose traced line has
// less horizontal slack than this allowance extends its editable text area by the shortfall, on the side
// away from its alignment, so that Word cannot wrap a line the PDF kept whole and grow the row by a line.
// Text that fits still renders at its natural width, so an unedited page is unchanged.
const WORD_LINE_MEASUREMENT_ALLOWANCE_POINTS = 3;
const WORD_LINE_MEASUREMENT_ALLOWANCE_RATIO = 0.0025;
// Word renders each horizontal rule charged to an `atLeast` row a fraction of a twip taller than its
// declared thickness (measured 0.6 to 0.9 twips per 1pt to 2pt rule over 40-row tables). A page whose
// grid fills the body area to the twip therefore overruns it by that fraction times the number of ruled
// rows, and Word answers by pushing the last row onto an overflow page. Reserve one twip per ruled edge
// when a flush grid is capped; the last row gives it up, which is invisible.
const WORD_RULE_ROUNDING_TWIPS = 1;
// Paragraph properties of every terminal paragraph Word requires after a story's last table: one exact
// twip of line pitch and a hidden paragraph mark, which Word lays out with no vertical extent.
const HIDDEN_TERMINAL_PARAGRAPH_PROPERTIES = '<w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"/>'
  + '<w:rPr><w:vanish/></w:rPr>';
const GEOMETRY_EPSILON = 0.13;
const CERTIFIED_GEOMETRY_TOLERANCE_POINTS = 0.5;
const NONE_BORDER = Object.freeze({ style: BorderStyle.NONE, size: 0, color: 'auto' });
const RESOLVED_TABLIX_FRAGMENT_BORDER = 'resolvedTablixFragmentBorder';
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
  // Images and charts always paint their declared owner even though they have no text, fill, or cell
  // border metadata. Treating them as empty lets a real drawing overlap bypass the shallow-edge cap.
  return ['textbox', 'tablixCell'].includes(item.kind)
    && !item.backgroundColor
    && !Object.values(item.borders || {}).some(Boolean);
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

function canonicalCoalescibleDrawingAxis(first, second, canonicalBounds) {
  const firstBounds = canonicalBounds?.get(first);
  const secondBounds = canonicalBounds?.get(second);
  if (!firstBounds || !secondBounds) return null;

  const left = firstBounds.x <= secondBounds.x ? firstBounds : secondBounds;
  const right = left === firstBounds ? secondBounds : firstBounds;
  const verticalSpan = Math.min(left.y + left.height, right.y + right.height)
    - Math.max(left.y, right.y);
  const horizontalEdgeOverlap = (left.x + left.width) - right.x;
  if (verticalSpan > GEOMETRY_EPSILON
    && horizontalEdgeOverlap >= -GEOMETRY_EPSILON
    && horizontalEdgeOverlap <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS + GEOMETRY_EPSILON) {
    return 'horizontal';
  }

  const upper = firstBounds.y <= secondBounds.y ? firstBounds : secondBounds;
  const lower = upper === firstBounds ? secondBounds : firstBounds;
  const horizontalSpan = Math.min(upper.x + upper.width, lower.x + lower.width)
    - Math.max(upper.x, lower.x);
  const verticalEdgeOverlap = (upper.y + upper.height) - lower.y;
  if (horizontalSpan > GEOMETRY_EPSILON
    && verticalEdgeOverlap >= -GEOMETRY_EPSILON
    && verticalEdgeOverlap <= CERTIFIED_GEOMETRY_TOLERANCE_POINTS + GEOMETRY_EPSILON) {
    return 'vertical';
  }
  return null;
}

function coalesceShallowEdgeOverlaps(items, canonicalBounds) {
  const adjustments = [];
  for (let leftIndex = 0; leftIndex < items.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < items.length; rightIndex += 1) {
      const left = items[leftIndex];
      const right = items[rightIndex];
      if (!positiveOverlap(left, right)) continue;
      const supportedKinds = ['textbox', 'tablixCell', 'image', 'chart'];
      if (!supportedKinds.includes(left.kind) || !supportedKinds.includes(right.kind)) continue;
      if (!/^(default|horizontal)?$/i.test(String(left.writingMode || 'default'))
        || !/^(default|horizontal)?$/i.test(String(right.writingMode || 'default'))) continue;

      const includesDrawing = ['image', 'chart'].includes(left.kind)
        || ['image', 'chart'].includes(right.kind);
      // A drawing has no text paint bounds with which to prove an arbitrary overlap harmless. Permit it
      // only when the unsnapped canonical PDF rectangles meet, or overlap by no more than the certified
      // 0.5pt Word geometry tolerance, on the corresponding edge. This handles both a three-decimal trace
      // edge straddling a 0.125pt Word-grid midpoint and the sub-point edge strips emitted by RDL designers,
      // without accepting a deliberately stacked logo, image, or chart.
      const canonicalAxis = includesDrawing
        ? canonicalCoalescibleDrawingAxis(left, right, canonicalBounds)
        : null;
      if (includesDrawing && !canonicalAxis) continue;

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
      const adjustment = canonicalAxis === 'horizontal'
        ? coalesceHorizontalEdge(left, right)
        : canonicalAxis === 'vertical'
          ? coalesceVerticalEdge(left, right)
          : preferVertical
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

// `w:sz` is measured in eighths of a point, and a double rule paints two strokes plus the gap between them.
function borderThicknessTwips(border) {
  if (!border || border.style === BorderStyle.NONE) return 0;
  const strokes = border.style === BorderStyle.DOUBLE ? 3 : 1;
  return Math.round((Number(border.size || 0) / 8) * 20 * strokes);
}

function bandFor(placement, row) {
  if (!placement) return null;
  return { index: row - placement.startRow, count: placement.endRow - placement.startRow };
}

// Thickest rule on the lower edge of one grid row, across every cell that row shows.
function rowBottomEdgeTwips(grid, row) {
  let edge = 0;
  let column = 0;
  while (column < grid.xBoundaries.length - 1) {
    const placement = grid.coverage[row][column];
    if (placement && placement.startColumn !== column) {
      column += 1;
      continue;
    }
    const geometry = cellGeometry(grid, row, column, placement, bandFor(placement, row));
    edge = Math.max(edge, borderThicknessTwips(geometry.borders.bottom));
    column = placement ? placement.endColumn : column + 1;
  }
  return edge;
}

// One resolved description of a page-grid cell, shared by the Word row arithmetic below and the cell
// builder so that both see the same margins and borders. Borders, background, and geometry are always
// resolved against the item's whole traced box, never against the individual band, so a merge cannot
// change what the canonical PDF painted.
function cellGeometry(grid, row, column, placement, band = null) {
  const rowSpan = placement ? placement.endRow - placement.startRow : 1;
  const columnSpan = placement ? placement.endColumn - placement.startColumn : 1;
  const owner = placement?.item || null;
  const merged = Boolean(band) && band.count > 1;
  const continuation = merged && band.index > 0;
  const box = cellBox(grid, placement ? placement.startRow : row, column, rowSpan, columnSpan);
  const borders = mergeBandBorders(resolvedCellBorders(box, owner, grid.decorators, grid.lines), band);
  const firstBandTwips = placement
    ? pointsToTwips(grid.yBoundaries[placement.startRow + 1] - grid.yBoundaries[placement.startRow])
    : Infinity;
  let margins = null;
  if (!continuation) {
    margins = owner ? cellMargins(owner, firstBandTwips) : {
      top: 0, right: 0, bottom: 0, left: 0, marginUnitType: WidthType.DXA,
    };
  }
  return {
    owner,
    columnSpan,
    merged,
    continuation,
    box,
    borders,
    margins,
    firstBandTwips,
    // `plain` covers both an unowned gap cell and a single-band item: Word measures either as ordinary
    // row content. `restart` is the first band of a vertical merge and `continue` every later band.
    kind: continuation ? 'continue' : merged ? 'restart' : 'plain',
  };
}

// Microsoft Word row arithmetic, measured in Word for Windows with synthetic tables. A row whose height
// rule is `atLeast` renders at
//
//   max(published value, tallest ordinary cell content) + largest top+bottom cell-margin pair
//   + thickest rule meeting at the horizontal edge above the row (+ the edge below it for the last row)
//
// An `exact` row renders at its published value regardless of content, margins, or borders. A vertical
// merge charges its content against the sum of the region's published values and grows only the region's
// last row; an `exact` last row therefore freezes the region. The one exception is a first band whose
// every cell starts a merge: Word then measures that band as if the merged content were ordinary content,
// unless the band is `exact`, which restores the region accounting.
//
// The reflowable profile publishes `atLeast` so that content a user adds grows the row, but an unedited
// row has to occupy exactly its canonical PDF height or the page grid drifts and Word repaginates the
// report differently from the PDF. Subtracting the margin and edge overhead from the traced height gives
// that, and the returned budgets tell each cell how much content Word will accept before it grows.
// `flowBudgetTwips` caps the grid's rendered height at the story band it must fit.
function wordRowPlan(grid, flowBudgetTwips = null) {
  const rowCount = grid.yBoundaries.length - 1;
  const rows = [];
  for (let row = 0; row < rowCount; row += 1) {
    const cells = [];
    let column = 0;
    while (column < grid.xBoundaries.length - 1) {
      const placement = grid.coverage[row][column];
      if (placement && placement.startColumn !== column) {
        column += 1;
        continue;
      }
      cells.push(cellGeometry(grid, row, column, placement, bandFor(placement, row)));
      column = placement ? placement.endColumn : column + 1;
    }
    rows.push({
      tracedTwips: Math.max(1, pointsToTwips(grid.yBoundaries[row + 1] - grid.yBoundaries[row])),
      marginTwips: Math.max(0, ...cells.map((cell) => (
        cell.margins ? Math.max(0, cell.margins.top) + Math.max(0, cell.margins.bottom) : 0
      ))),
      topEdgeTwips: Math.max(0, ...cells.map((cell) => borderThicknessTwips(cell.borders.top))),
      bottomEdgeTwips: Math.max(0, ...cells.map((cell) => borderThicknessTwips(cell.borders.bottom))),
      degenerate: cells.length > 0 && cells.every((cell) => cell.kind === 'restart'),
    });
  }
  rows.forEach((entry, row) => {
    const edgeAbove = Math.max(entry.topEdgeTwips, row > 0 ? rows[row - 1].bottomEdgeTwips : 0);
    const edgeBelow = row === rows.length - 1 ? entry.bottomEdgeTwips : 0;
    const overhead = entry.marginTwips + edgeAbove + edgeBelow;
    const atLeast = Math.floor(entry.tracedTwips - overhead);
    // A band thinner than its own rules and margins cannot be published as `atLeast` at all: Word would
    // render the overhead on top of a one-twip minimum and the page would grow by the difference.
    if (entry.degenerate || atLeast < 1) {
      entry.rule = HeightRule.EXACT;
      entry.valueTwips = entry.tracedTwips;
      entry.capacityTwips = Math.max(0, entry.tracedTwips - overhead);
      entry.overheadTwips = 0;
      entry.roundingTwips = 0;
    } else {
      entry.rule = HeightRule.ATLEAST;
      entry.valueTwips = atLeast;
      entry.capacityTwips = atLeast;
      entry.overheadTwips = overhead;
      entry.roundingTwips = (edgeAbove > 0 ? WORD_RULE_ROUNDING_TWIPS : 0)
        + (edgeBelow > 0 ? WORD_RULE_ROUNDING_TWIPS : 0);
    }
  });
  if (Number.isFinite(flowBudgetTwips)) {
    let rendered = rows.reduce((sum, entry) => (
      sum + entry.valueTwips + entry.overheadTwips + entry.roundingTwips
    ), 0);
    // Shave the excess from the bottom of the grid, never more than a row can give up; a row that is
    // already at its one-twip minimum passes the remainder to the row above it.
    for (let row = rows.length - 1; row >= 0 && rendered > flowBudgetTwips; row -= 1) {
      const entry = rows[row];
      const trim = Math.min(entry.valueTwips - 1, rendered - flowBudgetTwips);
      if (trim <= 0) continue;
      entry.valueTwips -= trim;
      entry.capacityTwips = Math.max(0, entry.capacityTwips - trim);
      rendered -= trim;
    }
  }
  return {
    rows,
    // The content Word accepts from an item before it grows a row: the item's own row for a single band,
    // the whole region for a vertical merge.
    budgetFor(row, placement) {
      const start = placement ? placement.startRow : row;
      const end = placement ? placement.endRow : row + 1;
      let budget = 0;
      for (let index = start; index < end; index += 1) budget += rows[index].capacityTwips;
      return budget;
    },
  };
}

// Shrinks a cell's paragraph spacing until Word will keep the row at its canonical height. The PDF box
// clips at its bottom edge, so the trailing padding goes first, then the displaced leading padding, and
// only then the line pitch, uniformly across every line so that all of the text stays legible.
function fitSpacingToBudget(spacing, lineCounts, budgetTwips, topPaddingTwips, bottomPaddingTwips) {
  const need = spacing.reduce((sum, entry, index) => (
    sum + entry.before + entry.after + entry.line * lineCounts[index]
  ), 0);
  let excess = need - budgetTwips;
  if (excess <= 0) return;
  const last = spacing[spacing.length - 1];
  const trailing = Math.min(excess, last.after, Math.max(0, bottomPaddingTwips));
  last.after -= trailing;
  excess -= trailing;
  if (excess <= 0) return;
  const first = spacing[0];
  const leading = Math.min(excess, first.before, Math.max(0, topPaddingTwips));
  first.before -= leading;
  excess -= leading;
  if (excess <= 0) return;
  const pitchTotal = spacing.reduce((sum, entry, index) => sum + entry.line * lineCounts[index], 0);
  const scale = pitchTotal > 0 ? Math.max(0, pitchTotal - excess) / pitchTotal : 0;
  spacing.forEach((entry) => {
    entry.line = Math.max(1, Math.floor(entry.line * scale));
  });
}

function strongerBorder(left, right) {
  if (!left) return right || null;
  if (!right) return left;
  const leftWidth = Number(left.width || 0);
  const rightWidth = Number(right.width || 0);
  return rightWidth >= leftWidth ? right : left;
}

// The extra editable width, in twips, a reflowable paragraph needs so that Word keeps each traced line
// whole (see `WORD_LINE_MEASUREMENT_ALLOWANCE_POINTS`). Zero when every line already has that slack.
function lineMeasurementIndentTwips(item, lines) {
  const inner = Number(item.width || 0)
    - Number(item.padding?.left || 0)
    - Number(item.padding?.right || 0);
  let need = 0;
  for (const line of lines) {
    const width = Number(line.width);
    if (!Number.isFinite(width) || width <= 0) continue;
    const allowance = WORD_LINE_MEASUREMENT_ALLOWANCE_POINTS + WORD_LINE_MEASUREMENT_ALLOWANCE_RATIO * width;
    need = Math.max(need, allowance - (inner - width));
  }
  return need > 0 ? Math.ceil(pointsToTwips(need)) : 0;
}

// Extends the paragraph's text area away from its alignment edge so the visible text does not move.
function measurementIndent(align, twips) {
  if (align === AlignmentType.RIGHT) return { left: -twips };
  if (align === AlignmentType.CENTER) {
    const half = Math.ceil(twips / 2);
    return { left: -half, right: -half };
  }
  return { right: -twips };
}

// `flow` is null for the page-locked profile. The reflowable profile passes the cell's Word content
// budget (see `wordRowPlan`) and the pitch an empty cell should offer a user.
function linesForParagraphs(
  item,
  bottomPaddingTwips = 0,
  fitTextCounter = null,
  topPaddingTwips = 0,
  flow = null,
) {
  const source = item.lines || [];
  if (source.length === 0) {
    const before = Math.max(0, topPaddingTwips);
    const after = Math.max(0, bottomPaddingTwips);
    return [new Paragraph({
      spacing: {
        before,
        after,
        // The page-locked profile keeps an empty cell physically negligible. The reflowable profile gives
        // it a real editable line, capped so that it still cannot grow the canonical row.
        line: flow ? Math.max(1, Math.min(flow.emptyLineTwips, flow.budgetTwips - before - after)) : 1,
        lineRule: LineRuleType.EXACT,
      },
      children: [new TextRun({ text: '' })],
    })];
  }

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

  const spacing = groups.map((group, groupIndex) => {
    const first = group.lines[0];
    const last = group.lines[group.lines.length - 1];
    return {
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
    };
  });
  if (flow) {
    fitSpacingToBudget(
      spacing,
      groups.map((group) => group.lines.length),
      flow.budgetTwips,
      topPaddingTwips,
      bottomPaddingTwips,
    );
  }

  return groups.map((group, groupIndex) => {
    const first = group.lines[0];
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
        runs.push(run.hyperlink
          ? new ExternalHyperlink({ link: run.hyperlink, children: [textRun] })
          : textRun);
      });
    });
    const align = alignment(first.alignment);
    const indentTwips = flow ? lineMeasurementIndentTwips(item, group.lines) : 0;
    return new Paragraph({
      alignment: align,
      indent: indentTwips > 0 ? measurementIndent(align, indentTwips) : undefined,
      spacing: {
        before: spacing[groupIndex].before,
        after: spacing[groupIndex].after,
        line: spacing[groupIndex].line,
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

// A bundled subreport's body is laid out inside this report's pages, so its embedded images and item
// definitions must be resolvable when a traced page is rebuilt as native Word content. Those child
// definitions hang off the Subreport items that resolved them; `model.subreports` is the parser's
// inventory of declared calls, not a map of loaded child reports.
function collectSubreportModels(items, result) {
  for (const item of items || []) {
    const child = item.type === 'Subreport' ? item.resolvedSubreport?.model : null;
    if (child && !result.includes(child)) {
      result.push(child);
      collectSubreportModels(child.body?.items, result);
      collectSubreportModels(child.page?.header?.items, result);
      collectSubreportModels(child.page?.footer?.items, result);
    }
    collectSubreportModels(item.items, result);
    for (const row of item.rows || []) {
      for (const cell of row.cells || []) collectSubreportModels(cell.items, result);
    }
  }
}

function collectModels(model, result = []) {
  result.push(model);
  collectSubreportModels(model.body?.items, result);
  collectSubreportModels(model.page?.header?.items, result);
  collectSubreportModels(model.page?.footer?.items, result);
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
  borderInsets = {},
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
    // Prefer the definition the canonical pass recorded: an item Name is unique only inside its own
    // report, so a chart drawn from a bundled subreport is either absent from the invoking report's index
    // or, worse, shadowed by a same-named chart there.
    const chart = item.chartItem || resources.items.get(item.itemName);
    if (!chart || !config || !tempDir) {
      unsupported('A PDF-traced chart cannot be materialized as a native Word picture', { item: item.itemName });
    }
    const datasets = normalizeDatasets(model, request);
    const globals = {
      PageNumber: item.pageNumber || 1,
      TotalPages: item.totalPages || 1,
      ExecutionTime: new Date(),
      variables: model.variables || {},
      culture: resolveReportCulture(model, { parameters: request.parameters || {} }),
    };
    // Prefer the series the canonical PDF pass already resolved: it was materialized in the chart's own
    // scope (a List/canvas cell or group instance sees only that instance's rows). Re-materializing here
    // from the report-level datasets draws a different chart — every category in the dataset — so Word and
    // the PDF disagreed. Fall back only for a trace written before charts carried their data.
    const chartData = item.chartData || materializeChart(chart, datasets, request.parameters || {}, globals);
    // The chart's own expressions — title, axis titles, legend title, expression-backed styles — resolve
    // in the chart's data scope, which for a canvas cell, group instance or bundled subreport is not the
    // invoking report's dataset set. Reuse the scope the canonical pass recorded so Word evaluates the
    // same captions the PDF drew; only the page globals are this page's. Fall back to the report scope
    // for a trace written before charts carried theirs.
    const chartContext = item.chartContext
      ? { ...item.chartContext, globals: { ...(item.chartContext.globals || {}), ...globals } }
      : { datasets, parameters: request.parameters || {}, globals, fields: {}, dataset: [] };
    const rendered = await renderChartPng(
      chart,
      chartData,
      config,
      tempDir,
      chartContext,
      chartIndex,
    );
    if (!rendered?.data) unsupported('A PDF-traced chart could not be rendered for native Word', {
      item: item.itemName,
    });
    data = rendered.data;
    type = 'png';
  }

  // A floating Word drawing is painted above its host cell's native borders.  SSRS paints a container
  // border after its positioned children, so an image or chart that fills a bordered Rectangle must not
  // be allowed to cover the equivalent Word border.  Constrain the drawing to the border's inner box;
  // the physical table cell and its edges remain at the canonical PDF geometry.
  const inset = {
    top: Math.max(0, Number(borderInsets.top || 0)),
    right: Math.max(0, Number(borderInsets.right || 0)),
    bottom: Math.max(0, Number(borderInsets.bottom || 0)),
    left: Math.max(0, Number(borderInsets.left || 0)),
  };
  const innerWidth = Math.max(0.01, Number(item.width || 0) - inset.left - inset.right);
  const innerHeight = Math.max(0.01, Number(item.height || 0) - inset.top - inset.bottom);
  let width = innerWidth;
  let height = innerHeight;
  const sizing = String(item.sizing || 'FitProportional');
  if (/^(Clip|AutoSize)$/i.test(sizing)) {
    unsupported(`RDL image sizing '${sizing}' is not safely representable in the page-locked editable Word contract`, {
      item: item.itemName,
    });
  }
  if (!/^Fit$/i.test(sizing)) {
    const natural = naturalImageSize(data);
    if (natural) {
      const scale = Math.min(innerWidth / natural.width, innerHeight / natural.height);
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
          offset: pointsToDrawingEmus(inset.left + Math.max(0, (innerWidth - width) / 2)),
        },
        verticalPosition: {
          relative: VerticalPositionRelativeFrom.PARAGRAPH,
          offset: pointsToDrawingEmus(inset.top + Math.max(0, (innerHeight - height) / 2)),
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

function sideCoordinate(bounds, side) {
  if (side === 'top') return bounds.y;
  if (side === 'bottom') return bounds.y + bounds.height;
  if (side === 'left') return bounds.x;
  return bounds.x + bounds.width;
}

function perpendicularOverlap(first, second, side) {
  if (side === 'top' || side === 'bottom') {
    return Math.min(first.x + first.width, second.x + second.width) - Math.max(first.x, second.x);
  }
  return Math.min(first.y + first.height, second.y + second.height) - Math.max(first.y, second.y);
}

function alignResolvedFragmentBordersToCellEdges(lines, owners, canonicalBounds) {
  for (const line of lines) {
    const side = String(line.fragmentSide || '').toLowerCase();
    if (line.traceRole !== RESOLVED_TABLIX_FRAGMENT_BORDER
      || !['top', 'right', 'bottom', 'left'].includes(side)
      || !line.tablixName) continue;

    const canonicalLine = canonicalBounds.get(line);
    if (!canonicalLine) continue;
    const lineCoordinate = sideCoordinate(canonicalLine, side);
    const matchingCells = owners.filter((owner) => {
      if (owner.kind !== 'tablixCell' || owner.tablixName !== line.tablixName) return false;
      const canonicalCell = canonicalBounds.get(owner);
      return canonicalCell
        && Math.abs(sideCoordinate(canonicalCell, side) - lineCoordinate) <= GEOMETRY_EPSILON
        && perpendicularOverlap(canonicalLine, canonicalCell, side) > GEOMETRY_EPSILON;
    });
    if (matchingCells.length === 0) continue;

    // The trace stores origins and dimensions to three decimal places. Two physically identical PDF
    // edges can therefore serialize on opposite sides of a 0.125pt Word-grid rounding midpoint (for
    // example, a cell bottom at 732.125pt and its fragment closure at 732.124pt). Only coalesce a
    // provenance-linked outer tablix border whose canonical edge already coincides with the covered
    // cells. A genuine crossing, including a mislabeled fragment line, remains on the rejection path.
    const targetEdges = matchingCells.map((cell) => sideCoordinate(cell, side));
    const target = targetEdges[0];
    if (!targetEdges.every((edge) => Math.abs(edge - target) <= GEOMETRY_EPSILON)) continue;
    const normalizedLineCoordinate = sideCoordinate(line, side);
    if (Math.abs(normalizedLineCoordinate - target)
      > CERTIFIED_GEOMETRY_TOLERANCE_POINTS + GEOMETRY_EPSILON) continue;
    if (side === 'top' || side === 'bottom') line.y = target;
    else line.x = target;
    // The canonical PDF has already proved that this fragment rule closes these exact cell edges. Keep
    // that association with the owner instead of depending solely on a second match after Word-grid
    // rounding. A split tablix can end on a fractional edge that is represented by a vertical merge in
    // Word; if that later line-to-band comparison misses by a snap interval, Word receives no bottom
    // border and leaves an open table corner despite the canonical PDF's closing rule.
    for (const cell of matchingCells) {
      cell.fragmentBorders = {
        ...cell.fragmentBorders,
        [side]: strongerBorder(cell.fragmentBorders?.[side], lineBorder(line)),
      };
    }
  }
}

function resolvedRenderedBorders(box, owner, decorators, lines) {
  return Object.fromEntries(['top', 'right', 'bottom', 'left'].map((side) => {
    let resolved = owner?.borders?.[side] || null;
    resolved = strongerBorder(resolved, owner?.fragmentBorders?.[side]);
    for (const decorator of decorators) {
      if (edgeMatches(decorator, side, box)) resolved = strongerBorder(resolved, decorator.borders?.[side]);
    }
    for (const line of lines) {
      if (lineOwnsCellSide(line, side) && lineMatches(line, side, box)) {
        resolved = strongerBorder(resolved, lineBorder(line));
      }
    }
    return [side, resolved];
  }));
}

function resolvedCellBorders(box, owner, decorators, lines) {
  return Object.fromEntries(Object.entries(
    resolvedRenderedBorders(box, owner, decorators, lines),
  ).map(([side, border]) => [side, wordBorder(border)]));
}

function drawingBorderInsets(box, owner, decorators, lines) {
  const borders = resolvedRenderedBorders(box, owner, decorators, lines);
  return Object.fromEntries(Object.entries(borders).map(([side, border]) => [
    side,
    border && !/^none$/i.test(String(border.style || 'None'))
      ? Math.max(0, Number(border.width || 0))
      : 0,
  ]));
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

// The grid canvas is the story's text area: `originX`/`originY` locate it on the physical page and the
// RDL page margins become real Word section margins around it, so Word's print pre-check sees the
// margins the report declares. Content may still extend into the right margin (Word draws a wider
// table past the margin), but never off the page.
function preparePageGrid(page, {
  items = page.items,
  originX = 0,
  originY = 0,
  canvasWidth = page.width,
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
  const canonicalBounds = new WeakMap();
  // Horizontal edges snap on the physical page grid and only then move to the story origin, which is
  // itself snapped, so every column boundary and coalescing decision is exactly what a page-origin grid
  // produces: the text area merely starts at the section margin instead of at the page edge.
  const gridOriginX = snap(originX);
  const normalized = items.map((item) => {
    const canonical = {
      x: Number(item.x || 0) - gridOriginX,
      y: Number(item.y || 0) - originY,
      width: Number(item.width || 0),
      height: Number(item.height || 0),
    };
    const x = snap(Number(item.x || 0)) - gridOriginX;
    const y = snap(canonical.y);
    const right = snap(Number(item.x || 0) + canonical.width) - gridOriginX;
    const bottom = snap(canonical.y + canonical.height);
    const normalizedItem = {
      ...item,
      x,
      y,
      width: Math.max(0, right - x),
      height: Math.max(0, bottom - y),
      // Text paint bounds are compared with the item box, so traced line and run abscissae move to the
      // same story origin as the box.
      lines: gridOriginX === 0 ? item.lines : (item.lines || []).map((line) => ({
        ...line,
        x: line.x === undefined ? line.x : Number(line.x) - gridOriginX,
        runs: (line.runs || []).map((run) => ({
          ...run,
          x: run.x === undefined ? run.x : Number(run.x) - gridOriginX,
        })),
      })),
    };
    canonicalBounds.set(normalizedItem, canonical);
    return normalizedItem;
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
      || item.x + item.width > snap(page.width) - gridOriginX + GEOMETRY_EPSILON
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
  const coalescedEdges = coalesceShallowEdgeOverlaps(owners, canonicalBounds);
  moveCoalescedBorderLines(lines, coalescedEdges);
  alignResolvedFragmentBordersToCellEdges(lines, owners, canonicalBounds);
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
    snap(originX + canvasWidth) - gridOriginX,
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

// Word requires a paragraph after the last table of every story. A hidden one takes no vertical space,
// so a story table that closes flush on its band cannot push the body or spill onto a blank page. The
// paragraph mark itself is hidden by `finalizePackage`, which is where the section paragraphs live.
function emptyStoryParagraph() {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text: '', vanish: true })],
  });
}

function headerLayout(page) {
  const region = page.regions?.header;
  if (!region || region.height <= GEOMETRY_EPSILON) return null;
  // A tablix fragment edge exactly on the body/header boundary is classified as a header primitive by
  // the trace's inclusive region test, but it still belongs to the body tablix.  The body grid retains
  // these provenance-marked edges so it can close the correct cell.  Emitting one in the native header
  // story as well gives Word a second, detached rule immediately below the page header.
  const items = page.items.filter((item) => (
    item.region === 'header' && item.traceRole !== RESOLVED_TABLIX_FRAGMENT_BORDER
  ));
  const contentBottom = items.length > 0
    ? Math.max(...items.map((item) => Number(item.y || 0) + Number(item.height || 0)))
    : region.y + region.height;
  const height = Math.max(region.height, contentBottom - region.y);
  return {
    region,
    items,
    height,
    topDistance: Math.max(0, region.y),
  };
}

function footerLayout(page) {
  const region = page.regions?.footer;
  if (!region || region.height <= GEOMETRY_EPSILON) return null;
  const items = page.items.filter((item) => (
    item.region === 'footer' && item.traceRole !== RESOLVED_TABLIX_FRAGMENT_BORDER
  ));
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
    return new Footer({ children: [emptyStoryParagraph()] });
  }

  const grid = preparePageGrid(page, {
    items: layout.items,
    originX: layout.region.x,
    originY: layout.region.y,
    canvasWidth: layout.region.width,
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
      emptyStoryParagraph(),
    ],
  });
}

async function nativePageHeader(
  page,
  resources,
  model,
  request,
  config,
  tempDir,
  chartCounter,
  fitTextCounter,
) {
  const layout = headerLayout(page);
  if (!layout) return null;

  if (layout.items.length === 0) {
    // Each canonical page owns its header relationship. An explicit empty header prevents Word from
    // inheriting visible content from a preceding one-page section when the RDL suppresses a header.
    return new Header({ children: [emptyStoryParagraph()] });
  }

  const grid = preparePageGrid(page, {
    items: layout.items,
    originX: layout.region.x,
    originY: layout.region.y,
    canvasWidth: layout.region.width,
    canvasHeight: layout.height,
    reserveSectionAnchor: false,
  });
  return new Header({
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
      emptyStoryParagraph(),
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
  flow = null,
) {
  const {
    owner, columnSpan, merged, continuation, box, borders: cellBorders, margins, firstBandTwips,
  } = cellGeometry(grid, row, column, placement, band);
  if (continuation) {
    return new TableCell({
      width: { size: pointsToTwips(box.width), type: WidthType.DXA },
      columnSpan,
      verticalMerge: VerticalMergeType.CONTINUE,
      borders: cellBorders,
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
      // The drawing floats, so only its one-twip anchor paragraph and the paddings it carries count as
      // Word row content; keep those inside the row budget as well.
      flow
        ? Math.max(0, Math.min(bottomPaddingTwips, flow.budgetTwips - displacedTopPaddingTwips - 1))
        : bottomPaddingTwips,
      displacedTopPaddingTwips,
      drawingBorderInsets(box, owner, grid.decorators, grid.lines),
    )];
  } else {
    children = owner
      ? linesForParagraphs(owner, bottomPaddingTwips, fitTextCounter, displacedTopPaddingTwips, flow)
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
    borders: cellBorders,
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
  flowBudgetTwips = null,
) {
  // The page-locked profile publishes every traced band as an exact Word row. The reflowable profile
  // publishes Word's own row arithmetic instead, so that an unedited row still renders at its traced
  // height while content a user adds can grow it (see `wordRowPlan`).
  const plan = request.__docxReflowable === true ? wordRowPlan(grid, flowBudgetTwips) : null;
  // Exact rows are published in whole twips, so a grid that closes flush on its band can sum to a twip
  // or two more than the band itself. Word has no slack left with real margins, so that rounding excess
  // comes out of the last row (never more than the row can give).
  const lastRowIndex = grid.yBoundaries.length - 2;
  let pageLockedRoundingTrim = 0;
  if (!plan && Number.isFinite(flowBudgetTwips)) {
    const total = grid.yBoundaries.slice(1).reduce((sum, value, index) => (
      sum + Math.max(1, pointsToTwips(value - grid.yBoundaries[index]))
    ), 0);
    pageLockedRoundingTrim = Math.max(0, total - flowBudgetTwips);
  }
  const rows = [];
  for (let row = 0; row < grid.yBoundaries.length - 1; row += 1) {
    const children = [];
    let column = 0;
    while (column < grid.xBoundaries.length - 1) {
      const placement = grid.coverage[row][column];
      if (placement && placement.startColumn !== column) {
        column += 1;
        continue;
      }
      const flow = plan && placement
        ? { budgetTwips: plan.budgetFor(row, placement), emptyLineTwips: REFLOWABLE_EMPTY_LINE_TWIPS }
        : null;
      children.push(await tableCellFor(
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
        bandFor(placement, row),
        flow,
      ));
      column = placement ? placement.endColumn : column + 1;
    }
    const tracedHeightTwips = Math.max(
      1,
      pointsToTwips(grid.yBoundaries[row + 1] - grid.yBoundaries[row]),
    );
    const planned = plan?.rows[row];
    // Word draws the bottom rule of a table's last row below that row, outside its exact height (top
    // and inner rules add nothing). The page-locked last row therefore gives up that thickness so the
    // story renders at exactly its traced height and the rule lands on the traced edge; the
    // reflowable plan already carries the same edge as the last row's overhead.
    const lastRowEdgeTwips = !planned && row === lastRowIndex
      ? rowBottomEdgeTwips(grid, row) + pageLockedRoundingTrim
      : 0;
    rows.push(new TableRow({
      // Reflowable Word intentionally lets a long edited row continue on the next physical page.
      // Page-locked DOCX retains the canonical no-split page fragment behavior.
      cantSplit: !planned || planned.rule === HeightRule.EXACT,
      height: planned
        ? { value: planned.valueTwips, rule: planned.rule }
        : { value: Math.max(1, tracedHeightTwips - lastRowEdgeTwips), rule: HeightRule.EXACT },
      children,
    }));
  }
  const columnWidths = grid.xBoundaries.slice(1).map((value, index) => (
    pointsToTwips(value - grid.xBoundaries[index])
  ));
  return new Table({
    width: { size: columnWidths.reduce((sum, width) => sum + width, 0), type: WidthType.DXA },
    columnWidths,
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
    indent: { size: 0, type: WidthType.DXA },
    // Both profiles keep the canonical column grid fixed. Autofit would let Word re-measure every column
    // from its content and redraw the report's geometry on open; vertical reflow of edited text happens
    // inside the fixed columns regardless, which is the only growth the reflowable profile promises.
    layout: TableLayoutType.FIXED,
    borders: TableBorders.NONE,
    rows,
  });
}

function pageProperties(page, index) {
  const landscape = page.width > page.height;
  const body = page.regions?.body || {};
  const bodyTop = Number(body.y || 0);
  // Snapped like the grid origin (see `preparePageGrid`), so margin plus grid reproduces the exact
  // page-origin positions the page-locked profile is certified against.
  const bodyLeft = snap(Number(body.x || 0));
  const bodyRight = Math.max(
    0,
    page.width - snap(Number(body.x || 0) + Number(body.width ?? page.width - Number(body.x || 0))),
  );
  const headerDistance = headerLayout(page)?.topDistance || 0;
  const footerDistance = footerLayout(page)?.bottomDistance || 0;
  // The canonical body band ends where the PDF stopped placing body content. Everything below it - the
  // page footer band and the RDL bottom margin - is a real Word bottom margin, so the footer story
  // starts exactly on the body boundary and grown reflowable content cannot flow through it. Word's
  // terminal paragraphs are hidden (see `finalizePackage`), so a grid closing flush on that boundary
  // needs no extra room; the page-locked last row gives up only the bottom rule Word draws below it.
  const bottomMargin = Math.max(
    0,
    page.height - Number(page.bodyBottom ?? page.regions?.footer?.y ?? page.height),
  );
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
        // Body flow begins at the canonical body band. Keeping it at the physical page origin places
        // an opaque Word body table over the native header story, so headers appear to disappear even
        // though their relationships exist in the package.
        top: pointsToTwips(bodyTop),
        // The RDL left/right margins are real Word margins; every story grid starts at the text area.
        right: pointsToTwips(bodyRight),
        bottom: pointsToTwips(bottomMargin),
        left: pointsToTwips(bodyLeft),
        // The RDL PageHeader height controls the native header grid. Word's Header from Top is the
        // traced header band's physical offset from the page edge, which is the RDL top margin.
        header: pointsToTwips(headerDistance),
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
      // The section paragraph's mark is hidden as well: a hidden paragraph takes no vertical space, so a
      // page grid closing flush on the body boundary keeps its page. The last section stores its
      // properties on the body, so it receives an explicit hidden paragraph after its page table.
      xml = xml
        .replace(/<w:docGrid\b[^>]*\/>/g, '')
        .replace(
          /<w:p><w:pPr>(?=<w:sectPr\b)/g,
          `<w:p><w:pPr>${HIDDEN_TERMINAL_PARAGRAPH_PROPERTIES}`,
        )
        .replace(
          /<\/w:tbl>(?=<w:sectPr\b)/,
          `</w:tbl><w:p><w:pPr>${HIDDEN_TERMINAL_PARAGRAPH_PROPERTIES}</w:pPr></w:p>`,
        );
    } else {
      // Header and footer stories end in `emptyStoryParagraph`; hide its mark the same way so a story
      // table filling its band does not grow the story by one line and displace the body.
      xml = xml.replace(
        /<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"\/><\/w:pPr>(<w:r><w:rPr><w:vanish\/><\/w:rPr>)/g,
        `<w:p><w:pPr>${HIDDEN_TERMINAL_PARAGRAPH_PROPERTIES}</w:pPr>$1`,
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

export async function renderPagedEditableDocx(model, request, config, tempDir, telemetry, {
  reflowable = false,
} = {}) {
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
    const canonicalRequest = {
      ...request,
      __canonicalPageCount: canonical.pageCount,
      __docxReflowable: reflowable,
    };
    const chartCounter = { value: 0 };
    // FitText is a page-locking mechanism: Word compresses a run to its traced PDF width. It must not
    // be emitted for the SSRS-style profile, where manual text must preserve its normal font size.
    const fitTextCounter = reflowable ? null : { value: 1 };
    const sections = [];
    for (const [index, page] of trace.pages.entries()) {
      // Header/footer items must not participate in body flow. Word always inserts a terminal paragraph after
      // the section's body table; keeping footer rows in that table lets the terminal paragraph push the
      // final footer row onto a blank page. A native footer is positioned independently from body flow.
      const bodyGrid = preparePageGrid(page, {
        // A tablix closure exactly on the printable body/footer boundary is recorded at that coordinate.
        // Region classification places boundary primitives in the footer, but the edge belongs to the
        // last body cell. Keep it in the body grid so its closing rule is emitted on that cell rather than
        // being detached into the footer story; all ordinary footer items remain independent.
        items: page.items.filter((item) => (
          (item.region !== 'footer' && item.region !== 'header')
            || item.traceRole === RESOLVED_TABLIX_FRAGMENT_BORDER
        )),
        // The native body table flows inside Word's body area. Its coordinates are therefore relative
        // to the canonical body band, while the section top margin restores the absolute PDF position.
        originX: Number(page.regions?.body?.x || 0),
        originY: Number(page.regions?.body?.y || 0),
        canvasWidth: Number(page.regions?.body?.width || page.width),
        canvasHeight: page.height - Number(page.regions?.body?.y || 0),
      });
      const header = await nativePageHeader(
        page,
        resources,
        model,
        canonicalRequest,
        config,
        workingTempDir,
        chartCounter,
        fitTextCounter,
      );
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
        headers: header ? { default: header } : undefined,
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
          pointsToTwips(
            Number(page.bodyBottom ?? page.regions?.footer?.y ?? page.height)
              - Number(page.regions?.body?.y || 0),
          ),
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
      description: reflowable
        ? 'Windows Word reflowable editable rendering derived from the canonical PDF layout trace'
        : 'Windows Word page-locked editable rendering derived from the canonical PDF layout trace',
      compatibilityModeVersion: 15,
      features: { updateFields: false },
      fonts: embeddedFonts.map((embedded) => ({
        name: embedded.family,
        data: embedded.files.regular.data,
      })),
      styles: {
        default: {
          document: {
            // A one-point default is necessary to hide empty cells in the page-locked grid, but it makes
            // user-entered text tiny. The reflowable profile uses a normal 10pt Word default instead.
            run: { font: 'Arial', size: reflowable ? REFLOWABLE_DEFAULT_FONT_HALF_POINTS : 2 },
            paragraph: reflowable
              ? { spacing: { before: 0, after: 0 } }
              : { spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT } },
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
      layoutMode: reflowable ? 'windows-reflowable-editable' : 'windows-paged-editable',
      editableTextRatio: 1,
      canonicalPdfSha256: createHash('sha256').update(canonical.buffer).digest('hex'),
    };
  } finally {
    await cleanupInternalArtifacts(internalFiles);
    if (ownedTempDir) await fs.rm(ownedTempDir, { recursive: true, force: true });
    reportTelemetry('docx.internal-artifacts-cleaned');
  }
}
