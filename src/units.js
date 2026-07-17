import { ServiceError } from './errors.js';

const POINTS_PER_INCH = 72;
const UNIT_FACTORS = {
  in: POINTS_PER_INCH,
  cm: POINTS_PER_INCH / 2.54,
  mm: POINTS_PER_INCH / 25.4,
  pt: 1,
  pc: 12,
  px: POINTS_PER_INCH / 96,
};

export function toPoints(value, fallback = 0) {
  if (value === undefined || value === null || value === '') return fallback;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const match = String(value).trim().match(/^(-?\d+(?:\.\d+)?)\s*(in|cm|mm|pt|pc|px)$/i);
  if (!match) throw new ServiceError('RDL_INVALID', `Invalid RDL size: ${String(value).slice(0, 64)}`);
  return Number(match[1]) * UNIT_FACTORS[match[2].toLowerCase()];
}

export function pointsToTwips(points) {
  return Math.round(points * 20);
}

export function pointsToDisplayPixels(points) {
  return Math.round((points / 72) * 96);
}
