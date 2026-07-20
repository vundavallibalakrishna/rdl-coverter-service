// Recurrence guard for the "RDL property can be an expression" bug class. Every RDL Style property can be an
// =expression evaluated per row, so a renderer must NEVER regex/compare/do arithmetic on a style property's
// RAW value — it must resolve it first via styleValue / styleColor / styleSize / isHidden. This test scans
// the renderer sources for the specific raw-access signatures we removed; reintroducing any fails the build.
// It is intentionally source-level (not behavioural) so a NEW raw consumption anywhere is caught, not only
// the ones with a dedicated regression test.
import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';

const renderers = ['pdf.js', 'docx.js', 'excel.js'];
const sources = Object.fromEntries(await Promise.all(renderers.map(async (name) => [
  name, await fs.readFile(new URL(`../src/render/${name}`, import.meta.url), 'utf8'),
])));

// Each rule: a regex that must NOT match, and why. Written to match the exact defective form, so a correct
// resolved form (e.g. `styleValue(style.verticalAlign, ...)`) does not trip it.
const FORBIDDEN = [
  { re: /embeddedImages\s*\??\.?\s*\[\s*item\.value\s*\]/, why: 'raw Image Value used as embeddedImages key — resolve with styleValue first' },
  { re: /\.test\(\s*(?:item\.)?style\s*\??\.\s*verticalAlign\s*\)/, why: 'raw verticalAlign regex-tested — resolve with styleValue first' },
  { re: /\.test\(\s*(?:item\.)?style\s*\??\.\s*textAlign\s*\)/, why: 'raw textAlign regex-tested — resolve with styleValue first' },
  { re: /Number\(\s*styleValue\(\s*(?:item\.)?style\s*\??\.\s*fontSize/, why: 'fontSize read via Number(styleValue()) — use styleSize (NaN on "14pt")' },
  { re: /\bcolor\(\s*item\.style\s*\??\.\s*border\s*\??\.\s*color\s*\)/, why: 'Line/shape border colour via color() — use styleColor (expression -> fallback)' },
  { re: /pointToEmu\(\s*border\s*\??\.\s*width\s*\|\|/, why: 'border width via pointToEmu(border.width) — use styleSize (NaN on expression)' },
  { re: /Math\.max\([^,]*,\s*item\.style\s*\??\.\s*border\s*\??\.\s*width\s*\|\|/, why: 'Line border width via Math.max(..) — use styleSize' },
  // Raw padding arithmetic (subtraction/pointsToTwips/pointToEmu) on style.paddingX — must go through styleSize.
  { re: /pointsToTwips\(\s*style\.padding(?:Top|Right|Bottom|Left)\s*\)/, why: 'padding via pointsToTwips(style.paddingX) — use styleSize' },
  { re: /-\s*textbox\.style\.padding(?:Left|Right|Top|Bottom)\b/, why: 'padding arithmetic on raw textbox.style.paddingX — use styleSize' },
];

for (const name of renderers) {
  test(`${name} resolves expression-capable style properties (no raw consumption)`, () => {
    for (const { re, why } of FORBIDDEN) {
      const match = re.exec(sources[name]);
      assert.equal(match, null, `${name}: forbidden raw style access — ${why}\n  matched: ${match?.[0]}`);
    }
  });
}
