// Report-culture helpers. SSRS formats every value in the report `Language` (a BCP-47 culture): date order,
// separators, currency symbol, month/day names, and AM/PM all follow it. These helpers resolve a declared
// Language to a canonical locale and map a locale to its default currency for the .NET `C` specifier.
//
// When no Language is declared the renderers pass `null` and the formatter keeps its legacy defaults, so a
// report that does not declare a culture is unchanged.

// Region → ISO 4217 currency for the standard `C`/currency formatting. Not exhaustive; unlisted regions
// fall back to USD, matching the engine's historical default.
const REGION_CURRENCY = {
  US: 'USD', GB: 'GBP', ZA: 'ZAR', IN: 'INR', AU: 'AUD', CA: 'CAD', NZ: 'NZD', JP: 'JPY', CN: 'CNY',
  CH: 'CHF', SE: 'SEK', NO: 'NOK', DK: 'DKK', PL: 'PLN', BR: 'BRL', MX: 'MXN', RU: 'RUB', KR: 'KRW',
  SG: 'SGD', HK: 'HKD', AE: 'AED', SA: 'SAR', NG: 'NGN', KE: 'KES', IL: 'ILS', TR: 'TRY', TH: 'THB',
  DE: 'EUR', FR: 'EUR', ES: 'EUR', IT: 'EUR', NL: 'EUR', IE: 'EUR', AT: 'EUR', BE: 'EUR', PT: 'EUR',
  FI: 'EUR', GR: 'EUR', LU: 'EUR', SK: 'EUR', SI: 'EUR', EE: 'EUR', LV: 'EUR', LT: 'EUR', CY: 'EUR',
  MT: 'EUR',
};

// Canonicalize a declared Language into a usable locale, or return null when it is absent/invalid so the
// caller keeps its legacy default. Never throws.
export function canonicalizeCulture(language) {
  const raw = String(language ?? '').trim();
  if (!raw) return null;
  try {
    const [canonical] = Intl.getCanonicalLocales(raw);
    return canonical || null;
  } catch {
    return null;
  }
}

// The ISO 4217 currency code for a locale's region (en-ZA → ZAR, en-GB → GBP, de-DE → EUR). Falls back to
// USD, which is also the code used when culture is null (legacy default).
export function currencyForCulture(culture) {
  if (!culture) return 'USD';
  let region = null;
  try {
    const locale = new Intl.Locale(culture);
    region = locale.region || locale.maximize().region || null;
  } catch {
    region = null;
  }
  return (region && REGION_CURRENCY[region]) || 'USD';
}
