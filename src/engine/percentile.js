/**
 * percentile.js — CEI Buying & CEI Profit Percentile Tier Engine
 *
 * Reads "CEI Buying" and "CEI Profit" sheets from the loaded workbook
 * and segments each supplier into one of four tiers based on the
 * Cumulative % column:
 *
 *   Tier 1 → top80      (cumulative < 80%)   — top 80% of value
 *   Tier 2 → mid80_95   (80% ≤ cum < 95%)    — 80–95%
 *   Tier 3 → tail95     (cum ≥ 95%)           — 95–100%
 *   Tier 4 → noValue    (no value for year)   — supplier has no data
 *
 * Within each tier, suppliers are further grouped by value magnitude:
 *   millions    → value ≥ 1,000,000
 *   thousands   → 1,000 ≤ value < 1,000,000
 *   subThousand → value < 1,000
 *
 * Each supplier entry includes KPI pillar scores (from the Results sheet
 * scoresMap) when scoresMap is provided:
 *   turnoverAndMargin (30%), assortmentAndInnovation (30%),
 *   qualityAssurance (25%), fulfillmentOperations (15%),
 *   termsAndConditions, totalPerformanceEvaluation,
 *   businessClass, performance, overallClass
 *
 * The result is split by year (2024 and 2025) for both sheets.
 *
 * Sheet column layouts (0-indexed, first data row = index 1):
 *   CEI Buying  : [0]=VendorNo, [1]=VendorName, [2]=2024 Buy (EUR), [3]=2025 Buy (EUR), [4]=Cumulative%
 *   CEI Profit  : [0]=VendorNo, [1]=VendorName, [2]=2024 Profit,    [3]=2025 Profit,    [4]=Cumulative%
 *
 * ✅ ZDR compliant — reads only from the pre-loaded workbook/scoresMap in RAM.
 * No disk writes. No file paths required.
 */

const XLSX = require('xlsx');

// ── Tier classification labels ────────────────────────────────────────────────
const TIER_LABELS = {
  top80:    { label: 'Top 80%',    range: '0% – 80%',   threshold: { min: 0,    max: 0.80 } },
  mid80_95: { label: '80% – 95%', range: '80% – 95%',  threshold: { min: 0.80, max: 0.95 } },
  tail95:   { label: '95% – 100%',range: '95% – 100%', threshold: { min: 0.95, max: 1.00 } },
  noValue:  { label: 'No Value',  range: 'N/A',         threshold: null },
};

// ── Value-range group labels ──────────────────────────────────────────────────
const VALUE_GROUPS = {
  millions:    { label: '≥ 1 Million',  range: '≥ 1,000,000' },
  thousands:   { label: '1K – 1M',     range: '1,000 – 999,999' },
  subThousand: { label: 'Below 1K',    range: '< 1,000' },
};

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/** Safely parse a money/value number (rounded to 2dp for display) */
function safeNum(v) {
  if (v == null || v === '' || isNaN(Number(v))) return null;
  return parseFloat(Number(v).toFixed(2));
}

/**
 * Read cumulative % at FULL float precision — no rounding.
 * Rounding (e.g. toFixed(2)) turns 0.7975... into 0.80 which
 * breaks the strict < 0.80 tier boundary and misclassifies suppliers.
 */
function rawPct(v) {
  if (v == null || v === '' || isNaN(Number(v))) return null;
  return Number(v); // keep IEEE-754 double precision
}

/**
 * Classify a supplier into a tier based on its cumulative %.
 * cumulativePct is the RUNNING total (0–1 scale, full precision).
 *
 * Boundaries (strict):
 *   top80    → cum < 0.80           (strictly below 80%)
 *   mid80_95 → 0.80 ≤ cum < 0.95   (80% up to but not including 95%)
 *   tail95   → cum ≥ 0.95           (95% and above)
 */
function classifyTier(cumulativePct) {
  if (cumulativePct == null) return 'noValue';
  if (cumulativePct < 0.80)  return 'top80';
  if (cumulativePct < 0.95)  return 'mid80_95';
  return 'tail95';
}

/** Classify a value into a magnitude group key */
function classifyValueGroup(value) {
  if (value == null)        return null;
  if (value >= 1_000_000)   return 'millions';
  if (value >= 1_000)       return 'thousands';
  return 'subThousand';
}

/**
 * Extract KPI pillar scores for a vendor from the scoresMap.
 * Returns null if scoresMap is not provided or vendor not found.
 *
 * Pillar weights (as shown in the Performance Evaluation sheet):
 *   Turnover & Margin      → 30%
 *   Assortment & Innovation → 30%
 *   Quality Assurance       → 25%
 *   Fulfillment Operations  → 15%
 *   Terms & Conditions      → unweighted display column
 *   Total Performance Eval  → weighted sum of all pillars
 */
function extractKpi(vendorNo, scoresMap) {
  if (!scoresMap) return null;
  const s = scoresMap[vendorNo];
  if (!s) return null;
  return {
    turnoverAndMargin:        s.turnoverMargin       ?? null,  // weight 30%
    assortmentAndInnovation:  s.assortmentInnovation ?? null,  // weight 30%
    qualityAssurance:         s.quality              ?? null,  // weight 25%
    fulfillmentOperations:    s.fulfillment          ?? null,  // weight 15%
    termsAndConditions:       s.terms                ?? null,
    totalPerformanceEvaluation: s.total              ?? null,
    businessClass:            s.businessClass        ?? null,
    performance:              s.performance          ?? null,
    overallClass:             s.overallClass         ?? null,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Entry builders
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Create a canonical supplier entry object.
 *
 * @param {object} s          - raw supplier record { vendorNo, vendorName, ... }
 * @param {string} valueField - 'value2024' | 'value2025'
 * @param {number|null} value - the actual money value
 * @param {number|null} rawCum - full-precision 0–1 cumulative % for classification
 * @param {object|null} scoresMap - Results sheet scores lookup (optional)
 */
function makeEntry(s, valueField, value, rawCum, scoresMap) {
  const tier = rawCum != null ? classifyTier(rawCum) : null;
  return {
    vendorNo:      s.vendorNo,
    vendorName:    s.vendorName,
    value,
    // Display: multiplied by 100, rounded to 4dp  e.g. 79.7534
    cumulativePct: rawCum != null ? parseFloat((rawCum * 100).toFixed(4)) : null,
    kpi:           extractKpi(s.vendorNo, scoresMap),
    _tier:         tier,
    _valueGroup:   classifyValueGroup(value),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Tier builders
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Build tier buckets using the pre-computed cumulative % from Excel.
 * Used for the 2025 year (sheet is already sorted + has cumulative col).
 */
function buildTiersFromCumulative(suppliers, valueField, cumField, scoresMap) {
  const withValue = [];
  const noValue   = [];

  for (const s of suppliers) {
    const val = s[valueField];
    const cum = s[cumField];
    if (val == null || val === 0) {
      noValue.push(makeEntry(s, valueField, null, null, scoresMap));
    } else {
      withValue.push(makeEntry(s, valueField, val, cum, scoresMap));
    }
  }

  // Sort by value desc (should already be pre-sorted in sheet, but be safe)
  withValue.sort((a, b) => b.value - a.value);

  const top80    = withValue.filter(s => s._tier === 'top80');
  const mid80_95 = withValue.filter(s => s._tier === 'mid80_95');
  const tail95   = withValue.filter(s => s._tier === 'tail95');

  return buildOutput(top80, mid80_95, tail95, noValue, withValue);
}

/**
 * Build tier buckets by computing cumulative % from scratch.
 * Used for 2024 (no pre-computed cumulative % column in Excel).
 */
function buildTiersComputed(suppliers, valueField, scoresMap) {
  const withValue = [];
  const noValue   = [];

  for (const s of suppliers) {
    const val = s[valueField];
    if (val == null || val === 0) {
      noValue.push(makeEntry(s, valueField, null, null, scoresMap));
    } else {
      withValue.push(makeEntry(s, valueField, val, null, scoresMap));
    }
  }

  // Sort descending by value
  withValue.sort((a, b) => b.value - a.value);

  // Compute cumulative % — keep full IEEE-754 precision for classifyTier()
  const total = withValue.reduce((sum, s) => sum + s.value, 0);
  let running = 0;
  for (const s of withValue) {
    running        += s.value;
    const rawCum   = total > 0 ? (running / total) : null;
    s.cumulativePct = rawCum != null ? parseFloat((rawCum * 100).toFixed(4)) : null;
    s._tier        = classifyTier(rawCum);
  }

  const top80    = withValue.filter(s => s._tier === 'top80');
  const mid80_95 = withValue.filter(s => s._tier === 'mid80_95');
  const tail95   = withValue.filter(s => s._tier === 'tail95');

  return buildOutput(top80, mid80_95, tail95, noValue, withValue);
}

// ─────────────────────────────────────────────────────────────────────────────
// Output assembler
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Group a ranked supplier list by value magnitude.
 * Returns an object with three buckets: millions, thousands, subThousand.
 */
function groupByValueRange(rankedSuppliers) {
  const millions    = [];
  const thousands   = [];
  const subThousand = [];

  for (const s of rankedSuppliers) {
    switch (s._valueGroup) {
      case 'millions':    millions.push(s);    break;
      case 'thousands':   thousands.push(s);   break;
      default:            subThousand.push(s); break;
    }
  }

  const strip = (arr) => arr.map(({ _tier, _valueGroup, ...rest }) => rest);

  return {
    millions:    { ...VALUE_GROUPS.millions,    count: millions.length,    suppliers: strip(millions)    },
    thousands:   { ...VALUE_GROUPS.thousands,   count: thousands.length,   suppliers: strip(thousands)   },
    subThousand: { ...VALUE_GROUPS.subThousand, count: subThousand.length, suppliers: strip(subThousand) },
  };
}

/**
 * Assemble the final tier output object.
 * Each tier contains:
 *   - summary counts & total value
 *   - suppliers[]        (flat ranked list, for simple iteration)
 *   - valueGroups{}      (same suppliers split by value magnitude)
 */
function buildOutput(top80, mid80_95, tail95, noValue, withValue) {
  const totalValue = withValue.reduce((s, x) => s + (x.value || 0), 0);

  // Add rank, return flat list with internal fields stripped
  const addRank = (arr) => arr.map((s, i) => {
    const { _tier, _valueGroup, ...rest } = s;
    return { rank: i + 1, ...rest };
  });

  const top80Ranked    = addRank(top80);
  const mid80_95Ranked = addRank(mid80_95);
  const tail95Ranked   = addRank(tail95);
  const noValueClean   = noValue.map(({ _tier, _valueGroup, ...rest }) => rest);

  return {
    summary: {
      totalSuppliers: withValue.length + noValue.length,
      withValue:      withValue.length,
      noValue:        noValue.length,
      totalValue:     parseFloat(totalValue.toFixed(2)),
      top80Count:     top80.length,
      mid80_95Count:  mid80_95.length,
      tail95Count:    tail95.length,
    },
    tiers: {
      top80: {
        ...TIER_LABELS.top80,
        count:       top80.length,
        totalValue:  parseFloat(top80.reduce((s, x) => s + (x.value || 0), 0).toFixed(2)),
        suppliers:   top80Ranked,
        valueGroups: groupByValueRange(top80Ranked.map((s, i) => ({ ...s, _tier: 'top80', _valueGroup: classifyValueGroup(s.value) }))),
      },
      mid80_95: {
        ...TIER_LABELS.mid80_95,
        count:       mid80_95.length,
        totalValue:  parseFloat(mid80_95.reduce((s, x) => s + (x.value || 0), 0).toFixed(2)),
        suppliers:   mid80_95Ranked,
        valueGroups: groupByValueRange(mid80_95Ranked.map(s => ({ ...s, _tier: 'mid80_95', _valueGroup: classifyValueGroup(s.value) }))),
      },
      tail95: {
        ...TIER_LABELS.tail95,
        count:       tail95.length,
        totalValue:  parseFloat(tail95.reduce((s, x) => s + (x.value || 0), 0).toFixed(2)),
        suppliers:   tail95Ranked,
        valueGroups: groupByValueRange(tail95Ranked.map(s => ({ ...s, _tier: 'tail95', _valueGroup: classifyValueGroup(s.value) }))),
      },
      noValue: {
        ...TIER_LABELS.noValue,
        count:       noValue.length,
        totalValue:  0,
        suppliers:   noValueClean,
        // No value grouping for noValue tier
      },
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Per-year builder
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Build yearly tier results from a raw sheet row array.
 *
 * For 2025: uses pre-computed cumulative % column from Excel (col[4]).
 * For 2024: recomputes cumulative % from scratch by sorting 2024 values.
 *
 * @param {Array[]} rows       - sheet_to_json result (header:1)
 * @param {string}  metricLabel
 * @param {object|null} scoresMap - Results sheet scores lookup
 * @returns {{ metric, '2025': {...}, '2024': {...} }}
 */
function buildYearlyTiers(rows, metricLabel, scoresMap) {
  const all = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r[0] == null) continue;

    const vendorNo = String(r[0]).trim();
    if (!vendorNo)  continue;

    // Vendor name may include "vendorNo_VendorName" prefix — strip it
    let vendorName = r[1] != null ? String(r[1]).trim() : vendorNo;
    const uIdx = vendorName.indexOf('_');
    if (uIdx !== -1) {
      const prefix = vendorName.slice(0, uIdx);
      if (/^\d+$/.test(prefix)) vendorName = vendorName.slice(uIdx + 1);
    }

    const value2024         = safeNum(r[2]);
    const value2025         = safeNum(r[3]);
    // ⚠️ FULL precision — do NOT round before classifyTier()
    const cumulativePct2025 = rawPct(r[4]);

    all.push({ vendorNo, vendorName, value2024, value2025, cumulativePct2025 });
  }

  return {
    metric: metricLabel,
    '2025': buildTiersFromCumulative(all, 'value2025', 'cumulativePct2025', scoresMap),
    '2024': buildTiersComputed(all, 'value2024', scoresMap),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// PUBLIC API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Parse CEI Buying and CEI Profit sheets from a workbook and return
 * percentile tier data split by year (2024 / 2025).
 *
 * @param {object}      wb        - XLSX workbook object (already loaded in RAM)
 * @param {object|null} scoresMap - Optional: Results sheet scores map from parseWorkbook()
 *                                  Enables KPI pillar scores per supplier.
 * @returns {{ ceiBuying, ceiProfit }}
 */
function buildPercentileTiers(wb, scoresMap = null) {
  const ceiBuyingResult = (() => {
    const ws = wb.Sheets['CEI Buying'];
    if (!ws) return null;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    return buildYearlyTiers(rows, 'CEI Buying (EUR)', scoresMap);
  })();

  const ceiProfitResult = (() => {
    const ws = wb.Sheets['CEI Profit'];
    if (!ws) return null;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    return buildYearlyTiers(rows, 'CEI Profit', scoresMap);
  })();

  return {
    ceiBuying: ceiBuyingResult,
    ceiProfit: ceiProfitResult,
  };
}

/**
 * Parse a workbook file/buffer and return percentile tier data.
 * Lightweight standalone version (no KPI enrichment without scoresMap).
 *
 * @param {string|Buffer} fileOrBuffer
 * @param {object|null}   scoresMap - optional, enables KPI fields
 * @returns {{ ceiBuying, ceiProfit }}
 */
function parsePercentiles(fileOrBuffer, scoresMap = null) {
  let wb;
  if (Buffer.isBuffer(fileOrBuffer)) {
    wb = XLSX.read(fileOrBuffer, { type: 'buffer', cellFormula: false, cellHTML: false });
  } else {
    wb = XLSX.readFile(fileOrBuffer, { cellFormula: false, cellHTML: false });
  }
  return buildPercentileTiers(wb, scoresMap);
}

module.exports = { parsePercentiles, buildPercentileTiers };
