/**
 * parser.js — Updated for NEW Supplier Performance Measurement file
 *
 * NEW Results sheet column layout (Row 1 = group headers, Row 2 = KPI headers, Row 3+ = data):
 *   [0]  Vendor No.
 *   [1]  Vendor Name
 *   [2]  Team
 *   ── Turnover and Margin ──
 *   [3]  2025 CEI Buying          score (0-3 numeric)
 *   [4]  2025 CEE Retail          score (0-3 numeric)
 *   [5]  2025 CEI Margin/Profit   score (0-3 numeric)
 *   [6]  2025 CEE CM1             score (0-3 numeric)
 *   ── Assortment & Innovation ──
 *   [7]  2025 New item%           score (0-3 numeric)
 *   [8]  2025 Pipeline development score (0-3 numeric)
 *   ── Quality Assurance ──
 *   [9]  Inspection Pass Rate     score (0-3 numeric)
 *   [10] Inspection Defect Rate   score (0-3 numeric)
 *   [11] Number of Re-inspection  score (0-3 numeric)
 *   [12] Return Rate              score (0-3 numeric)
 *   [13] Complain                 score (0-3 numeric)
 *   ── Fulfillment Operations ──
 *   [14] On-time rate             score (0-3 numeric)
 *   [15] Vessel booking           score (0-3 numeric)
 *   [16] Inspection booking       score (0-3 numeric)
 *   [17] Order confirmation       score (0-3 numeric)
 *   ── Terms & Conditions ──
 *   [18] Payment Terms            score (0-3 numeric)
 *   [19] Service remission %      score (0-3 numeric)
 *   [20] Agreed bonus %           score (0-3 numeric)
 *   [21] MOV required?            score (0-3 numeric)
 *   ── Pre-computed Pillar Scores ──  (single source of truth from Excel)
 *   [22] Turnover and Margin score   (weight 30%)
 *   [23] Assortment & Innovation score (weight 30%)
 *   [24] Quality Assurance score     (weight 25%)
 *   [25] Fulfillment Operations score (weight 15%)
 *   [26] Terms & Conditions score
 *   [27] Total Performance Evaluation
 *   [28] Business Class  (A/B/C/D)
 *   [29] Performance     (1.top / 2.prefered / 3.under / 4.critical)
 *   [30] Overall Class
 *   [31] CEI Buying 2025 (actual EUR buy value)
 *
 * Actual EUR values also read from individual KPI sheets:
 *   CEI Buying sheet  : col[0]=VendorNo, col[2]=2024 EUR, col[3]=2025 EUR
 *   CEI Profit sheet  : col[0]=VendorNo, col[2]=2024 HKD, col[3]=2025 HKD
 *   CEE Retail sheet  : col[13]=VendorNo, col[14]=SUM GMV FP ac (EUR)
 *   CEE CM1 sheet     : col[11]=VendorNo, col[12]=SUM CM1 Goods ac (EUR)
 */

const XLSX = require('xlsx');

/**
 * Map numeric KPI score (0–3) to a letter grade for backward-compatible display.
 *   3 → A,  2 → B,  1 → C,  0 → D (no data / lowest)
 */
function numToGrade(val) {
  if (val == null || isNaN(Number(val))) return null;
  const n = Number(val);
  if (n >= 3) return 'A';
  if (n >= 2) return 'B';
  if (n >= 1) return 'C';
  return 'D';
}

/** Safely parse a numeric value from a cell, returning 0 if invalid */
function safeNum(v, dp = 2) {
  if (v == null || isNaN(Number(v))) return 0;
  return parseFloat(Number(v).toFixed(dp));
}

// ─────────────────────────────────────────────────────────────────────────────
// Build lookup maps from the individual KPI sheets
// These provide the actual EUR/HKD values for each turnover KPI
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Build CEI Buying map from "CEI Buying" sheet.
 * Row layout: col[0]=VendorNo, col[2]=2024 Buy (EUR), col[3]=2025 Buy (EUR)
 */
function buildCeiBuyingMap(wb) {
  const map = {};
  const ws = wb.Sheets['CEI Buying'];
  if (!ws) return map;
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r[0] == null) continue;
    const vno = String(r[0]).trim();
    if (!vno) continue;
    map[vno] = {
      ceiBuying2024: safeNum(r[2], 2),
      ceiBuying2025: safeNum(r[3], 2),
    };
  }
  return map;
}

/**
 * Build CEI Profit map from "CEI Profit" sheet.
 * Row layout: col[0]=VendorNo, col[2]=2024 Profit (HKD), col[3]=2025 Profit (HKD)
 */
function buildCeiProfitMap(wb) {
  const map = {};
  const ws = wb.Sheets['CEI Profit'];
  if (!ws) return map;
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r[0] == null) continue;
    const vno = String(r[0]).trim();
    if (!vno) continue;
    map[vno] = {
      ceiProfit2024: safeNum(r[2], 2),
      ceiProfit2025: safeNum(r[3], 2),
    };
  }
  return map;
}

/**
 * Build CEE Retail map from "CEE Retail" sheet.
 * The sheet has both product-level rows AND aggregated supplier pivot.
 * Aggregated columns: col[13]=VendorNo, col[14]=SUM of GMV FP ac (EUR)
 */
function buildCeeRetailMap(wb) {
  const map = {};
  const ws = wb.Sheets['CEE Retail'];
  if (!ws) return map;
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    // col[13] = Supplier (aggregated), col[14] = SUM of GMV FP ac
    const vno = r[13] != null ? String(r[13]).trim() : null;
    if (!vno || isNaN(Number(r[14]))) continue;
    // Only overwrite if not already set (first occurrence = highest rank)
    if (!map[vno]) {
      map[vno] = { ceeRetail2025: safeNum(r[14], 2) };
    }
  }
  return map;
}

/**
 * Build CEE CM1 map from "CEE CM1" sheet.
 * Aggregated columns: col[11]=VendorNo, col[12]=SUM of CM1 Goods ac (EUR)
 */
function buildCeeCm1Map(wb) {
  const map = {};
  const ws = wb.Sheets['CEE CM1'];
  if (!ws) return map;
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    // col[11] = Supplier (aggregated), col[12] = SUM of CM1 Goods ac
    const vno = r[11] != null ? String(r[11]).trim() : null;
    if (!vno || isNaN(Number(r[12]))) continue;
    if (!map[vno]) {
      map[vno] = { ceeCm12025: safeNum(r[12], 2) };
    }
  }
  return map;
}

// ─────────────────────────────────────────────────────────────────────────────

/**
 * Parse the workbook and build lookup maps keyed by vendor number.
 *
 * @param {string|Buffer} fileOrBuffer
 * @returns {object} { gradesMap, scoresMap, results, turnoverMap }
 */
function parseWorkbook(fileOrBuffer) {
  let wb;

  if (Buffer.isBuffer(fileOrBuffer)) {
    wb = XLSX.read(fileOrBuffer, { type: 'buffer', cellFormula: false, cellHTML: false });
  } else {
    wb = XLSX.readFile(fileOrBuffer, { cellFormula: false, cellHTML: false });
  }

  return buildFromResults(wb);
}

function buildFromResults(wb) {
  const ws = wb.Sheets['Results'];
  if (!ws) return { gradesMap: {}, scoresMap: {}, results: [], turnoverMap: {} };

  // ── Build all KPI sheet lookup maps ─────────────────────────────────────────
  const ceiBuyingMap = buildCeiBuyingMap(wb);
  const ceiProfitMap = buildCeiProfitMap(wb);
  const ceeRetailMap = buildCeeRetailMap(wb);
  const ceeCm1Map = buildCeeCm1Map(wb);

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  const gradesMap = {};
  const scoresMap = {};
  const turnoverMap = {};
  const results = [];

  // Helper: safely extract a numeric score cell from Results sheet
  const n = (row, col, dp = 2) => {
    const v = row[col];
    if (v == null || String(v).trim() === '' || isNaN(Number(v))) return 'NA';
    return parseFloat(Number(v).toFixed(dp));
  };

  // Helper: extract raw numeric KPI (0–3) and convert to letter grade
  const g = (row, col) => {
    const v = row[col];
    if (v == null) return null;
    const str = String(v).trim().toUpperCase();
    if (['A', 'B', 'C', 'D'].includes(str)) return str;
    return numToGrade(v);
  };

  // Helper: extract raw numeric KPI value (0–3)
  const kpiNum = (row, col) => {
    const v = row[col];
    if (v == null || isNaN(Number(v))) return null;
    return Number(v);
  };

  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row[0] == null) break;

    const vendorNo = String(row[0]).trim();
    if (!vendorNo) continue;

    // ── Grade columns (3–21) — numeric KPI scores converted to A/B/C/D ────────
    gradesMap[vendorNo] = {
      // Turnover and Margin
      ceiBuying: g(row, 3),
      ceeRetail: g(row, 4),
      ceiProfit: g(row, 5),
      ceeCm1: g(row, 6),
      // Assortment & Innovation
      newItem: g(row, 7),
      pipeline: g(row, 8),
      // Quality Assurance
      passRate: g(row, 9),
      defectRate: g(row, 10),
      reInspect: g(row, 11),
      returnRate: g(row, 12),
      complain: g(row, 13),
      // Fulfillment Operations
      onTime: g(row, 14),
      vessel: g(row, 15),
      inspBook: g(row, 16),
      orderConf: g(row, 17),
      // Terms & Conditions
      payment: g(row, 18),
      remission: g(row, 19),
      bonus: g(row, 20),
      mov: g(row, 21),
    };

    // ── Pre-computed pillar scores from Excel (cols 22–27) ───────────────────
    const turnoverMarginScore = n(row, 22, 2);
    const assortmentInnovationScore = n(row, 23, 2);
    const qualityScore = n(row, 24, 2);
    const fulfillmentScore = n(row, 25, 2);
    const termsScore = n(row, 26, 2);
    const total = n(row, 27, 2);

    // Extra classification fields
    const businessClass = row[28] != null ? String(row[28]).trim() : null;
    const performance = row[29] != null ? String(row[29]).trim() : null;
    const overallClass = row[30] != null ? String(row[30]).trim() : null;

    // ── Actual financial values from individual KPI sheets ───────────────────
    const ceiBuy = ceiBuyingMap[vendorNo] || {};
    const ceiProf = ceiProfitMap[vendorNo] || {};
    const ceeRet = ceeRetailMap[vendorNo] || {};
    const ceeCm1 = ceeCm1Map[vendorNo] || {};

    scoresMap[vendorNo] = {
      turnoverMargin: turnoverMarginScore,
      assortmentInnovation: assortmentInnovationScore,
      quality: qualityScore,
      fulfillment: fulfillmentScore,
      terms: termsScore,
      total,
      businessClass,
      performance,
      overallClass,
      // Actual EUR values
      ceiBuying2024: ceiBuy.ceiBuying2024 || 0,
      ceiBuying2025: ceiBuy.ceiBuying2025 || safeNum(row[31], 2),  // fallback to Results col 31
      ceeRetail2025: ceeRet.ceeRetail2025 || 0,
      ceiProfit2024: ceiProf.ceiProfit2024 || 0,
      ceiProfit2025: ceiProf.ceiProfit2025 || 0,
      ceeCm12025: ceeCm1.ceeCm12025 || 0,
    };

    // ── Turnover map — full data for chart endpoint ──────────────────────────
    turnoverMap[vendorNo] = {
      vendorNo,
      vendorName: row[1] ? String(row[1]).trim() : '',
      team: row[2] ? String(row[2]).trim() : '',
      // Pillar score out of 30
      turnoverScore: turnoverMarginScore,
      // Business classification
      businessClass,
      performance,
      // ── Actual EUR/HKD financial values (all 4 turnover KPIs) ──────────────
      actuals: {
        ceiBuying2024: ceiBuy.ceiBuying2024 || 0,   // EUR — previous year buy value
        ceiBuying2025: ceiBuy.ceiBuying2025 || safeNum(row[31], 2), // EUR — current year buy value
        ceeRetail2025: ceeRet.ceeRetail2025 || 0,   // EUR — retail GMV current year
        ceiProfit2024: ceiProf.ceiProfit2024 || 0,   // HKD — profit previous year
        ceiProfit2025: ceiProf.ceiProfit2025 || 0,   // HKD — profit current year
        ceeCm12025: ceeCm1.ceeCm12025 || 0,   // EUR — CM1 Goods current year
      },
      // Individual KPI numeric scores (0–3) for radar/breakdown chart
      kpis: {
        ceiBuying: kpiNum(row, 3),
        ceeRetail: kpiNum(row, 4),
        ceiProfit: kpiNum(row, 5),
        ceeCm1: kpiNum(row, 6),
      },
      // Letter grades for badge display
      grades: {
        ceiBuying: g(row, 3),
        ceeRetail: g(row, 4),
        ceiProfit: g(row, 5),
        ceeCm1: g(row, 6),
      },
    };

    results.push({
      vendorNo,
      vendorName: row[1] ? String(row[1]).trim() : '',
      team: row[2] ? String(row[2]).trim() : '',
    });
  }

  return { gradesMap, scoresMap, results, turnoverMap };
}

module.exports = { parseWorkbook };
