/**
 * scorer.js
 * Returns pillar scores and tier from pre-computed Excel values.
 *
 * NEW FILE — 4 pillars:
 *   Turnover and Margin     30%  (max 30)
 *   Assortment & Innovation 30%  (max 30)
 *   Quality Assurance       25%  (max 25)
 *   Fulfillment Operations  15%  (max 15)
 *   Terms & Conditions      (included in total — no fixed weight shown in header)
 *
 * Scores are read directly from the Results sheet (cols 22–27) to ensure
 * 100% parity with the Excel formula — no re-computation needed.
 */

/**
 * Build the pillars object and total score from pre-computed Excel scores.
 *
 * @param {object} excelScores - from parser.scoresMap[vendorNo]
 * @returns {object} { pillars, totalScore, tier, businessClass, performance, ceiBuying2025 }
 */
function calcScores(excelScores) {
  const round2OrNA = (v) => v === 'NA' ? 'NA' : parseFloat(Number(v || 0).toFixed(2));

  const turnoverMargin = round2OrNA(excelScores.turnoverMargin);
  const assortmentInnovation = round2OrNA(excelScores.assortmentInnovation);
  const quality = round2OrNA(excelScores.quality);
  const fulfillment = round2OrNA(excelScores.fulfillment);
  const terms = round2OrNA(excelScores.terms);

  // Use Excel's own total
  const totalScore = round2OrNA(excelScores.total);

  return {
    pillars: {
      turnoverMargin: { score: turnoverMargin, weight: 'NA', maxScore: null, label: 'Turnover & Margin' },
      assortmentInnovation: { score: assortmentInnovation, weight: '30%', maxScore: 30, label: 'Assortment & Innovation' },
      quality: { score: quality, weight: '30%', maxScore: 30, label: 'Quality Assurance' },
      fulfillment: { score: fulfillment, weight: '25%', maxScore: 25, label: 'Fulfillment Operations' },
      terms: { score: terms, weight: '15%', maxScore: 15, label: 'Terms & Conditions' },
    },
    totalScore,
    tier: assignTier(totalScore),
    businessClass: excelScores.businessClass || null,
    performance: excelScores.performance || null,
    overallClass: excelScores.overallClass || null,
    // ── All 4 turnover KPI actual financial values ─────────────────────────
    actuals: {
      ceiBuying2024: excelScores.ceiBuying2024 || 0,  // EUR — previous year
      ceiBuying2025: excelScores.ceiBuying2025 || 0,  // EUR — current year
      ceeRetail2025: excelScores.ceeRetail2025 || 0,  // EUR — retail GMV
      ceiProfit2024: excelScores.ceiProfit2024 || 0,  // HKD — profit prev year
      ceiProfit2025: excelScores.ceiProfit2025 || 0,  // HKD — profit curr year
      ceeCm12025: excelScores.ceeCm12025 || 0,  // EUR — CM1 Goods
    },
  };
}

function assignTier(score) {
  if (score >= 85) return 'Strategic Partner';
  if (score >= 70) return 'Preferred Supplier';
  if (score >= 50) return 'Approved Supplier';
  return 'At Risk';
}

module.exports = { calcScores, assignTier };
