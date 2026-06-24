/**
 * grader.js
 * Assigns A/B/C/D grades to each KPI for a given supplier.
 * Reads pre-calculated grades directly from the Results sheet via parser.js.
 *
 * NEW FILE: 4 pillars
 *   - Turnover and Margin     (4 KPIs: ceiBuying, ceeRetail, ceiProfit, ceeCm1)
 *   - Assortment & Innovation (2 KPIs: newItem, pipeline)
 *   - Quality Assurance       (5 KPIs: passRate, defectRate, reInspect, returnRate, complain)
 *   - Fulfillment Operations  (4 KPIs: onTime, vessel, inspBook, orderConf)
 *   - Terms & Conditions      (4 KPIs: payment, remission, bonus, mov)
 */

const VALID_GRADES = new Set(['A', 'B', 'C', 'D']);

/**
 * Grade all KPIs for a single supplier.
 *
 * @param {string} vendorNo
 * @param {object} lookups - from parser.parseWorkbook() — must contain { gradesMap }
 * @returns {{ grades: object, dataWarnings: string[] }}
 */
function gradeSupplier(vendorNo, lookups) {
  const warnings  = [];
  const rawGrades = lookups.gradesMap?.[vendorNo] || {};

  function grade(rawVal, kpiLabel) {
    const v = rawVal != null ? String(rawVal).toUpperCase().trim() : null;
    if (v && VALID_GRADES.has(v)) return v;
    warnings.push(kpiLabel);
    return null;
  }

  // ── Turnover and Margin (4 KPIs) ────────────────────────────────────────────
  const ceiBuying  = grade(rawGrades.ceiBuying,  'CEI Buying');
  const ceeRetail  = grade(rawGrades.ceeRetail,  'CEE Retail');
  const ceiProfit  = grade(rawGrades.ceiProfit,  'CEI Profit');
  const ceeCm1     = grade(rawGrades.ceeCm1,     'CEE CM1');

  // ── Assortment & Innovation (2 KPIs) ────────────────────────────────────────
  const newItem    = grade(rawGrades.newItem,    'New Item%');
  const pipeline   = grade(rawGrades.pipeline,   'Pipeline Development');

  // ── Quality Assurance (5 KPIs) ──────────────────────────────────────────────
  const passRate   = grade(rawGrades.passRate,   'Pass Rate');
  const defectRate = grade(rawGrades.defectRate, 'Defect Rate');
  const reInspect  = grade(rawGrades.reInspect,  'Re-inspection');
  const returnRate = grade(rawGrades.returnRate, 'Return Rate');
  const complain   = grade(rawGrades.complain,   'Complain');

  // ── Fulfillment Operations (4 KPIs) ─────────────────────────────────────────
  const onTime     = grade(rawGrades.onTime,     'On-time Rate');
  const vessel     = grade(rawGrades.vessel,     'Vessel Booking');
  const inspBook   = grade(rawGrades.inspBook,   'Inspection Booking');
  const orderConf  = grade(rawGrades.orderConf,  'Order Confirmation');

  // ── Terms & Conditions (4 KPIs) ─────────────────────────────────────────────
  const payment    = grade(rawGrades.payment,    'Payment Terms');
  const remission  = grade(rawGrades.remission,  'Remission %');
  const bonus      = grade(rawGrades.bonus,      'Agreed Bonus');
  const mov        = grade(rawGrades.mov,        'MOV Required');

  return {
    grades: {
      // Turnover and Margin
      ceiBuying, ceeRetail, ceiProfit, ceeCm1,
      // Assortment & Innovation
      newItem, pipeline,
      // Quality
      passRate, defectRate, reInspect, returnRate, complain,
      // Fulfillment
      onTime, vessel, inspBook, orderConf,
      // Terms
      payment, remission, bonus, mov,
    },
    dataWarnings: warnings,
  };
}

module.exports = { gradeSupplier };
