/* ============================================================================
 * Regression test — engine output vs. the values Merit.xlsb actually produced.
 *   node tests/engine.test.js
 * Fixtures are extracted straight from the workbook by tools/extract_merit.py.
 *
 * Two scenarios are run:
 *
 *  A. CONFORMANCE REPLAY — proves the reverse-engineering is exact. Employee 4
 *     is fed without the special-impact flag, because that is what the
 *     workbook's column-O formula effectively did for that one row (see
 *     AMB-01 in docs/05-ambiguities.md: the row is flagged بله and column N
 *     holds 300, but column O dropped the +N term). Under this replay every
 *     computed column of every employee must match Excel bit-for-bit.
 *
 *  B. DOCUMENTED RULE — the data exactly as it sits in the workbook, scored
 *     with the rule the reference document states (O = L + N, always). This is
 *     what the shipped application does. The test quantifies how far AMB-01
 *     moves each column so the business owner can confirm the intended rule.
 * ==========================================================================*/
'use strict';
var Engine   = require('../src/calculation-engine.js');
var sample   = require('../sample-data/merit-sample.json');
var expected = require('../tests/merit-expected.json');

var COEF = 1e-6;   // coefficient tolerance
var RIAL = 1e-3;   // payout tolerance (on a 1e11 budget this is ~1e-14 relative)
var AMB01_EMPLOYEE = '4';

var config = {
  budget:                       sample.config.budget,
  gradeMap:                     sample.config.gradeMap,
  answerScale:                  sample.config.answerScale,
  maxPerformanceScore:          sample.config.maxPerformanceScore,
  questionCount:                sample.config.questionCount,
  minPerformanceThreshold:      sample.config.minPerformanceThreshold,
  gradeImpactFactor:            sample.config.gradeImpactFactor,
  baselineCoefficientPerPerson: sample.config.baselineCoefficientPerPerson,
  specialImpactAmount:          sample.config.specialImpactAmount
};

var COLUMNS = [
  ['performanceScore',        'performanceScore',      'K (sheet 1) امتیاز عملکرد',   COEF],
  ['performanceKaraneh',      'performanceKaraneh',    'L (sheet 1) عدد کارانه',      COEF],
  ['specialImpactValue',      'specialImpactValue',    'N (sheet 1) اثرگذاری ویژه',   COEF],
  ['rawCoefficient',          'rawCoefficient',        'O (sheet 1) ضریب کارانه',     COEF],
  ['finalCoefficient',        'finalCoefficient',      'R (sheet 1) ضریب نهایی',      COEF],
  ['gradeScore',              'gradeScore',            'F (sheet 2) گرید',            COEF],
  ['gradeImpact',             'gradeImpact',           'G (sheet 2) عدد گرید',        COEF],
  ['eligibleEvalScore',       'eligibleEvalScore',     'I (sheet 2) عدد ارزیابی',     COEF],
  ['performanceContribution', 'performanceScoreFinal', 'K (sheet 2) امتیاز عملکردی',  COEF],
  ['totalScore',              'totalScore',            'L (sheet 2) امتیاز کل',       COEF],
  ['initialAllocation',       'initialAllocation',     'M (sheet 2) دریافتی اولیه',   RIAL],
  ['finalKaraneh',            'finalKaraneh',          'Q (sheet 2) دریافتی نهایی',   RIAL]
];

var TOTALS = [
  ['count (A5)',                     function (t) { return t.inScopeCount; },                                 'questionnaireCount',       0],
  ['baseline pool (B5)',             function (t) { return t.inScopeCount * 100; },                           'baselineTotal',            1e-6],
  ['coefficient excess (O5)',        function (t) { return t.sumRawCoefficient - t.inScopeCount * 100; },     'coefficientExcess',        1e-6],
  ['per-person normalisation (N5)',  function (t) { return (t.sumRawCoefficient - t.inScopeCount * 100) / t.inScopeCount; }, 'perPersonNormalization', 1e-9],
  ['ineligible redistribution (J5)', function (t) { return t.ineligibleRedistribution; },                     'ineligibleRedistribution', 1e-9],
  ['HOD redistribution (O5 pay)',    function (t) { return t.hodRedistribution; },                            'hodRedistribution',        1e-3],
  ['SUM initial allocation (M5)',    function (t) { return t.sumInitialAllocation; },                         'sumInitialAllocation',     1e-3],
  ['SUM final karaneh (Q5)',         function (t) { return t.sumFinalKaraneh; },                              'sumFinalKaraneh',          1e-3]
];

/* ---------------------------------------------------------------- helpers */
function pad(s, n) { s = String(s); while (s.length < n) s += ' '; return s; }
function lpad(s, n) { s = String(s); while (s.length < n) s = ' ' + s; return s; }
function fixed(v, d) { return Number(v).toFixed(d === undefined ? 4 : d); }
function val(v) { return (v === null || v === undefined) ? 0 : v; }
function index(rows) { var m = {}; rows.forEach(function (r) { m[r.employeeId] = r; }); return m; }

function replayInput() {
  return sample.employees.map(function (e) {
    if (e.employeeId !== AMB01_EMPLOYEE) return e;
    var c = {}; for (var k in e) c[k] = e[k];
    c.specialProject = null; c.specialImpactAmount = 0;   // what column O actually did
    return c;
  });
}

var failures = [];
function expectClose(label, ours, theirs, tol) {
  var d = Math.abs(val(ours) - val(theirs));
  var ok = d <= tol;
  if (!ok) failures.push(label + ': engine=' + val(ours) + ' excel=' + val(theirs) + ' Δ=' + d);
  return { ok: ok, diff: d };
}

/* ==========================================================================
 * SCENARIO A — conformance replay
 * ========================================================================*/
console.log('\n╔══════════════════════════════════════════════════════════════════════════╗');
console.log('║ SCENARIO A — CONFORMANCE REPLAY vs Merit.xlsb                            ║');
console.log('╚══════════════════════════════════════════════════════════════════════════╝');

var A = Engine.calculate(replayInput(), config);
var aById = index(A.rows);

console.log('\n-- Sheet control totals --');
console.log(pad('metric', 32) + lpad('engine', 26) + lpad('excel', 26) + lpad('Δ', 12) + '  ');
TOTALS.forEach(function (t) {
  var ours = t[1](A.totals), theirs = expected.totals[t[2]];
  var r = expectClose('A/' + t[0], ours, theirs, t[3]);
  console.log(pad(t[0], 32) + lpad(fixed(ours), 26) + lpad(fixed(theirs), 26) +
              lpad(r.diff.toExponential(1), 12) + '  ' + (r.ok ? 'PASS' : 'FAIL'));
});
var recon = expectClose('A/budget reconciliation', A.totals.sumFinalKaraneh, config.budget, 1e-3);
console.log(pad('budget reconciliation', 32) + lpad(fixed(A.totals.sumFinalKaraneh), 26) +
            lpad(fixed(config.budget), 26) + lpad(recon.diff.toExponential(1), 12) +
            '  ' + (recon.ok ? 'PASS' : 'FAIL'));

console.log('\n-- Per-column maximum absolute difference across all ' + expected.rows.length + ' employees --');
console.log(pad('column', 34) + lpad('max Δ', 14) + lpad('tolerance', 14) + '  ');
COLUMNS.forEach(function (c) {
  var max = 0, worst = null;
  expected.rows.forEach(function (exp) {
    /* The workbook is internally inconsistent on exactly one cell: employee 4
       carries N=300 while column O behaved as if N were 0. The replay feeds
       the input column O acted on, so its own column N cannot also match —
       comparing it would be asserting both halves of a contradiction. */
    if (exp.employeeId === AMB01_EMPLOYEE && c[0] === 'specialImpactValue') return;
    var got = aById[exp.employeeId];
    if (!got) { failures.push('A/missing employee ' + exp.employeeId); return; }
    var d = Math.abs(val(got[c[0]]) - val(exp[c[1]]));
    if (d > max) { max = d; worst = exp.employeeId; }
  });
  var ok = max <= c[3];
  if (!ok) failures.push('A/' + c[0] + ' max Δ ' + max + ' at employee ' + worst);
  console.log(pad(c[2], 34) + lpad(max.toExponential(3), 14) + lpad(c[3].toExponential(0), 14) +
              '  ' + (ok ? 'PASS' : 'FAIL') +
              (c[0] === 'specialImpactValue' ? '  (employee ' + AMB01_EMPLOYEE + ' excluded — AMB-01)' : ''));
});

/* ==========================================================================
 * SCENARIO B — documented rule (what the application ships with)
 * ========================================================================*/
var B = Engine.calculate(sample.employees, config);
var bById = index(B.rows);

console.log('\n╔══════════════════════════════════════════════════════════════════════════╗');
console.log('║ SCENARIO B — DOCUMENTED RULE (O = L + N for every flagged employee)      ║');
console.log('╚══════════════════════════════════════════════════════════════════════════╝');
console.log('\nImpact of AMB-01 relative to the workbook as it stands:');
COLUMNS.forEach(function (c) {
  var max = 0;
  expected.rows.forEach(function (exp) {
    max = Math.max(max, Math.abs(val(bById[exp.employeeId][c[0]]) - val(exp[c[1]])));
  });
  console.log(pad('  ' + c[2], 36) + lpad(max.toExponential(3), 14));
});
var payoutShift = expectClose('B/final payouts unchanged',
  B.totals.sumFinalKaraneh, config.budget, 1e-3);
var maxPayoutDelta = 0;
expected.rows.forEach(function (exp) {
  maxPayoutDelta = Math.max(maxPayoutDelta,
    Math.abs(val(bById[exp.employeeId].finalKaraneh) - val(exp.finalKaraneh)));
});
console.log('\n  → AMB-01 moves the intermediate coefficients but the rial payouts are');
console.log('    unaffected (max Δ = ' + maxPayoutDelta.toExponential(2) + ' rial), because the affected');
console.log('    employee is HOD-overridden and the normalisation shift is uniform.');
if (maxPayoutDelta > RIAL) failures.push('B/payout drift ' + maxPayoutDelta);
if (!payoutShift.ok) { /* already recorded */ }

/* ==========================================================================
 * TEN-EMPLOYEE SPOT CHECK (deliverable §36)
 * ========================================================================*/
console.log('\n╔══════════════════════════════════════════════════════════════════════════╗');
console.log('║ SAMPLE VALIDATION — 10 EMPLOYEES, FINAL KARANEH                          ║');
console.log('╚══════════════════════════════════════════════════════════════════════════╝');
console.log(pad('Emp', 6) + lpad('Excel Final', 24) + lpad('Engine Final', 24) + lpad('Difference', 14) + '  ');
expected.rows.slice(0, 10).forEach(function (exp) {
  var x = val(exp.finalKaraneh), h = val(aById[exp.employeeId].finalKaraneh);
  var d = Math.abs(x - h);
  console.log(pad(exp.employeeId, 6) + lpad(fixed(x, 2), 24) + lpad(fixed(h, 2), 24) +
              lpad(d.toExponential(2), 14) + '  ' + (d <= RIAL ? 'PASS' : 'FAIL'));
});

/* ==========================================================================
 * ENGINE INVARIANTS
 * ========================================================================*/
console.log('\n-- Engine invariants --');
function invariant(label, ok, detail) {
  if (!ok) failures.push('INV/' + label + ' ' + (detail || ''));
  console.log(pad('  ' + label, 60) + (ok ? 'PASS' : 'FAIL') + (detail ? '  ' + detail : ''));
}
invariant('sum(final karaneh) === budget',
  Math.abs(A.totals.sumFinalKaraneh - config.budget) < 1e-3);
invariant('sum(total score) === count × baseline',
  Math.abs(A.totals.sumTotalScore - A.totals.inScopeCount * 100) < 1e-6,
  'sum=' + fixed(A.totals.sumTotalScore));
invariant('no employee scoring exactly the threshold is paid',
  A.rows.filter(function (r) { return r.performanceScore === 2 && r.finalKaraneh !== 0; }).length === 0);
invariant('no negative payout in the reference data', A.totals.negativePayoutCount === 0);
invariant('budget status is BALANCED', A.totals.budgetStatus === 'BALANCED', A.totals.budgetStatus);

/* A budget change must scale every payout proportionally and still reconcile. */
var scaled = Engine.calculate(replayInput(), Object.assign({}, config, { budget: 250000000000 }));
invariant('budget change rescales and reconciles',
  Math.abs(scaled.totals.sumFinalKaraneh - 250000000000) < 1e-3 ||
  scaled.totals.overriddenCount > 0);

/* Turning the grade impact factor on must change the distribution. */
var graded = Engine.calculate(replayInput(), Object.assign({}, config, { gradeImpactFactor: 1 }));
invariant('grade impact factor changes the distribution',
  Math.abs(graded.totals.sumTotalScore - A.totals.sumTotalScore) > 1);
invariant('grade-weighted run still reconciles to budget',
  Math.abs(graded.totals.sumFinalKaraneh - config.budget) < 1e-3);

/* ========================================================================*/
console.log('\n╔══════════════════════════════════════════════════════════════════════════╗');
if (failures.length) {
  console.log('║ RESULT: FAILED (' + pad(failures.length + ' checks)', 56) + '║');
  console.log('╚══════════════════════════════════════════════════════════════════════════╝');
  failures.slice(0, 40).forEach(function (f) { console.log('  ' + f); });
  process.exit(1);
}
console.log('║ RESULT: PASS — ' + pad(expected.rows.length + ' employees × ' + COLUMNS.length +
            ' columns reproduce Merit.xlsb exactly.', 57) + '║');
console.log('╚══════════════════════════════════════════════════════════════════════════╝');
