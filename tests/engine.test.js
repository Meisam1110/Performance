/* ============================================================================
 * Regression test — engine output vs. the values Merit.xlsb actually produced.
 *   node tests/engine.test.js
 * Fixtures are extracted straight from the workbook by tools/extract_merit.py.
 *
 * Every computed column of every employee must match, with no exceptions.
 * The one row that used to disagree (employee 4, flagged for special impact
 * yet scoring 90) is explained by the special-impact gate: the bonus applies
 * only above `specialImpactMinScore`. Column N of the workbook holds the
 * amount the manager's answer implies; column O applies the gate.
 * ==========================================================================*/
'use strict';
var Engine   = require('../src/calculation-engine.js');
var sample   = require('../sample-data/merit-sample.json');
var expected = require('../tests/merit-expected.json');

var COEF = 1e-6;   // coefficient tolerance
var RIAL = 1e-3;   // payout tolerance (on a 1e11 budget this is ~1e-14 relative)

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
  ['specialImpactEntered',    'specialImpactValue',    'N (sheet 1) اثرگذاری ویژه',   COEF],
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

var failures = [];
function record(label, ok, detail) {
  if (!ok) failures.push(label + (detail ? ' — ' + detail : ''));
  return ok;
}

var result = Engine.calculate(sample.employees, config);
var byId = {};
result.rows.forEach(function (r) { byId[r.employeeId] = r; });

/* ==========================================================================
 * Sheet control totals
 * ========================================================================*/
console.log('\n╔══════════════════════════════════════════════════════════════════════════╗');
console.log('║ CONFORMANCE vs Merit.xlsb — sheet control totals                         ║');
console.log('╚══════════════════════════════════════════════════════════════════════════╝');
console.log(pad('metric', 32) + lpad('engine', 26) + lpad('excel', 26) + lpad('Δ', 12));
TOTALS.forEach(function (t) {
  var ours = t[1](result.totals), theirs = expected.totals[t[2]];
  var d = Math.abs(val(ours) - val(theirs));
  var ok = record('total ' + t[0], d <= t[3], 'Δ=' + d);
  console.log(pad(t[0], 32) + lpad(fixed(ours), 26) + lpad(fixed(theirs), 26) +
              lpad(d.toExponential(1), 12) + '  ' + (ok ? 'PASS' : 'FAIL'));
});
var reconD = Math.abs(result.totals.sumFinalKaraneh - config.budget);
console.log(pad('budget reconciliation', 32) + lpad(fixed(result.totals.sumFinalKaraneh), 26) +
            lpad(fixed(config.budget), 26) + lpad(reconD.toExponential(1), 12) +
            '  ' + (record('budget reconciliation', reconD <= 1e-3) ? 'PASS' : 'FAIL'));

/* ==========================================================================
 * Every computed column, every employee
 * ========================================================================*/
console.log('\n-- Per-column maximum absolute difference across all ' +
            expected.rows.length + ' employees (no exclusions) --');
console.log(pad('column', 34) + lpad('max Δ', 14) + lpad('tolerance', 14));
COLUMNS.forEach(function (c) {
  var max = 0, worst = null;
  expected.rows.forEach(function (exp) {
    var got = byId[exp.employeeId];
    if (!got) { record('missing employee ' + exp.employeeId, false); return; }
    var d = Math.abs(val(got[c[0]]) - val(exp[c[1]]));
    if (d > max) { max = d; worst = exp.employeeId; }
  });
  var ok = record(c[0], max <= c[3], 'max Δ ' + max + ' at employee ' + worst);
  console.log(pad(c[2], 34) + lpad(max.toExponential(3), 14) +
              lpad(c[3].toExponential(0), 14) + '  ' + (ok ? 'PASS' : 'FAIL'));
});

/* ==========================================================================
 * The special-impact gate — the rule that resolves the last discrepancy
 * ========================================================================*/
console.log('\n-- Special-impact gate (bonus applies only at ' +
            Engine.DEFAULT_CONFIG.specialImpactMinScore + ' karaneh points or above) --');
console.log(pad('emp', 6) + lpad('score', 8) + lpad('karaneh', 10) + lpad('flagged', 10) +
            lpad('entered', 10) + lpad('applied', 10) + lpad('excel O', 12));
expected.rows.forEach(function (exp) {
  var r = byId[exp.employeeId];
  if (!r.specialProject) return;
  var ok = record('gate/' + exp.employeeId,
    Math.abs(r.rawCoefficient - exp.rawCoefficient) < COEF);
  console.log(pad(exp.employeeId, 6) + lpad(fixed(r.performanceScore, 2), 8) +
              lpad(fixed(r.performanceKaraneh, 2), 10) + lpad('بله', 10) +
              lpad(fixed(r.specialImpactEntered, 0), 10) +
              lpad(fixed(r.specialImpactValue, 0), 10) +
              lpad(fixed(exp.rawCoefficient, 0), 12) + '  ' + (ok ? 'PASS' : 'FAIL'));
});

/* ==========================================================================
 * Ten-employee spot check
 * ========================================================================*/
console.log('\n-- Sample validation: 10 employees, final karaneh --');
console.log(pad('Emp', 6) + lpad('Excel Final', 24) + lpad('Engine Final', 24) + lpad('Difference', 14));
expected.rows.slice(0, 10).forEach(function (exp) {
  var x = val(exp.finalKaraneh), h = val(byId[exp.employeeId].finalKaraneh);
  var d = Math.abs(x - h);
  record('spot/' + exp.employeeId, d <= RIAL);
  console.log(pad(exp.employeeId, 6) + lpad(fixed(x, 2), 24) + lpad(fixed(h, 2), 24) +
              lpad(d.toExponential(2), 14) + '  ' + (d <= RIAL ? 'PASS' : 'FAIL'));
});

/* ==========================================================================
 * Invariants and configurability
 * ========================================================================*/
console.log('\n-- Engine invariants --');
function invariant(label, ok, detail) {
  record('INV/' + label, ok, detail);
  console.log(pad('  ' + label, 62) + (ok ? 'PASS' : 'FAIL') + (detail ? '  ' + detail : ''));
}

invariant('sum(final karaneh) === budget',
  Math.abs(result.totals.sumFinalKaraneh - config.budget) < 1e-3);
invariant('sum(total score) === count × baseline',
  Math.abs(result.totals.sumTotalScore - result.totals.inScopeCount * 100) < 1e-6,
  'sum=' + fixed(result.totals.sumTotalScore));
invariant('no employee scoring exactly the threshold is paid',
  result.rows.filter(function (r) { return r.performanceScore === 2 && r.finalKaraneh !== 0; }).length === 0);
invariant('no negative payout in the reference data', result.totals.negativePayoutCount === 0);
invariant('budget status is BALANCED', result.totals.budgetStatus === 'BALANCED', result.totals.budgetStatus);

/* Gate is configurable: dropping it to zero must pay the blocked employee. */
var ungated = Engine.calculate(sample.employees,
  Object.assign({}, config, { specialImpactMinScore: 0 }));
var emp4Gated = byId['4'].specialImpactValue;
var emp4Ungated = ungated.rows.filter(function (r) { return r.employeeId === '4'; })[0].specialImpactValue;
invariant('gate is configurable (0 pays the blocked employee)',
  emp4Gated === 0 && emp4Ungated === 300, emp4Gated + ' → ' + emp4Ungated);
invariant('removing the gate still reconciles to budget',
  Math.abs(ungated.totals.sumFinalKaraneh - config.budget) < 1e-3);

/* Question weights: doubling one question must move the score, and an
   all-equal weighting must reproduce the plain average exactly. */
var weighted = Engine.calculate(sample.employees, Object.assign({}, config, {
  questions: Engine.DEFAULT_CONFIG.questions.map(function (q, i) {
    return { id: q.id, text: q.text, weight: i === 0 ? 2 : q.weight, scored: q.scored };
  })
}));
invariant('question weights change the score',
  Math.abs(weighted.totals.sumRawCoefficient - result.totals.sumRawCoefficient) > 1);
invariant('weighted run still reconciles to budget',
  Math.abs(weighted.totals.sumFinalKaraneh - config.budget) < 1e-3);

var explicitEqual = Engine.calculate(sample.employees, Object.assign({}, config, {
  questions: Engine.DEFAULT_CONFIG.questions.slice()
}));
invariant('equal weights reproduce the plain average',
  Math.abs(explicitEqual.totals.sumRawCoefficient - result.totals.sumRawCoefficient) < 1e-9);

/* Adding a fifth scored question must be possible without touching code. */
var fiveQ = Engine.calculate(sample.employees, Object.assign({}, config, {
  questions: Engine.DEFAULT_CONFIG.questions.map(function (q) {
    return { id: q.id, text: q.text, weight: 1, scored: true };
  }),
  questionCount: 5
}));
invariant('a fifth scored question is picked up from config',
  fiveQ.totals.inScopeCount === 100 &&
  Math.abs(fiveQ.totals.sumRawCoefficient - result.totals.sumRawCoefficient) > 1);

var scaled = Engine.calculate(sample.employees, Object.assign({}, config, { budget: 250000000000 }));
invariant('budget change rescales and reconciles',
  Math.abs(scaled.totals.sumFinalKaraneh - 250000000000) < 1e-3 || scaled.totals.overriddenCount > 0);

var graded = Engine.calculate(sample.employees, Object.assign({}, config, { gradeImpactFactor: 1 }));
invariant('grade impact factor changes the distribution',
  Math.abs(graded.totals.sumTotalScore - result.totals.sumTotalScore) > 1);
invariant('grade-weighted run still reconciles to budget',
  Math.abs(graded.totals.sumFinalKaraneh - config.budget) < 1e-3);

/* ========================================================================*/
console.log('\n╔══════════════════════════════════════════════════════════════════════════╗');
if (failures.length) {
  console.log('║ RESULT: FAILED (' + pad(failures.length + ' checks)', 56) + '║');
  console.log('╚══════════════════════════════════════════════════════════════════════════╝');
  failures.slice(0, 40).forEach(function (f) { console.log('  • ' + f); });
  process.exit(1);
}
console.log('║ RESULT: PASS — ' + pad(expected.rows.length + ' employees × ' + COLUMNS.length +
            ' columns reproduce Merit.xlsb exactly.', 57) + '║');
console.log('╚══════════════════════════════════════════════════════════════════════════╝');
