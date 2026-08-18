/* ============================================================================
 * End-to-end test of the built single-file application.
 *   node tools/build.js && node tests/ui.e2e.js
 *
 * Drives the real page in Chromium: loads the sample dataset, checks the
 * numbers on screen against the values Merit.xlsb produced, exercises the
 * HOD budget guard, and confirms the Excel export round-trips.
 * ==========================================================================*/
'use strict';
var path = require('path');
var fs = require('fs');
var os = require('os');
var { chromium } = require('playwright');
var expected = require('./merit-expected.json');

var APP = 'file://' + path.join(__dirname, '..', 'karaneh-system.html');
var failures = [];

function check(label, ok, detail) {
  console.log((ok ? '  PASS  ' : '  FAIL  ') + label + (detail ? '   ' + detail : ''));
  if (!ok) failures.push(label + (detail ? ' — ' + detail : ''));
}

function near(a, b, tol) { return Math.abs(a - b) <= tol; }

(async function () {
  var downloadDir = fs.mkdtempSync(path.join(os.tmpdir(), 'karaneh-'));
  var browser = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  var ctx = await browser.newContext({ acceptDownloads: true });
  var page = await ctx.newPage();

  var pageErrors = [];
  page.on('pageerror', function (e) { pageErrors.push(String(e)); });
  page.on('console', function (m) {
    if (m.type() === 'error') pageErrors.push('console: ' + m.text());
  });

  console.log('\n== BOOT ==');
  await page.goto(APP);
  await page.waitForSelector('.topbar .brand', { timeout: 15000 });
  check('page boots and renders the shell', true);
  check('no page errors on boot', pageErrors.length === 0, pageErrors.join(' | '));

  var navCount = await page.locator('.navitem').count();
  check('all navigation entries render', navCount === 10, navCount + ' items');

  console.log('\n== LOAD SAMPLE DATA ==');
  /* Load through the app's own code path, then confirm it persisted. */
  await page.evaluate(function () {
    var s = window.SAMPLE_DATA, now = new Date().toISOString();
    var st = window.App.state;
    Object.keys(s.config).forEach(function (k) {
      if (st.config[k] !== undefined) st.config[k] = s.config[k];
    });
    st.questionnaires = s.employees.map(function (e, i) {
      var c = JSON.parse(JSON.stringify(e));
      c._key = 'q' + (i + 1); c.importedAt = now;
      c.hodComment = c.hodAdjustment != null ? 'از فایل مرجع' : '';
      /* Conformance replay: employee 4's column-O formula in the workbook
         ignored its own special-impact value. Feed the app what that formula
         actually acted on so the browser can be held to Excel exactly.
         See docs/05-ambiguities.md, AMB-01. */
      if (c.employeeId === '4') { c.specialProject = null; c.specialImpactAmount = 0; }
      return c;
    });
    st.employees = s.employees.map(function (e) {
      return { employeeId: e.employeeId, fullName: e.fullName, division: e.division,
               positionTitle: e.positionTitle, jobLevel: e.jobLevel, employeeStatus: 'Active' };
    });
  });
  await page.evaluate(function () { window.App.save(); window.App.recalc(); window.App.go('dashboard'); });
  await page.waitForTimeout(400);

  var totals = await page.evaluate(function () {
    var t = window.App.result.totals;
    return {
      inScope: t.inScopeCount, eligible: t.eligibleCount, ineligible: t.ineligibleCount,
      overridden: t.overriddenCount, sumFinal: t.sumFinalKaraneh, budget: t.budget,
      remaining: t.remainingBudget, status: t.budgetStatus,
      ineligibleRedist: t.ineligibleRedistribution, hodRedist: t.hodRedistribution,
      sumTotalScore: t.sumTotalScore
    };
  });
  check('100 questionnaire records in scope', totals.inScope === 100, String(totals.inScope));
  check('95 employees clear the threshold', totals.eligible === 95, String(totals.eligible));
  check('5 employees below threshold', totals.ineligible === 5, String(totals.ineligible));
  check('2 HOD overrides carried from the workbook', totals.overridden === 2, String(totals.overridden));
  check('sum of final karaneh equals the budget',
    near(totals.sumFinal, totals.budget, 1e-3), totals.sumFinal.toFixed(4));
  check('budget status is BALANCED', totals.status === 'BALANCED', totals.status);
  check('ineligible redistribution matches Excel J5',
    near(totals.ineligibleRedist, expected.totals.ineligibleRedistribution, 1e-9),
    totals.ineligibleRedist.toFixed(9));
  check('sum of total score equals 10,000',
    near(totals.sumTotalScore, 10000, 1e-6), totals.sumTotalScore.toFixed(6));

  console.log('\n== DASHBOARD RENDERING ==');
  var kpiText = await page.locator('.kpi').allTextContents();
  check('KPI tiles render', kpiText.length >= 8, kpiText.length + ' tiles');
  var topbar = await page.locator('.topbar').textContent();
  check('top bar shows the budget', topbar.indexOf('100.00B') !== -1, topbar.replace(/\s+/g, ' ').trim());
  var divRows = await page.locator('table.grid tbody tr').count();
  check('per-division summary renders', divRows >= 8, divRows + ' divisions');

  console.log('\n== PAYMENT TABLE vs EXCEL ==');
  await page.evaluate(function () { window.App.go('payment'); });
  await page.waitForSelector('table.grid tbody tr');
  var onScreen = await page.evaluate(function () {
    var out = {};
    window.App.result.rows.forEach(function (r) {
      out[r.employeeId] = {
        gradeScore: r.gradeScore, gradeImpact: r.gradeImpact,
        perf: r.performanceContribution, total: r.totalScore,
        alloc: r.initialAllocation, final: r.finalKaraneh
      };
    });
    return out;
  });
  var maxAlloc = 0, maxFinal = 0, worst = null;
  expected.rows.forEach(function (e) {
    var g = onScreen[e.employeeId];
    var da = Math.abs((g.alloc || 0) - (e.initialAllocation || 0));
    var df = Math.abs((g.final || 0) - (e.finalKaraneh || 0));
    if (da > maxAlloc) { maxAlloc = da; worst = e.employeeId; }
    if (df > maxFinal) maxFinal = df;
  });
  check('initial allocation matches Excel for all 100 employees',
    maxAlloc < 1e-3, 'max Δ ' + maxAlloc.toExponential(2) + ' rial (employee ' + worst + ')');
  check('final karaneh matches Excel for all 100 employees',
    maxFinal < 1e-3, 'max Δ ' + maxFinal.toExponential(2) + ' rial');

  var footer = await page.locator('table.grid tfoot').textContent();
  check('payment table footer totals render', /100,000,000,000/.test(footer.replace(/\s+/g, '')),
    footer.replace(/\s+/g, ' ').trim().slice(0, 90));

  console.log('\n== CALCULATION DETAIL PANEL ==');
  await page.evaluate(function () { window.App.showEmployeeDetail('1'); });
  await page.waitForSelector('.modal .calc-step');
  var steps = await page.locator('.modal .calc-step').count();
  var detailText = await page.locator('.modal').textContent();
  check('breakdown shows every calculation step', steps >= 12, steps + ' steps');
  check('breakdown shows the performance score', /3\.50/.test(detailText));
  check('breakdown shows the karaneh number 105', /105\.00/.test(detailText));
  check('breakdown shows the special impact of 300', /300/.test(detailText));
  check('breakdown quotes the source Excel columns', /ستون K/.test(detailText) && /ستون Q/.test(detailText));
  await page.locator('.modal header .close').click();

  console.log('\n== REAL-TIME RECALCULATION ==');
  var before = await page.evaluate(function () { return window.App.result.totals.eligibleCount; });
  await page.evaluate(function () {
    /* Raise the threshold: more people should drop out, budget must still balance. */
    window.App.state.config.minPerformanceThreshold = 3;
    window.App.go('payment');
  });
  await page.evaluate(function () { window.App.state.config.minPerformanceThreshold = 3; });
  var after = await page.evaluate(function () {
    window.App.state.config.minPerformanceThreshold = 3;
    var r = window.KaranehEngine.calculate(
      window.App.result.rows.map(function (x) { return x._input; }), window.App.state.config);
    return { eligible: r.totals.eligibleCount, sum: r.totals.sumFinalKaraneh,
             status: r.totals.budgetStatus };
  });
  check('raising the threshold reduces the eligible population',
    after.eligible < before, before + ' → ' + after.eligible);
  check('budget still reconciles after the change',
    near(after.sum, 100000000000, 1e-3), after.sum.toFixed(2));
  await page.evaluate(function () {
    window.App.state.config.minPerformanceThreshold = 2;
    window.App.go('payment');
  });

  console.log('\n== BUDGET GUARD ON HOD ADJUSTMENT ==');
  var guard = await page.evaluate(function () {
    var E = window.KaranehEngine, App = window.App;
    var ceiling = E.maxAllowedAdjustment(App.result, '3');
    /* Push one person far past the ceiling and confirm the engine flags it. */
    var records = App.result.rows.map(function (x) {
      var c = {}; for (var k in x._input) c[k] = x._input[k];
      if (c.employeeId === '3') { c.hodAdjustment = 9.5e10; c.hodComment = 'تست'; }
      return c;
    });
    var res = E.calculate(records, App.state.config);
    return {
      ceiling: ceiling,
      negatives: res.totals.negativePayoutCount,
      status: res.totals.budgetStatus,
      sum: res.totals.sumFinalKaraneh,
      valid: E.validateBudget(res).ok
    };
  });
  check('a ceiling is computed for the override', guard.ceiling > 0, guard.ceiling.toFixed(0));
  check('an oversized override drives peers negative', guard.negatives > 0, guard.negatives + ' employees');
  check('budget status flips to INVALID', guard.status === 'INVALID', guard.status);
  check('validateBudget rejects it', guard.valid === false);
  check('total still reconciles to the budget (it is the split that breaks)',
    near(guard.sum, 100000000000, 1e-2), guard.sum.toFixed(2));

  console.log('\n== HOD COMMENT IS MANDATORY ==');
  await page.evaluate(function () { window.App.go('hod'); });
  await page.waitForSelector('table.grid tbody tr');
  await page.evaluate(function () {
    var btns = Array.prototype.slice.call(document.querySelectorAll('table.grid tbody tr button'));
    for (var i = 0; i < btns.length; i++) {
      if (btns[i].textContent.indexOf('تعیین') !== -1) { btns[i].click(); return; }
    }
  });
  await page.waitForSelector('.modal input[type="number"]');
  await page.locator('.modal input[type="number"]').fill('1000000000');
  await page.locator('.modal footer button.primary').click();
  await page.waitForTimeout(300);
  var stillOpen = await page.locator('.modal').count();
  var toastText = await page.locator('.toast').allTextContents();
  check('saving without a comment is blocked', stillOpen > 0);
  check('the user is told a comment is required',
    toastText.join(' ').indexOf('توضیح') !== -1, toastText.join(' | '));

  await page.locator('.modal textarea').fill('اثرگذاری ویژه در پروژه انتقال داده');
  await page.locator('.modal footer button.primary').click();
  await page.waitForTimeout(400);
  var afterHod = await page.evaluate(function () {
    var t = window.App.result.totals;
    return { overridden: t.overriddenCount, sum: t.sumFinalKaraneh, status: t.budgetStatus,
             audits: window.App.state.auditLog.length };
  });
  check('the override is applied once a comment is given',
    afterHod.overridden === 3, String(afterHod.overridden));
  check('budget still reconciles after the override',
    near(afterHod.sum, 100000000000, 1e-2), afterHod.sum.toFixed(2));
  check('the change is written to the audit log', afterHod.audits > 0, afterHod.audits + ' entries');

  console.log('\n== VALIDATION CENTER ==');
  await page.evaluate(function () { window.App.go('validation'); });
  await page.waitForSelector('.kpi');
  var vText = await page.locator('#main').textContent();
  check('validation center renders', vText.indexOf('مرکز اعتبارسنجی') !== -1);
  var belowThreshold = await page.evaluate(function () {
    return window.App.validation.filter(function (i) { return i.code === 'BELOW_THRESHOLD'; }).length;
  });
  check('the 5 below-threshold employees are reported', belowThreshold === 5, String(belowThreshold));

  console.log('\n== EXCEL EXPORT ==');
  var downloadPromise = page.waitForEvent('download', { timeout: 20000 });
  await page.evaluate(function () { window.App.go('reports'); });
  await page.waitForSelector('#main button.primary');
  await page.locator('#main button.primary').first().click();
  var download = await downloadPromise;
  var file = path.join(downloadDir, download.suggestedFilename());
  await download.saveAs(file);
  var size = fs.statSync(file).size;
  check('an .xlsx file is produced', /\.xlsx$/.test(file) && size > 5000,
    download.suggestedFilename() + ' (' + (size / 1024).toFixed(0) + ' KB)');

  var XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
  var wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
  var want = ['Employee Master', 'Consolidated Questionnaire', 'روش پرداخت کارانه',
              'Validation Report', 'Summary', 'Configuration', 'Audit Log'];
  check('all seven required sheets are present',
    want.every(function (n) { return wb.SheetNames.indexOf(n) !== -1; }),
    wb.SheetNames.join(' | '));

  var pay = XLSX.utils.sheet_to_json(wb.Sheets['روش پرداخت کارانه'], { header: 1 });
  check('payment sheet keeps the reference column order',
    pay[0][0] === 'شماره پرسنلی' && pay[0][12] === 'دریافتی قبل از تغییرات معاون بخش' &&
    pay[0][16] === 'HOD Comment', pay[0].slice(0, 3).join(' , '));
  check('payment sheet carries one row per employee plus a total row',
    pay.length >= 102, pay.length + ' rows');
  /* SheetJS drops both of these on read, so inspect the archive itself.
     The export is written uncompressed, so the worksheet XML is verbatim. */
  var rawXlsx = fs.readFileSync(file).toString('latin1');
  check('exported sheets are right-to-left', /rightToLeft="1"/.test(rawXlsx));
  check('exported sheets have frozen panes',
    /<pane [^>]*state="frozen"/.test(rawXlsx),
    (rawXlsx.match(/state="frozen"/g) || []).length + ' frozen sheets');
  check('exported sheets carry an autofilter', /<autoFilter/.test(rawXlsx));
  check('rial columns use a thousand separator', /#,##0/.test(rawXlsx));

  var exportedSum = 0;
  pay.slice(1).forEach(function (row) {
    if (row[0] === 'جمع' || !row[0]) return;
    exportedSum += Number(row[15]) || 0;
  });
  check('exported payouts sum to the budget',
    near(exportedSum, 100000000000, 1e-2), exportedSum.toFixed(2));

  console.log('\n== PERSISTENCE ==');
  await page.reload();
  await page.waitForSelector('.topbar .brand');
  await page.waitForTimeout(700);
  var reloaded = await page.evaluate(function () {
    return { q: window.App.state.questionnaires.length,
             overridden: window.App.result.totals.overriddenCount,
             sum: window.App.result.totals.sumFinalKaraneh };
  });
  check('data survives a reload', reloaded.q === 100, reloaded.q + ' records');
  check('HOD overrides survive a reload', reloaded.overridden === 3, String(reloaded.overridden));
  check('totals are identical after reload',
    near(reloaded.sum, 100000000000, 1e-2), reloaded.sum.toFixed(2));

  check('no uncaught page errors during the whole run',
    pageErrors.length === 0, pageErrors.slice(0, 3).join(' | '));

  await browser.close();
  fs.rmSync(downloadDir, { recursive: true, force: true });

  console.log('\n' + '='.repeat(74));
  if (failures.length) {
    console.log('E2E RESULT: FAILED (' + failures.length + ')');
    failures.forEach(function (f) { console.log('  • ' + f); });
    process.exit(1);
  }
  console.log('E2E RESULT: PASS — the built application works end to end.');
  console.log('='.repeat(74));
}()).catch(function (e) {
  console.error('\nE2E harness error:', e);
  process.exit(1);
});
