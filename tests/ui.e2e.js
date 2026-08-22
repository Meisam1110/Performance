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

var APP = 'file://' + path.join(__dirname, '..', 'karaneh-hr.html');
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
  await page.waitForSelector('.brandbar .title', { timeout: 15000 });
  check('page boots and renders the shell', true);
  check('no page errors on boot', pageErrors.length === 0, pageErrors.join(' | '));

  var navCount = await page.locator('.navitem').count();
  var phaseCount = await page.locator('.navphase-tag').count();
  check('navigation is split into two phases', phaseCount === 2, phaseCount + ' phase headers');
  check('navigation sits in a horizontal bar',
    await page.locator('.navbar .navitem').count() > 0);
  check('all navigation entries render for the admin role', navCount === 12, navCount + ' items');
  var themed = await page.evaluate(function () {
    var before = document.documentElement.getAttribute('data-theme');
    document.getElementById('themeBtn').click();
    var after = document.documentElement.getAttribute('data-theme');
    document.getElementById('themeBtn').click();
    return { before: before, after: after,
             restored: document.documentElement.getAttribute('data-theme') };
  });
  check('the theme toggle switches modes',
    themed.before === 'light' && themed.after === 'dark' && themed.restored === 'light',
    themed.before + ' → ' + themed.after + ' → ' + themed.restored);
  var brand = await page.evaluate(function () {
    return getComputedStyle(document.documentElement).getPropertyValue('--brand').trim();
  });
  check('the MTN Irancell brand colour is in use', brand.toUpperCase() === '#FFCC00', brand);
  var fontOk = await page.evaluate(function () {
    return document.fonts ? document.fonts.check('700 14px "MTN Irancell"') : true;
  });
  check('the brand font is embedded and loaded', fontOk === true, String(fontOk));
  var logo = await page.locator('.brandbar .logo svg').count();
  check('the brand mark renders', logo === 1);

  var shell = await page.evaluate(function () {
    var nav = Array.prototype.map.call(document.querySelectorAll('.navitem'),
      function (n) { return n.textContent.replace(/[0-9🔒]/g, '').trim(); });
    var svg = document.querySelector('.brandbar .logo svg');
    var parts = Array.prototype.map.call(svg.querySelectorAll('rect, text'), function (n) {
      var r = n.getBoundingClientRect();
      return { tag: n.tagName, txt: n.textContent, left: Math.round(r.left), width: Math.round(r.width) };
    });
    return { first: nav[0], parts: parts, box: Math.round(svg.getBoundingClientRect().width) };
  });
  check('the guide is the first tab', /راهنما/.test(shell.first), shell.first);
  /* The wordmark is anchored inside an RTL text element, where `start` is the
     right edge; anchored the other way it rendered off the viewBox. */
  check('the brand mark draws both of its parts inside the box',
    shell.parts.length === 3 && shell.parts.every(function (p) {
      return p.width > 8 && p.left >= 0;
    }), JSON.stringify(shell.parts));
  check('the wordmark sits beside the badge, not on top of it',
    shell.parts[2].left >= shell.parts[0].left + shell.parts[0].width,
    shell.parts[0].left + '+' + shell.parts[0].width + ' vs ' + shell.parts[2].left);

  console.log('\n== LOAD SAMPLE DATA ==');
  /* The workbook data goes in exactly as it stands — no fix-ups. The
     special-impact gate is what makes it reproduce Excel. */
  await page.evaluate(function () {
    var s = window.SAMPLE_DATA, now = new Date().toISOString();
    var st = window.App.state;
    Object.keys(s.config).forEach(function (k) {
      if (st.config[k] === undefined) return;
      if (k === 'gradeMap') {
        Object.keys(s.config.gradeMap).forEach(function (jl) {
          st.config.gradeMap[jl] = s.config.gradeMap[jl];
        });
        return;
      }
      st.config[k] = s.config[k];
    });
    st.questionnaires = s.employees.map(function (e, i) {
      var c = JSON.parse(JSON.stringify(e));
      c._key = 'q' + (i + 1); c.importedAt = now;
      c.hodComment = c.hodAdjustment != null ? 'از فایل مرجع' : '';
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
  var strip = await page.locator('.strip').textContent();
  check('budget strip shows the budget', /100,000,000,000/.test(strip.replace(/\s+/g, '')),
    strip.replace(/\s+/g, ' ').trim().slice(0, 90));
  var railPct = await page.locator('.ringwrap .pct').textContent();
  check('progress ring reports a percentage', /\d+٪/.test(railPct), railPct.trim());
  var stepCount = await page.locator('.step').count();
  check('progress rail lists the process steps', stepCount === 8, stepCount + ' steps');
  await page.waitForSelector('.chart svg', { timeout: 5000 });
  var chartCount = await page.locator('.chart svg').count();
  check('dashboard renders its charts', chartCount >= 5, chartCount + ' charts');
  var marks = await page.locator('.chart .chart-mark').count();
  check('charts draw real marks', marks > 20, marks + ' marks');
  var legendItems = await page.locator('.chart-legend-item').count();
  check('the status chart carries a legend', legendItems >= 3, legendItems + ' legend entries');
  var unifiedRows = await page.locator('table.grid tbody tr').count();
  check('unified table renders alongside the charts', unifiedRows > 20, unifiedRows + ' rows');
  var facets = await page.locator('.table-toolbar select').count();
  check('unified table exposes filters', facets >= 6, facets + ' filters');

  /* Filtering the table must not change what the charts describe. */
  await page.locator('.table-toolbar select').first().selectOption({ index: 1 });
  await page.waitForTimeout(250);
  var filtered = await page.locator('table.grid tbody tr').count();
  check('a filter narrows the table', filtered < unifiedRows && filtered > 0,
    unifiedRows + ' → ' + filtered);
  await page.locator('.table-toolbar select').first().selectOption({ index: 0 });
  await page.waitForTimeout(200);

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

  var footer = await page.locator('table.grid tfoot').last().textContent();
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

  console.log('\n== SPECIAL-IMPACT GATE ==');
  var gate = await page.evaluate(function () {
    var cfg = window.App.state.config;
    var rows = window.App.result.rows;
    var flagged = rows.filter(function (r) { return r.specialProject; });
    return {
      floor: cfg.specialImpactMinScore,
      flagged: flagged.length,
      blocked: rows.filter(function (r) { return r.specialImpactBlocked; }).length,
      blockedIds: rows.filter(function (r) { return r.specialImpactBlocked; })
        .map(function (r) { return r.employeeId; }),
      /* the blocked employee must have entered 300 but applied 0 */
      entered: (rows.filter(function (r) { return r.specialImpactBlocked; })[0] || {}).specialImpactEntered,
      applied: (rows.filter(function (r) { return r.specialImpactBlocked; })[0] || {}).specialImpactValue
    };
  });
  check('gate floor is configured', gate.floor === 100, String(gate.floor));
  check('4 employees are flagged for special impact', gate.flagged === 4, String(gate.flagged));
  check('the sub-threshold employee is blocked', gate.blocked === 1 && gate.blockedIds[0] === '4',
    gate.blockedIds.join(','));
  check('blocked employee records entered 300 but applied 0',
    gate.entered === 300 && gate.applied === 0, gate.entered + ' → ' + gate.applied);

  await page.evaluate(function () { window.App.go('questionnaires'); });
  await page.waitForSelector('table.grid tbody tr');
  var lockUi = await page.evaluate(function () {
    /* find the blocked employee's row and read its special-impact cell */
    var grid = window.App.grids.questionnaires;
    grid.state.filter = '4';
    grid.render();
    var locks = document.querySelectorAll('table.grid tbody .locked-note').length;
    grid.state.filter = '';
    grid.render();
    return locks;
  });
  check('the UI locks the special-impact control below the floor', lockUi > 0, lockUi + ' locked cells');

  console.log('\n== QUESTIONNAIRE DESIGNER ==');
  await page.evaluate(function () { window.App.go('designer'); });
  await page.waitForSelector('.qrow');
  var qrows = await page.locator('.qrow').count();
  check('designer lists every question', qrows === 5, qrows + ' questions');
  var sigShown = await page.locator('#main .mono').first().textContent();
  check('designer shows the template signature', /^[0-9a-f]{8}$/.test(sigShown.trim()), sigShown.trim());

  var weighted = await page.evaluate(function () {
    var before = window.App.result.totals.sumRawCoefficient;
    window.App.state.config.questions[0].weight = 3;
    window.App.recalc();
    var after = window.App.result.totals.sumRawCoefficient;
    window.App.state.config.questions[0].weight = 1;
    window.App.recalc();
    return { before: before, after: after, restored: window.App.result.totals.sumRawCoefficient };
  });
  check('changing a question weight moves the score',
    Math.abs(weighted.after - weighted.before) > 1,
    weighted.before.toFixed(2) + ' → ' + weighted.after.toFixed(2));
  check('restoring the weight restores the score',
    Math.abs(weighted.restored - weighted.before) < 1e-9);

  console.log('\n== GRADE TABLE AND IMPACT FACTOR ==');
  await page.evaluate(function () { window.App.go('settings'); });
  await page.waitForSelector('#main table.grid');
  var gradeUi = await page.evaluate(function () {
    var cfg = window.App.state.config;
    /* the grade table is the first grid on the settings page */
    var rows = Array.prototype.slice.call(
      document.querySelectorAll('#gradeTable tbody tr'));
    var levels = rows.map(function (tr) { return tr.children[0].textContent.trim(); });
    return {
      map: cfg.gradeMap,
      levelsShown: levels,
      factorControls: document.querySelectorAll('#main input[type="range"]').length,
      has2H: levels.some(function (l) { return l.indexOf('2H') === 0; })
    };
  });
  check('grade table lists 2H', gradeUi.has2H, gradeUi.levelsShown.join(' , '));
  check('2H is configured at 225', gradeUi.map['2H'] === 225, String(gradeUi.map['2H']));
  check('grade ladder is shown in rank order',
    gradeUi.levelsShown.join(',') === '0,1,2,2H,3,3H,4', gradeUi.levelsShown.join(','));
  check('the impact factor has a slider control', gradeUi.factorControls === 1,
    gradeUi.factorControls + ' sliders');

  /* Moving the slider must preview without committing. */
  var previewOnly = await page.evaluate(function () {
    var before = window.App.state.config.gradeImpactFactor;
    var slider = document.querySelector('#main input[type="range"]');
    slider.value = '0.5';
    slider.dispatchEvent(new Event('input', { bubbles: true }));
    return { before: before, after: window.App.state.config.gradeImpactFactor,
             previewText: document.querySelector('#main .alert') ?
               document.querySelector('#main .alert').textContent : '' };
  });
  check('moving the slider previews without committing',
    previewOnly.before === 0 && previewOnly.after === 0, 'config still ' + previewOnly.after);
  check('the preview states what the change would move',
    /منتقل می‌شود/.test(previewOnly.previewText), previewOnly.previewText.slice(0, 70).trim());

  /* Committing changes the distribution and still reconciles. */
  var committed = await page.evaluate(function () {
    var slider = document.querySelector('#main input[type="range"]');
    slider.value = '0.5';
    slider.dispatchEvent(new Event('input', { bubbles: true }));
    var apply = Array.prototype.slice.call(document.querySelectorAll('#main button'))
      .filter(function (b) { return b.textContent.trim() === 'اعمال'; })[0];
    apply.click();
    return {
      factor: window.App.state.config.gradeImpactFactor,
      sum: window.App.result.totals.sumFinalKaraneh,
      gradeImpacted: window.App.result.rows.filter(function (r) {
        return r.inScope && r.gradeImpact > 0;
      }).length
    };
  });
  check('applying the factor commits it', committed.factor === 0.5, String(committed.factor));
  check('grade now contributes to the score', committed.gradeImpacted > 90,
    committed.gradeImpacted + ' employees');
  check('budget still reconciles with grade switched on',
    Math.abs(committed.sum - 100000000000) < 1e-2, committed.sum.toFixed(2));

  var restored = await page.evaluate(function () {
    window.App.state.config.gradeImpactFactor = 0;
    window.App.recalc();
    return window.App.result.totals.sumFinalKaraneh;
  });
  check('restoring the factor to zero reconciles',
    Math.abs(restored - 100000000000) < 1e-2, restored.toFixed(2));

  /* An unmapped level must be offered a one-click fix. */
  var unmapped = await page.evaluate(function () {
    window.App.state.questionnaires[0].jobLevel = '5H';
    window.App.state.employees[0].jobLevel = '5H';
    window.App.recalc();
    window.App.go('settings');
    var errs = window.App.validation.filter(function (i) { return i.code === 'UNMAPPED_JL'; }).length;
    var fixBtn = Array.prototype.slice.call(document.querySelectorAll('#main button'))
      .filter(function (b) { return b.textContent.indexOf('افزودن 5H') !== -1; })[0];
    var had = !!fixBtn;
    if (fixBtn) fixBtn.click();
    return { errs: errs, offered: had, added: window.App.state.config.gradeMap['5H'] };
  });
  check('an unmapped job level is reported', unmapped.errs > 0, unmapped.errs + ' issues');
  check('the settings page offers a one-click fix', unmapped.offered);
  check('the suggested score seats it above the top level',
    unmapped.added === 400, String(unmapped.added));

  await page.evaluate(function () {
    delete window.App.state.config.gradeMap['5H'];
    window.App.state.questionnaires[0].jobLevel = '3H';
    window.App.state.employees[0].jobLevel = '3H';
    window.App.recalc();
  });

  console.log('\n== TEMPLATE DOWNLOAD AND RE-IMPORT ==');
  /* Navigate explicitly rather than relying on where the previous block left
     the app — otherwise reordering blocks silently breaks this one. */
  await page.evaluate(function () {
    window.App.state.employees = window.App.state.employees.slice(0, 6);
    window.App.recalc();
    window.App.go('designer');
  });
  await page.waitForSelector('.qrow');
  var tplDownload = page.waitForEvent('download', { timeout: 20000 });
  await page.evaluate(function () {
    document.querySelectorAll('#main button').forEach(function (b) {
      if (b.textContent.trim() === 'تمپلیت با فهرست پرسنل') b.click();
    });
  });
  var tplFile = path.join(downloadDir, (await tplDownload).suggestedFilename());
  await (await tplDownload).saveAs(tplFile);
  check('questionnaire template downloads', fs.existsSync(tplFile),
    path.basename(tplFile) + ' (' + (fs.statSync(tplFile).size / 1024).toFixed(0) + ' KB)');

  var XLSX0 = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
  var tplWb = XLSX0.read(fs.readFileSync(tplFile), { type: 'buffer' });
  check('template carries the questionnaire sheet',
    tplWb.SheetNames.indexOf('پرسشنامه کارانه تیمی') !== -1, tplWb.SheetNames.join(' | '));
  check('template carries a hidden signature sheet',
    tplWb.SheetNames.indexOf('_Template') !== -1);
  var tplRaw = fs.readFileSync(tplFile).toString('latin1');
  check('template answer cells are dropdowns', /<dataValidation type="list"/.test(tplRaw),
    (tplRaw.match(/<dataValidation /g) || []).length + ' validations');
  check('template freezes its header', /state="frozen"/.test(tplRaw));

  var tplRows = XLSX0.utils.sheet_to_json(tplWb.Sheets['پرسشنامه کارانه تیمی'],
    { header: 1, defval: null, blankrows: false });
  var codeRow = tplRows.filter(function (r) { return r[0] === 'شماره پرسنلی'; })[0];
  check('template header row uses short question codes',
    codeRow && codeRow.indexOf('Q1') !== -1 && codeRow.indexOf('Q4') !== -1,
    codeRow ? codeRow.slice(5, 11).join(',') : 'not found');

  /* A template cut from a different design must be refused. */
  var mismatch = await page.evaluate(function () {
    var cfg = JSON.parse(JSON.stringify(window.App.state.config));
    cfg.questions.push({ id: 'q9', text: 'سؤال تازه', weight: 1, scored: true });
    var wb = window.Templates.buildQuestionnaireTemplate(cfg, [], { XLSX: window.XLSX });
    var bytes = window.XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: false });
    var back = window.XLSX.read(bytes, { type: 'array' });
    var v = window.Templates.verifyAgainstTemplate(back, window.App.state.config, window.XLSX);
    var same = window.Templates.verifyAgainstTemplate(
      window.XLSX.read(window.XLSX.write(
        window.Templates.buildQuestionnaireTemplate(window.App.state.config, [], { XLSX: window.XLSX }),
        { type: 'array', bookType: 'xlsx', compression: false }), { type: 'array' }),
      window.App.state.config, window.XLSX);
    return { foreign: v, own: same };
  });
  check('a template from a different design is rejected',
    mismatch.foreign.ok === false && mismatch.foreign.level === 'mismatch',
    mismatch.foreign.level);
  check('the rejection names the mismatch',
    /مجموعه سؤالات/.test(mismatch.foreign.problems.join(' ')));
  check('a template from the current design is accepted',
    mismatch.own.ok === true && mismatch.own.level === 'match', mismatch.own.level);

  console.log('\n== ROLE-BASED ACCESS ==');
  var roleTest = await page.evaluate(function () {
    var divisions = {};
    window.App.result.rows.forEach(function (r) { if (r.division) divisions[r.division] = 1; });
    var first = Object.keys(divisions)[0];
    window.App.state.role = 'hod';
    window.App.state.hodScope = [first];
    window.App.recalc();
    window.App.go('dashboard');
    var navItems = document.querySelectorAll('.navitem').length;
    var canSettings = window.App.go && (function () {
      window.App.go('settings');
      return window.App.view === 'settings';
    }());
    window.App.state.role = 'admin';
    window.App.state.hodScope = [];
    window.App.recalc();
    return { division: first, navItems: navItems, reachedSettings: canSettings };
  });
  check('division head sees a reduced menu', roleTest.navItems < 11, roleTest.navItems + ' items');
  check('division head cannot reach settings', roleTest.reachedSettings === false);

  console.log('\n== PHASE GATE ==');
  var gateTest = await page.evaluate(function () {
    /* Break one answer so phase 1 has a blocking error. */
    var q = window.App.state.questionnaires[0];
    var keep = q.q1;
    q.q1 = 'یک پاسخ نامعتبر';
    window.App.recalc();
    var blocked = window.App.validation.filter(function (i) {
      return i.severity === 'err' && i.stage === 'data';
    }).length;
    window.App.go('hod');
    var lockedView = document.querySelector('#main .alert.err') !== null;
    q.q1 = keep;
    window.App.recalc();
    return { blocked: blocked, lockedView: lockedView, reopened: window.App.view };
  });
  check('an invalid answer blocks phase 1', gateTest.blocked > 0, gateTest.blocked + ' blockers');
  check('HOD adjustments are locked while phase 1 is incomplete', gateTest.lockedView);

  console.log('\n== PAYROLL-FORMAT EXPORT ==');
  await page.evaluate(function () { window.App.go('dashboard'); });
  await page.waitForSelector('#main button');
  var payrollDl = page.waitForEvent('download', { timeout: 20000 });
  await page.evaluate(function () {
    document.querySelectorAll('#main button').forEach(function (b) {
      if (b.textContent.trim() === 'خروجی کارانه') b.click();
    });
  });
  var payrollFile = path.join(downloadDir, (await payrollDl).suggestedFilename());
  await (await payrollDl).saveAs(payrollFile);
  var pwb = XLSX0.read(fs.readFileSync(payrollFile), { type: 'buffer' });
  var prows = XLSX0.utils.sheet_to_json(pwb.Sheets[pwb.SheetNames[0]], { header: 1, defval: null });
  var expectedHeaders = ['Emp No', 'Emp Status', 'First Name', 'Last Name'];
  check('payroll export uses the payroll team column layout',
    expectedHeaders.every(function (h, i) { return prows[0][i] === h; }),
    prows[0].slice(0, 4).join(' , '));
  var finalIdx = prows[0].indexOf('Final Karaneh');
  check('payroll export has a Final Karaneh column', finalIdx > 0, 'index ' + finalIdx);
  var paySum = 0, payRows = 0;
  prows.slice(1).forEach(function (r) {
    if (!r || !r[0] || r[0] === 'جمع') return;
    payRows++;
    paySum += Number(r[finalIdx]) || 0;
  });
  check('payroll export carries every employee', payRows === 100, payRows + ' rows');
  check('payroll amounts sum to exactly the budget, in whole rial',
    paySum === 100000000000, paySum.toLocaleString('en-US'));

  console.log('\n== EXPORT BY MANAGEMENT LEVEL ==');
  var byMgr = await page.evaluate(function () {
    /* Give the master records managers so the split has something to group on. */
    window.App.state.employees.forEach(function (e, i) {
      e.directManager = 'مدیر ' + (i % 3 + 1);
      e.managerLevel1 = 'معاون ' + (i % 2 + 1);
    });
    window.App.recalc();
    return window.App.state.employees.length;
  });
  var mgrDl = page.waitForEvent('download', { timeout: 20000 });
  await page.evaluate(function () { window.__runManagerExport(); });
  var mgrFile = path.join(downloadDir, (await mgrDl).suggestedFilename());
  await (await mgrDl).saveAs(mgrFile);
  var mwb = XLSX0.read(fs.readFileSync(mgrFile), { type: 'buffer' });
  check('per-manager workbook opens with an index sheet',
    mwb.SheetNames[0] === 'فهرست', mwb.SheetNames.join(' | '));
  check('per-manager workbook has one sheet per group',
    mwb.SheetNames.length >= 4 && mwb.SheetNames.indexOf('مدیر 1') !== -1 &&
    mwb.SheetNames.indexOf('مدیر 3') !== -1,
    mwb.SheetNames.length + ' sheets: ' + mwb.SheetNames.join(' | '));
  var mgrSum = 0;
  mwb.SheetNames.slice(1).forEach(function (n) {
    var rs = XLSX0.utils.sheet_to_json(mwb.Sheets[n], { header: 1, defval: null });
    var fi = rs[0].indexOf('Final Karaneh');
    rs.slice(1).forEach(function (r) {
      if (r && r[0] && r[0] !== 'جمع') mgrSum += Number(r[fi]) || 0;
    });
  });
  check('per-manager sheets sum to exactly the budget',
    mgrSum === 100000000000, mgrSum.toLocaleString('en-US'));
  var idxRows = XLSX0.utils.sheet_to_json(mwb.Sheets['فهرست'], { header: 1, defval: null });
  var idxTotal = idxRows.filter(function (r) { return r && r[0] === 'جمع کل'; })[0];
  check('index sheet totals every group', idxTotal && Math.abs(idxTotal[2] - 100000000000) < 200,
    idxTotal ? Number(idxTotal[2]).toLocaleString('en-US') : 'missing');

  console.log('\n== EXCEL EXPORT ==');
  await page.evaluate(function () { window.App.go('reports'); });
  await page.waitForSelector('#main button.primary');
  var downloadPromise = page.waitForEvent('download', { timeout: 20000 });
  await page.evaluate(function () {
    document.querySelectorAll('#main button').forEach(function (b) {
      if (b.textContent.trim() === 'خروجی کامل Excel') b.click();
    });
  });
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

  console.log('\n== BARS SCALE IN THE TEMPLATE ==');
  var barsSheet = XLSX0.utils.sheet_to_json(tplWb.Sheets['BARS'], { header: 1, defval: null });
  var barsHeader = barsSheet.filter(function (r) { return r && r[0] === 'حوزه'; })[0];
  check('the template carries a BARS sheet', !!tplWb.Sheets['BARS'],
    tplWb.SheetNames.join(' | '));
  check('BARS lists a level column per answer option',
    barsHeader && barsHeader.length === 7, barsHeader ? barsHeader.length + ' columns' : 'missing');
  var barsRows = barsSheet.filter(function (r) {
    return r && r[0] && ['عملکرد / کارایی', 'رفتار شغلی', 'خروجی کار', 'اثرگذاری'].indexOf(r[0]) !== -1;
  });
  check('BARS carries a row per domain', barsRows.length === 4,
    barsRows.map(function (r) { return r[0]; }).join(' , '));
  check('each BARS cell holds a labelled behaviour',
    barsRows.every(function (r) {
      return r.slice(2, 7).every(function (c) { return c && /^\(/.test(c) && c.length > 40; });
    }));
  check('BARS anchor cells are set to wrap',
    /wrapText="1"/.test(tplRaw), 'wrap style present');

  console.log('\n== SPECIAL SCORE STEPS OF 50 ==');
  var stepInfo = await page.evaluate(function () {
    var cfg = window.App.state.config;
    window.App.go('questionnaires');
    var grid = window.App.grids.questionnaires;
    grid.state.filter = '1';
    grid.render();
    var opts = [];
    document.querySelectorAll('table.grid tbody select').forEach(function (sel) {
      var vals = Array.prototype.map.call(sel.options, function (o) { return o.value; });
      if (vals.indexOf('300') !== -1 && vals.indexOf('50') !== -1) opts = vals;
    });
    grid.state.filter = '';
    grid.render();
    return { step: cfg.specialImpactStep, options: opts,
             snapped: [25, 74, 130].map(function (v) {
               return window.KaranehEngine.snapToStep(v, cfg);
             }) };
  });
  check('the special score step is 50', stepInfo.step === 50, String(stepInfo.step));
  check('the picker offers only multiples of 50',
    stepInfo.options.length > 1 &&
    stepInfo.options.filter(function (v) { return v; })
      .every(function (v) { return Number(v) % 50 === 0; }),
    stepInfo.options.join(','));
  check('off-step amounts snap onto the scale',
    stepInfo.snapped.join(',') === '50,50,150', stepInfo.snapped.join(','));

  console.log('\n== DELIVERY AND EMAIL ==');
  var mailState = await page.evaluate(function () {
    window.App.state.mail = {
      performance: 'performance@mtnirancell.ir',
      compensation: 'cnb@mtnirancell.ir',
      cc: 'hrops@mtnirancell.ir'
    };
    window.App.save();
    window.App.go('reports');
    var chips = Array.prototype.map.call(document.querySelectorAll('#main .chip'),
      function (c) { return c.textContent; }).join(' | ');
    return { chips: chips };
  });
  check('reports names both recipient teams',
    /performance@mtnirancell\.ir/.test(mailState.chips) &&
    /cnb@mtnirancell\.ir/.test(mailState.chips),
    mailState.chips.slice(0, 90));

  var pkgDl = page.waitForEvent('download', { timeout: 20000 });
  await page.evaluate(function () {
    document.querySelectorAll('#main button').forEach(function (b) {
      if (b.textContent.trim() === 'دریافت و ارسال بستهٔ نهایی') b.click();
    });
  });
  var perfFile = path.join(downloadDir, (await pkgDl).suggestedFilename());
  await (await pkgDl).saveAs(perfFile);
  check('the performance file is produced', /Performance/.test(path.basename(perfFile)),
    path.basename(perfFile));

  var perfWb = XLSX0.read(fs.readFileSync(perfFile), { type: 'buffer' });
  var perfRows = XLSX0.utils.sheet_to_json(perfWb.Sheets['Performance'], { header: 1, defval: null });
  check('the performance file carries the BARS sheet',
    perfWb.SheetNames.indexOf('BARS') !== -1, perfWb.SheetNames.join(' | '));
  check('the performance file records the chosen anchor per question',
    perfRows[0].filter(function (h) { return /— سطح$/.test(h || ''); }).length === 4,
    perfRows[0].filter(function (h) { return /— سطح$/.test(h || ''); }).join(' , '));
  check('the performance file contains no rial amount',
    perfRows[0].every(function (h) { return !/ریال|کارانه نهایی|دریافتی/.test(h || ''); }),
    perfRows[0].join(' , ').slice(0, 80));

  await page.waitForTimeout(900);
  var draftInfo = await page.evaluate(function () {
    var links = Array.prototype.map.call(document.querySelectorAll('.modal a.btn'),
      function (a) { return a.getAttribute('href') || ''; });
    return { links: links.filter(function (h) { return h.indexOf('mailto:') === 0; }) };
  });
  check('two addressed drafts are prepared', draftInfo.links.length === 2,
    draftInfo.links.length + ' mailto links');
  check('the performance draft is addressed to the performance team',
    draftInfo.links.some(function (h) { return h.indexOf('performance@mtnirancell.ir') !== -1; }));
  check('the compensation draft is addressed to C&B',
    draftInfo.links.some(function (h) { return h.indexOf('cnb@mtnirancell.ir') !== -1; }));
  check('both drafts carry the CC',
    draftInfo.links.every(function (h) { return h.indexOf('hrops%40mtnirancell.ir') !== -1; }));

  await page.evaluate(function () {
    var c = document.querySelector('.modal header .close'); if (c) c.click();
  });

  console.log('\n== DOWNLOAD TRAY ==');
  var tray = await page.evaluate(function () {
    var items = document.querySelectorAll('#dlTray .dl-item');
    return {
      count: items.length,
      hrefs: Array.prototype.map.call(items, function (a) {
        return (a.getAttribute('href') || '').slice(0, 5);
      }),
      hasDownloadAttr: Array.prototype.every.call(items, function (a) { return a.hasAttribute('download'); })
    };
  });
  check('every produced file stays clickable in the tray', tray.count >= 2, tray.count + ' files');
  check('tray entries are real blob links',
    tray.hrefs.every(function (h) { return h === 'blob:'; }), tray.hrefs.join(','));
  check('tray entries carry a download filename', tray.hasDownloadAttr);

  console.log('\n== FILENAMES BROWSERS WILL HONOUR ==');
  var nameTest = await page.evaluate(function () {
    /* Chromium discards a download filename containing non-ASCII and saves it
       as "download" with no extension, so names must transliterate. */
    var probe = document.createElement('a');
    return {
      persian: !!probe,
      tray: (window.App.downloads || []).map(function (d) { return d.filename; })
    };
  });
  check('no produced filename contains non-ASCII',
    nameTest.tray.every(function (n) { return /^[\x20-\x7E]+$/.test(n); }),
    nameTest.tray.slice(0, 2).join(' , '));
  check('every produced filename keeps its extension',
    nameTest.tray.every(function (n) { return /\.(xlsx|csv|json|html)$/.test(n); }),
    nameTest.tray.slice(0, 2).join(' , '));

  console.log('\n== BUDGET AND LEVEL REPORT ON THE PAYMENT SCREEN ==');
  await page.evaluate(function () { window.App.go('payment'); });
  await page.waitForSelector('#main table.grid');
  var payScreen = await page.evaluate(function () {
    var cards = Array.prototype.map.call(document.querySelectorAll('#main .card > h2'),
      function (h) { return h.textContent.trim(); });
    var budgetInput = null;
    document.querySelectorAll('#main input[type="number"]').forEach(function (i) {
      if (String(i.value) === String(window.App.state.config.budget)) budgetInput = i;
    });
    return { cards: cards, hasBudgetInput: !!budgetInput };
  });
  check('the payment screen leads with the per-level report',
    /سطح شغلی/.test(payScreen.cards[0] || ''), payScreen.cards.slice(0, 3).join(' | '));
  check('the budget is editable on the payment screen', payScreen.hasBudgetInput);

  var liveBudget = await page.evaluate(function () {
    var before = window.App.result.totals.sumFinalKaraneh;
    var input = null;
    document.querySelectorAll('#main input[type="number"]').forEach(function (i) {
      if (String(i.value) === String(window.App.state.config.budget)) input = i;
    });
    input.value = '60000000000';
    input.dispatchEvent(new Event('change', { bubbles: true }));
    var after = window.App.result.totals.sumFinalKaraneh;
    /* put it back */
    var input2 = null;
    document.querySelectorAll('#main input[type="number"]').forEach(function (i) {
      if (String(i.value) === '60000000000') input2 = i;
    });
    input2.value = '100000000000';
    input2.dispatchEvent(new Event('change', { bubbles: true }));
    return { before: before, after: after, restored: window.App.result.totals.sumFinalKaraneh };
  });
  check('editing the budget there recalculates immediately',
    Math.abs(liveBudget.after - 60000000000) < 1e-2,
    liveBudget.before.toFixed(0) + ' → ' + liveBudget.after.toFixed(0));
  check('restoring the budget restores the total',
    Math.abs(liveBudget.restored - 100000000000) < 1e-2, liveBudget.restored.toFixed(0));

  var levelReport = await page.evaluate(function () {
    var rows = document.querySelectorAll('#main table.grid tbody tr');
    var first = rows[0] ? Array.prototype.map.call(rows[0].children,
      function (td) { return td.textContent.trim(); }) : [];
    return { rows: rows.length, first: first, chart: document.querySelectorAll('#main .chart svg').length };
  });
  check('the level report lists a row per job level', levelReport.rows > 3,
    levelReport.rows + ' rows');
  check('the level report shows totals and averages per level',
    levelReport.first.length === 10, levelReport.first.slice(0, 6).join(' | '));
  check('the level report is drawn as a chart too', levelReport.chart >= 1,
    levelReport.chart + ' charts');

  console.log('\n== ROSTER GUIDANCE AND BRAND MARK ==');
  var roster = await page.evaluate(function () {
    window.App.go('employees');
    var txt = document.querySelector('#main').textContent;
    return {
      saysPayroll: /از تیم حقوق و دستمزد گرفته‌اید/.test(txt),
      noTray: !/سینی/.test(txt)
    };
  });
  check('the personnel screen says to load the payroll file here', roster.saysPayroll);

  var wording = await page.evaluate(function () {
    window.App.go('help');
    return { noTray: !/سینی/.test(document.querySelector('#main').textContent),
             steps: document.querySelectorAll('#main .card').length };
  });
  check('the guide no longer calls it a tray', wording.noTray);

  var logoCard = await page.evaluate(function () {
    window.App.go('settings');
    var cards = Array.prototype.slice.call(document.querySelectorAll('#main .card'));
    var card = cards.filter(function (c) { return /نشان سازمان/.test(c.textContent); })[0];
    if (!card) return { ok: false };
    return {
      ok: true,
      accepts: (card.querySelector('input[type="file"]') || {}).accept || '',
      hasPreview: !!card.querySelector('svg, img')
    };
  });
  check('settings offers to replace the brand mark with the official file',
    logoCard.ok && /svg/.test(logoCard.accepts), JSON.stringify(logoCard));
  check('the brand-mark card previews what will be shown', logoCard.hasPreview);

  console.log('\n== HANDOVER: HR ISSUES A DIVISION-HEAD FILE ==');
  await page.evaluate(function () {
    /* Earlier blocks trimmed the roster; restore it so the split is meaningful. */
    var s = window.SAMPLE_DATA;
    window.App.state.employees = s.employees.map(function (e) {
      return {
        employeeId: e.employeeId, fullName: e.fullName, division: e.division,
        positionTitle: e.positionTitle, jobLevel: e.jobLevel,
        employeeStatus: 'Active', workingDays: 93
      };
    });
    var mgr = ['مدیر الف', 'مدیر ب', 'مدیر ج'];
    window.App.state.employees.forEach(function (e, i) {
      e.directManager = mgr[i % 3];
      e.managerLevel1 = 'معاون ' + (i % 2 + 1);
    });
    window.App.recalc();
    window.App.go('employees');
  });
  await page.waitForSelector('#main .card');
  var pkgDl = page.waitForEvent('download', { timeout: 25000 });
  await page.evaluate(function () {
    document.querySelectorAll('#main button').forEach(function (b) {
      if (b.textContent.trim() === 'تولید فایل معاون بخش') b.click();
    });
  });
  await page.waitForSelector('.modal .checkline');
  var pickedGroup = await page.evaluate(function () {
    var lines = Array.prototype.slice.call(document.querySelectorAll('.modal .checkline'));
    var first = lines[0].textContent.trim();
    lines.forEach(function (l, i) {
      var cb = l.querySelector('input');
      if (i > 0 && cb.checked) cb.click();
    });
    Array.prototype.slice.call(document.querySelectorAll('.modal footer button'))
      .filter(function (b) { return b.textContent.trim() === 'تولید فایل‌ها'; })[0].click();
    return first;
  });
  var pkgFile = path.join(downloadDir, (await pkgDl).suggestedFilename());
  await (await pkgDl).saveAs(pkgFile);
  check('the handover file is named so the browser keeps it',
    /^[\x20-\x7E]+\.html$/.test(path.basename(pkgFile)), path.basename(pkgFile));
  check('the handover file is a complete application',
    fs.statSync(pkgFile).size > 900 * 1024,
    (fs.statSync(pkgFile).size / 1024).toFixed(0) + ' KB');

  var pkgHtml = fs.readFileSync(pkgFile, 'utf8');
  check('the handover file carries its own payload',
    pkgHtml.indexOf('__KARANEH_PACKAGE__') !== -1);
  check('the payload is parsed, not evaluated as a literal',
    /__KARANEH_PACKAGE__ = JSON\.parse\(/.test(pkgHtml));

  console.log('\n== HANDOVER: THE DIVISION HEAD OPENS IT ==');
  var hodPage = await ctx.newPage();
  var hodErrors = [];
  hodPage.on('pageerror', function (e) { hodErrors.push(String(e)); });
  hodPage.on('console', function (m) {
    if (m.type() === 'error') hodErrors.push('console: ' + m.text());
  });
  await hodPage.goto('file://' + pkgFile);
  await hodPage.waitForSelector('.brandbar .title', { timeout: 15000 });
  await hodPage.waitForTimeout(1200);

  var hod = await hodPage.evaluate(function () {
    return {
      title: document.title,
      role: window.App.state.role,
      label: window.App.state.package.label,
      scope: window.App.state.hodScope,
      employees: window.App.state.employees.length,
      inScope: window.App.result.totals.inScopeCount,
      nav: Array.prototype.map.call(document.querySelectorAll('.navitem'),
        function (n) { return n.textContent.replace(/[0-9🔒]/g, '').trim(); })
    };
  });
  /* The handover file is the same application, not a cut-down one: every
     screen and every setting is there, only the data is narrowed. */
  check('it opens with full access', hod.role === 'admin', hod.role);
  check('it is titled for that division', hod.title.indexOf(hod.label) !== -1, hod.title);
  check('it carries only that division\'s people',
    hod.employees > 0 && hod.employees < 100, hod.employees + ' of 100');
  check('every carried employee is in scope', hod.inScope === hod.employees,
    hod.inScope + ' / ' + hod.employees);
  check('the questionnaire designer is offered',
    hod.nav.some(function (n) { return n.indexOf('طراحی') !== -1; }), hod.nav.join(' , '));
  check('system settings are offered',
    hod.nav.some(function (n) { return n.indexOf('تنظیمات') !== -1; }));
  check('every screen of the cycle is there',
    ['پرسنل', 'ورود پاسخ‌ها', 'پاسخ‌ها', 'اعتبارسنجی', 'پرداخت کارانه', 'تعیین مبلغ', 'خروجی']
      .every(function (want) {
        return hod.nav.some(function (n) { return n.indexOf(want) !== -1; });
      }), hod.nav.join(' , '));
  check('a help section travels with the file',
    hod.nav.some(function (n) { return n.indexOf('راهنما') !== -1; }));
  check('the guide is the first tab', /راهنما/.test(hod.nav[0]), hod.nav[0]);

  var hodHelp = await hodPage.evaluate(function () {
    window.App.go('help');
    return {
      steps: document.querySelectorAll('#main .card').length,
      mentionsPackage: /صادرشده برای/.test(document.querySelector('#main').textContent),
      barsRows: document.querySelectorAll('#main .bars-scale-table tbody tr').length
    };
  });
  check('the help section names the file it belongs to', hodHelp.mentionsPackage);
  check('the help section includes the BARS scale', hodHelp.barsRows === 5,
    hodHelp.barsRows + ' domains');

  /* Carrying HR's whole budget into a ten-person file would read as ten
     billion each; the file opens on that group's own share instead. */
  var seeded = await hodPage.evaluate(function () {
    return { budget: window.App.state.config.budget,
             sum: window.App.result.totals.sumFinalKaraneh,
             source: window.App.state.budgetSource || '' };
  });
  check('the file opens on this group\'s share of the budget, not the whole one',
    seeded.budget > 0 && seeded.budget < 100000000000,
    Number(seeded.budget).toFixed(0));
  check('that seeded budget reconciles on open',
    Math.abs(seeded.sum - seeded.budget) < 1, Number(seeded.sum).toFixed(0));
  check('the budget field says where its opening number came from',
    /سهم/.test(seeded.source), seeded.source);

  var hodBudget = await hodPage.evaluate(function () {
    window.App.go('payment');
    var input = null;
    document.querySelectorAll('#main input[type="number"]').forEach(function (i) {
      if (String(i.value) === String(window.App.state.config.budget)) input = i;
    });
    if (!input) return { ok: false };
    input.value = '8000000000';
    input.dispatchEvent(new Event('change', { bubbles: true }));
    return { ok: true, sum: window.App.result.totals.sumFinalKaraneh,
             status: window.App.result.totals.budgetStatus };
  });
  check('the head can set their own budget in the file', hodBudget.ok);
  check('their budget reconciles exactly',
    Math.abs(hodBudget.sum - 8000000000) < 1e-2, Number(hodBudget.sum).toFixed(0));

  var hodSplit = await hodPage.evaluate(function () {
    window.App.go('employees');
    var labels = Array.prototype.map.call(document.querySelectorAll('#main button'),
      function (b) { return b.textContent.trim(); });
    return { labels: labels,
             hasSplit: labels.some(function (l) { return l.indexOf('به تفکیک مدیر مستقیم') !== -1; }) };
  });
  check('the personnel screen offers the manager split', hodSplit.hasSplit,
    hodSplit.labels.slice(0, 5).join(' , '));

  var splitDl = hodPage.waitForEvent('download', { timeout: 20000 });
  await hodPage.evaluate(function () {
    document.querySelectorAll('#main button').forEach(function (b) {
      if (b.textContent.trim() === 'تمپلیت به تفکیک مدیر مستقیم') b.click();
    });
  });
  await hodPage.waitForSelector('.modal .checkline');
  await hodPage.evaluate(function () {
    var lines = Array.prototype.slice.call(document.querySelectorAll('.modal .checkline'));
    lines.forEach(function (l, i) { var cb = l.querySelector('input'); if (i > 0 && cb.checked) cb.click(); });
    Array.prototype.slice.call(document.querySelectorAll('.modal footer button'))
      .filter(function (b) { return b.textContent.trim() === 'تولید فایل‌ها'; })[0].click();
  });
  var splitFile = path.join(downloadDir, (await splitDl).suggestedFilename());
  await (await splitDl).saveAs(splitFile);
  var splitWb = XLSX0.read(fs.readFileSync(splitFile), { type: 'buffer' });
  check('the split produces a questionnaire per manager',
    splitWb.SheetNames.indexOf('BARS') !== -1, splitWb.SheetNames.join(' | '));
  var splitRows = XLSX0.utils.sheet_to_json(
    splitWb.Sheets['پرسشنامه کارانه تیمی'], { header: 1, defval: null, blankrows: false });
  var splitData = splitRows.filter(function (r) {
    return r[0] && r[0] !== 'شماره پرسنلی' && /^\d/.test(String(r[0]));
  });
  check('the manager file lists only that manager\'s people',
    splitData.length > 0 && splitData.length < hod.employees,
    splitData.length + ' of ' + hod.employees);

  check('the division-head file ran without errors',
    hodErrors.length === 0, hodErrors.slice(0, 2).join(' | '));
  await hodPage.close();

  console.log('\n== PERSISTENCE ==');
  await page.reload();
  await page.waitForSelector('.brandbar .title');
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
