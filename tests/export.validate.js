/* Produce a real export from the built app and leave it on disk so an
 * independent reader (openpyxl) can confirm Excel will understand it.
 *   node tests/export.validate.js <outfile>
 */
'use strict';
var path = require('path'), fs = require('fs');
var { chromium } = require('playwright');
var out = process.argv[2] || path.join(__dirname, '..', 'tmp-export.xlsx');

(async function () {
  var browser = await chromium.launch({
    executablePath: process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium-1194/chrome-linux/chrome'
  });
  var ctx = await browser.newContext({ acceptDownloads: true });
  var page = await ctx.newPage();
  await page.goto('file://' + path.join(__dirname, '..', 'karaneh-hr.html'));
  await page.waitForSelector('.brandbar .title');
  await page.evaluate(function () {
    var s = window.SAMPLE_DATA, now = new Date().toISOString(), st = window.App.state;
    Object.keys(s.config).forEach(function (k) {
      if (st.config[k] === undefined) return;
      if (k === 'gradeMap') {
        /* Merge, so a level the sample does not mention (2H) survives. */
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
      c.hodComment = c.hodAdjustment != null ? 'مقدار وارد شده از فایل مرجع' : '';
      return c;
    });
    st.employees = s.employees.map(function (e) {
      return { employeeId: e.employeeId, fullName: e.fullName, division: e.division,
               positionTitle: e.positionTitle, jobLevel: e.jobLevel, employeeStatus: 'Active',
               employmentType: 'Local', workingDays: 93 };
    });
    window.App.recalc();
  });
  var dl = page.waitForEvent('download', { timeout: 20000 });
  await page.evaluate(function () { window.App.go('reports'); });
  /* Named, not positional: the delivery buttons sit above this one. */
  await page.evaluate(function () {
    document.querySelectorAll('#main button').forEach(function (b) {
      if (b.textContent.trim() === 'خروجی کامل Excel') b.click();
    });
  });
  var d = await dl;
  await d.saveAs(out);
  await browser.close();
  console.log('exported ' + out + ' (' + (fs.statSync(out).size / 1024).toFixed(0) + ' KB)');
}()).catch(function (e) { console.error(e); process.exit(1); });
