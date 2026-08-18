/* ============================================================================
 * app.js — Karaneh Management System
 * ----------------------------------------------------------------------------
 * Wires the calculation engine, the store, the importer and the UI together.
 *
 * The central rule: `state` holds only user input. Every derived number comes
 * out of KaranehEngine.calculate(), which is re-run by `recalc()` after any
 * change. There is no Calculate button because there is nothing to trigger —
 * editing an answer, a job level, the budget or an HOD override recomputes the
 * whole chain and repaints the current view immediately.
 * ==========================================================================*/
(function () {
  'use strict';

  var Engine = window.KaranehEngine;
  var Store  = window.DataStore;
  var Import = window.ExcelImport;
  var U      = window.UI;
  var el = U.el;

  var App = {
    state: null,      // persisted user input
    result: null,     // latest engine output — never persisted
    view: 'dashboard',
    grids: {},
    booted: false
  };
  window.App = App;

  /* ======================================================================
   * Navigation definition
   * ====================================================================*/
  var NAV = [
    { group: 'مرور کلی' },
    { id: 'dashboard',      icon: '▦', label: 'داشبورد' },
    { group: 'ورود اطلاعات' },
    { id: 'employees',      icon: '👤', label: 'اطلاعات پرسنل' },
    { id: 'import',         icon: '📥', label: 'ورود پرسشنامه‌ها' },
    { id: 'questionnaires', icon: '📝', label: 'مدیریت پرسشنامه' },
    { group: 'محاسبات' },
    { id: 'payment',        icon: '💰', label: 'روش پرداخت کارانه' },
    { id: 'hod',            icon: '✍️', label: 'تغییرات معاون بخش' },
    { group: 'کنترل و خروجی' },
    { id: 'validation',     icon: '🛡', label: 'مرکز اعتبارسنجی' },
    { id: 'reports',        icon: '📤', label: 'گزارش و خروجی' },
    { id: 'audit',          icon: '🧾', label: 'ردیابی تغییرات' },
    { id: 'settings',       icon: '⚙️', label: 'تنظیمات' }
  ];

  /* ======================================================================
   * Boot
   * ====================================================================*/
  function boot() {
    Store.read().then(function (saved) {
      App.state = saved || freshState();
      if (!App.state.config) App.state.config = defaultConfig();
      if (!App.state.columnMappings) App.state.columnMappings = cloneMappings();
      App._keySeq = App.state.questionnaires.length;
      renderShell();
      recalc();
      App.booted = true;
    }).catch(function (e) {
      console.error(e);
      App.state = freshState();
      renderShell();
      recalc();
    });
  }

  function freshState() {
    var s = Store.emptyState();
    s.config = defaultConfig();
    s.columnMappings = cloneMappings();
    return s;
  }

  function defaultConfig() {
    var c = {};
    Object.keys(Engine.DEFAULT_CONFIG).forEach(function (k) {
      var v = Engine.DEFAULT_CONFIG[k];
      c[k] = (v && typeof v === 'object') ? JSON.parse(JSON.stringify(v)) : v;
    });
    return c;
  }

  function cloneMappings() {
    return JSON.parse(JSON.stringify(Import.DEFAULT_MAPPINGS));
  }

  function save() {
    return Store.write(App.state).catch(function (e) {
      U.toast(e.message || 'ذخیره‌سازی ناموفق بود.', 'err', 6000);
    });
  }

  /* ======================================================================
   * The single recalculation path
   * ====================================================================*/
  function recalc(options) {
    var records = App.state.questionnaires.map(function (q) {
      var master = employeeById(q.employeeId);
      /* Master data wins for organisational attributes; the questionnaire
         wins for answers. A job level typed into a team file is a fallback
         only, because HR's master file is the system of record. */
      return {
        employeeId:    q.employeeId,
        fullName:      q.fullName || (master && master.fullName) || '',
        division:      (master && master.division) || q.division || '',
        positionTitle: (master && master.positionTitle) || q.positionTitle || '',
        jobLevel:      (master && master.jobLevel) || q.jobLevel || '',
        q1: q.q1, q2: q.q2, q3: q.q3, q4: q.q4, q5: q.q5,
        specialProject:      q.specialProject,
        specialImpactAmount: q.specialImpactAmount,
        hodAdjustment:       q.hodAdjustment,
        hodComment:          q.hodComment,
        excluded:            !!q.excluded,
        sourceFile:          q.sourceFile,
        _key:                q._key
      };
    });

    App.result = Engine.calculate(records, App.state.config);
    App.validation = buildValidation();

    renderNav();
    if (!options || options.repaint !== false) renderView();
    return App.result;
  }

  function employeeById(id) {
    if (!App._empIndex || App._empIndexStamp !== App.state.employees.length) {
      App._empIndex = {};
      App.state.employees.forEach(function (e) { App._empIndex[e.employeeId] = e; });
      App._empIndexStamp = App.state.employees.length;
    }
    return App._empIndex[id];
  }
  function invalidateEmployeeIndex() { App._empIndexStamp = -1; }

  function resultRow(employeeId) {
    if (!App.result) return null;
    var found = null;
    App.result.rows.some(function (r) {
      if (r.employeeId === employeeId) { found = r; return true; }
      return false;
    });
    return found;
  }

  function questionnaireByKey(key) {
    var found = null;
    App.state.questionnaires.some(function (q) {
      if (q._key === key) { found = q; return true; }
      return false;
    });
    return found;
  }

  /* ======================================================================
   * Validation center — everything that must be resolved before finalisation
   * ====================================================================*/
  function buildValidation() {
    var issues = [];
    var res = App.result;
    if (!res) return issues;

    var qById = {};
    App.state.questionnaires.forEach(function (q) {
      (qById[q.employeeId] || (qById[q.employeeId] = [])).push(q);
    });

    function add(severity, code, title, employeeId, employeeName, detail) {
      issues.push({
        severity: severity, code: code, title: title,
        employeeId: employeeId || '', employeeName: employeeName || '', detail: detail || ''
      });
    }

    /* Employees in the master file with no questionnaire at all. */
    App.state.employees.forEach(function (e) {
      if (!isPayrollEligible(e)) return;
      if (!qById[e.employeeId]) {
        add('warn', 'NO_QUESTIONNAIRE', 'پرسنل بدون پرسشنامه',
            e.employeeId, e.fullName, 'واحد: ' + (e.division || '—'));
      }
    });

    /* Questionnaires with no matching master record. */
    App.state.questionnaires.forEach(function (q) {
      if (!employeeById(q.employeeId) && App.state.employees.length) {
        add('warn', 'NO_MASTER', 'پرسشنامه بدون رکورد پرسنلی',
            q.employeeId, q.fullName, 'فایل: ' + (q.sourceFile || '—'));
      }
    });

    /* Duplicates across the consolidated questionnaire set. */
    Store.detectQuestionnaireDuplicates(App.state.questionnaires).forEach(function (d) {
      var unresolved = d.records.filter(function (r) { return !r.excluded; }).length > 1;
      add(unresolved ? 'err' : 'info', 'DUPLICATE', 'پرسشنامه تکراری',
          d.employeeId, d.records[0].fullName,
          'در ' + d.count + ' فایل: ' + d.sources.join('، ') +
          (unresolved ? ' — هنوز تعیین تکلیف نشده است.' : ' — تعیین تکلیف شده.'));
    });

    res.rows.forEach(function (r) {
      if (r.excluded) return;
      if (!r.hasQuestionnaire) {
        add('err', 'INCOMPLETE_ANSWERS', 'پاسخ سؤالات ناقص یا نامعتبر',
            r.employeeId, r.fullName, describeMissingAnswers(r));
      }
      if (r.gradeScore === null && r.jobLevel) {
        add('err', 'UNMAPPED_JL', 'سطح شغلی در جدول گرید تعریف نشده',
            r.employeeId, r.fullName, 'JL = ' + r.jobLevel);
      }
      if (!r.jobLevel) {
        add('err', 'MISSING_JL', 'سطح شغلی نامشخص', r.employeeId, r.fullName, '');
      }
      if (!r.eligible && r.hasQuestionnaire) {
        add('info', 'BELOW_THRESHOLD', 'امتیاز کمتر از حد نصاب',
            r.employeeId, r.fullName,
            'امتیاز ' + U.score(r.performanceScore) + ' ≤ حد نصاب ' +
            App.state.config.minPerformanceThreshold);
      }
      if (r.isOverridden && !r.hodComment) {
        add('err', 'MISSING_HOD_COMMENT', 'تغییر معاون بخش بدون توضیح',
            r.employeeId, r.fullName, 'مبلغ: ' + U.money(r.hodAdjustment));
      }
      if (r.isOverridden && r.hodAdjustment < 0) {
        add('err', 'NEGATIVE_OVERRIDE', 'مبلغ تغییر معاون بخش منفی است',
            r.employeeId, r.fullName, U.money(r.hodAdjustment));
      }
      if (r.negative) {
        add('err', 'NEGATIVE_PAYOUT', 'دریافتی نهایی منفی شده است',
            r.employeeId, r.fullName,
            'تغییرات معاون بخش سهم سایر افراد را منفی کرده است: ' + U.money(r.finalKaraneh));
      }
      if (r.specialProject && !(r.specialImpactValue > 0)) {
        add('warn', 'SPECIAL_NO_AMOUNT', 'اثرگذاری ویژه بدون امتیاز',
            r.employeeId, r.fullName, '');
      }
    });

    if (res.totals.budgetOverrun) {
      add('err', 'BUDGET_OVERRUN', 'عبور از بودجه', '', '',
          'تخصیص ' + U.money(res.totals.allocatedBudget) + ' در برابر بودجه ' +
          U.money(res.totals.budget));
    }
    if (!App.state.employees.length) {
      add('info', 'NO_MASTER_DATA', 'اطلاعات پرسنل وارد نشده است', '', '',
          'تطبیق پرسشنامه با اطلاعات پرسنلی انجام نمی‌شود.');
    }
    if (!App.state.questionnaires.length) {
      add('info', 'NO_QUESTIONNAIRES', 'هیچ پرسشنامه‌ای وارد نشده است', '', '', '');
    }
    return issues;
  }

  function describeMissingAnswers(r) {
    var missing = [], scale = App.state.config.answerScale;
    App.state.config.scoredQuestions.forEach(function (k, i) {
      var v = r[k];
      if (v === null || v === undefined || v === '') missing.push('Q' + (i + 1) + ' خالی');
      else if (typeof v !== 'number' && scale[String(v).trim()] === undefined) {
        missing.push('Q' + (i + 1) + ' نامعتبر («' + v + '»)');
      }
    });
    return missing.join('، ');
  }

  /** Who HR expects a questionnaire for. Leavers and non-active staff are out. */
  function isPayrollEligible(e) {
    var st = String(e.employeeStatus || '').trim().toLowerCase();
    if (st === 'non active' || st === 'inactive' || st === 'غیرفعال') return false;
    if (st === 'to be non active') return false;
    return true;
  }

  function issueCount(severity) {
    return (App.validation || []).filter(function (i) { return i.severity === severity; }).length;
  }

  /* ======================================================================
   * Shell
   * ====================================================================*/
  function renderShell() {
    var root = document.getElementById('app');
    U.clear(root);

    root.appendChild(el('div', { class: 'topbar' }, [
      el('span', { class: 'brand', text: 'سامانه مدیریت کارانه' }),
      el('span', { class: 'period', id: 'periodChip', text: App.state.period }),
      el('div', { class: 'spacer' }),
      el('div', { class: 'topstat', id: 'topBudget' }),
      el('div', { class: 'sep' }),
      el('div', { class: 'topstat', id: 'topAllocated' }),
      el('div', { class: 'sep' }),
      el('div', { class: 'topstat', id: 'topRemaining' }),
      el('div', { class: 'sep' }),
      el('div', { class: 'topstat', id: 'topStatus' })
    ]));

    var sidebar = el('nav', { class: 'sidebar', id: 'sidebar' });
    var main = el('main', { class: 'main', id: 'main' });
    root.appendChild(el('div', { class: 'shell' }, [main, sidebar]));
    renderNav();
  }

  function renderNav() {
    var sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    U.clear(sidebar);
    NAV.forEach(function (n) {
      if (n.group) { sidebar.appendChild(el('div', { class: 'navgroup', text: n.group })); return; }
      var badge = null;
      if (n.id === 'validation') {
        var errs = issueCount('err'), warns = issueCount('warn');
        if (errs) badge = el('span', { class: 'badge err', text: String(errs) });
        else if (warns) badge = el('span', { class: 'badge warn', text: String(warns) });
      } else if (n.id === 'questionnaires') {
        badge = el('span', { class: 'badge', text: String(App.state.questionnaires.length) });
      } else if (n.id === 'employees') {
        badge = el('span', { class: 'badge', text: String(App.state.employees.length) });
      } else if (n.id === 'hod') {
        var ov = App.result ? App.result.totals.overriddenCount : 0;
        if (ov) badge = el('span', { class: 'badge', text: String(ov) });
      }
      sidebar.appendChild(el('button', {
        class: 'navitem' + (App.view === n.id ? ' active' : ''),
        onclick: function () { go(n.id); }
      }, [
        el('span', { class: 'ico', text: n.icon }),
        el('span', { text: n.label }),
        badge
      ]));
    });
    renderTopStats();
  }

  function renderTopStats() {
    var t = App.result ? App.result.totals : null;
    function set(id, value, label, color) {
      var node = document.getElementById(id);
      if (!node) return;
      U.clear(node);
      node.appendChild(el('b', {}, [U.bidi(value)]));
      node.appendChild(el('span', { text: label }));
      node.style.color = color || '';
    }
    if (!t) return;
    set('topBudget', U.moneyShort(t.budget), 'بودجه');
    set('topAllocated', U.moneyShort(t.allocatedBudget), 'تخصیص‌یافته');
    set('topRemaining', U.moneyShort(t.remainingBudget), 'باقیمانده',
        t.remainingBudget < -1 ? '#ffd0cc' : '');
    var label = t.budgetStatus === 'BALANCED' ? 'متوازن'
              : t.budgetStatus === 'OVERRUN'  ? 'عبور از بودجه' : 'نامعتبر';
    set('topStatus', label, 'وضعیت', t.budgetStatus === 'BALANCED' ? '#c8f2dd' : '#ffd0cc');
  }

  /* Exposed alongside `go` so automated tests and the browser console can
     drive the same code paths the UI does, rather than a parallel one. */
  App.recalc = recalc;
  App.save = save;

  function go(view) {
    App.view = view;
    renderNav();
    renderView();
    var m = document.getElementById('main');
    if (m) m.scrollTop = 0;
  }
  App.go = go;

  function renderView() {
    var main = document.getElementById('main');
    if (!main) return;
    U.clear(main);
    var fn = VIEWS[App.view] || VIEWS.dashboard;
    fn(main);
  }

  function head(title, subtitle, actions) {
    return el('div', { class: 'view-head' }, [
      el('div', {}, [el('h1', { text: title }), el('p', { text: subtitle || '' })]),
      actions ? el('div', { class: 'actions' }, actions) : null
    ]);
  }

  function btn(label, onClick, kind) {
    return el('button', { class: 'btn ' + (kind || ''), text: label, onclick: onClick });
  }

  var VIEWS = {};

  /* ======================================================================
   * VIEW — Dashboard
   * ====================================================================*/
  VIEWS.dashboard = function (main) {
    var t = App.result.totals;
    var withQ = App.state.questionnaires.filter(function (q) { return !q.excluded; }).length;
    var missing = App.state.employees.filter(function (e) {
      return isPayrollEligible(e) && !App.state.questionnaires.some(function (q) {
        return q.employeeId === e.employeeId && !q.excluded;
      });
    }).length;

    main.appendChild(head(
      'داشبورد — ' + App.state.period,
      'وضعیت لحظه‌ای فرآیند کارانه. تمام اعداد با هر تغییر بلافاصله بازمحاسبه می‌شوند.',
      [
        btn('ورود پرسشنامه', function () { go('import'); }, 'primary'),
        btn('خروجی Excel', function () { exportWorkbook(); })
      ]));

    if (t.budgetStatus !== 'BALANCED') {
      main.appendChild(U.alert('err', 'بودجه در وضعیت نامعتبر است',
        t.budgetOverrun
          ? 'مجموع کارانه نهایی از بودجه تعیین‌شده بیشتر است. تا رفع این مورد، نهایی‌سازی ممکن نیست.'
          : t.negativePayoutCount + ' نفر با اعمال تغییرات معاون بخش دریافتی منفی پیدا کرده‌اند. تا رفع این مورد، نهایی‌سازی ممکن نیست.',
        btn('بررسی', function () { go('validation'); }, 'sm')));
    }
    if (issueCount('err')) {
      main.appendChild(U.alert('warn', issueCount('err') + ' خطای باز در مرکز اعتبارسنجی',
        'این موارد باید پیش از نهایی‌سازی برطرف شوند.',
        btn('مشاهده', function () { go('validation'); }, 'sm')));
    }

    var grid = el('div', { class: 'kpi-grid' });
    grid.appendChild(U.kpi('کل پرسنل', U.int(App.state.employees.length), { kind: 'brand' }));
    grid.appendChild(U.kpi('واجد شرایط', U.int(t.eligibleCount),
      { kind: 'ok', sub: 'امتیاز بالاتر از حد نصاب' }));
    grid.appendChild(U.kpi('دارای پرسشنامه', U.int(withQ), { kind: 'info' }));
    grid.appendChild(U.kpi('فاقد پرسشنامه', U.int(missing),
      { kind: missing ? 'warn' : '', sub: 'از میان پرسنل فعال' }));
    grid.appendChild(U.kpi('زیر حد نصاب', U.int(t.ineligibleCount),
      { kind: t.ineligibleCount ? 'warn' : '', sub: 'کارانه صفر' }));
    grid.appendChild(U.kpi('تغییرات معاون بخش', U.int(t.overriddenCount), { kind: 'info' }));
    grid.appendChild(U.kpi('موارد استثنا', U.int(issueCount('err')),
      { kind: issueCount('err') ? 'err' : 'ok' }));
    grid.appendChild(U.kpi('کارانه نهایی', U.moneyShort(t.sumFinalKaraneh),
      { kind: 'brand', sub: 'ریال' }));
    main.appendChild(grid);

    var used = t.budget ? Math.min(1, t.allocatedBudget / t.budget) : 0;
    var over = t.budget && t.allocatedBudget > t.budget
      ? Math.min(1, (t.allocatedBudget - t.budget) / t.budget) : 0;
    var bar = el('div', {}, [
      el('div', { class: 'budget-bar' }, [
        el('i', { class: 'used', style: 'width:' + (used * 100).toFixed(3) + '%' }),
        over ? el('i', { class: 'over', style: 'width:' + (over * 100).toFixed(3) + '%' }) : null
      ]),
      el('div', { class: 'budget-legend' }, [
        el('span', { html: 'بودجه: <b class="num">' + U.money(t.budget) + '</b> ریال' }),
        el('span', { html: 'تخصیص‌یافته: <b class="num">' + U.money(t.allocatedBudget) + '</b>' }),
        el('span', { html: 'باقیمانده: <b class="num">' + U.money(t.remainingBudget) + '</b>' }),
        el('span', {}, [
          document.createTextNode('وضعیت: '),
          el('span', {
            class: 'chip ' + (t.budgetStatus === 'BALANCED' ? 'ok' : 'err'),
            text: t.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر'
          })
        ])
      ])
    ]);
    main.appendChild(U.card('کنترل بودجه', bar, {
      hint: 'مجموع تخصیص همواره دقیقاً برابر بودجه است؛ تغییرات معاون بخش بین سایر افراد سرشکن می‌شود.'
    }));

    var byDiv = {};
    App.result.rows.forEach(function (r) {
      if (!r.inScope) return;
      var d = r.division || '—';
      var g = byDiv[d] || (byDiv[d] = { division: d, count: 0, eligible: 0, score: 0, amount: 0, overrides: 0 });
      g.count++;
      if (r.eligible) g.eligible++;
      g.score += r.totalScore;
      g.amount += r.finalKaraneh;
      if (r.isOverridden) g.overrides++;
    });
    var divRows = Object.keys(byDiv).map(function (k) { return byDiv[k]; });

    if (divRows.length) {
      var dg = U.DataGrid({
        title: 'خلاصه به تفکیک واحد سازمانی',
        rows: divRows,
        sortKey: 'amount', sortDir: 'desc',
        searchFields: ['division'],
        columns: [
          { key: 'division', label: 'واحد سازمانی', alwaysVisible: true },
          { key: 'count', label: 'تعداد', type: 'int' },
          { key: 'eligible', label: 'واجد شرایط', type: 'int' },
          { key: 'score', label: 'امتیاز کل', type: 'score', calculated: true },
          { key: 'overrides', label: 'تغییر معاون', type: 'int' },
          { key: 'amount', label: 'کارانه نهایی (ریال)', type: 'money', calculated: true },
          { key: 'share', label: 'سهم از بودجه', type: 'percent', calculated: true,
            value: function (r) { return t.budget ? r.amount / t.budget : 0; } }
        ],
        footer: function (rows) {
          var s = { count: 0, eligible: 0, score: 0, amount: 0, overrides: 0 };
          rows.forEach(function (r) {
            s.count += r.count; s.eligible += r.eligible;
            s.score += r.score; s.amount += r.amount; s.overrides += r.overrides;
          });
          return {
            division: 'جمع', count: U.int(s.count), eligible: U.int(s.eligible),
            score: U.score(s.score), overrides: U.int(s.overrides),
            amount: U.money(s.amount), share: U.percent(t.budget ? s.amount / t.budget : 0)
          };
        }
      });
      main.appendChild(dg.node);
    }

    main.appendChild(U.card('مسیر فرآیند', workflowNode(), { hint: 'وضعیت هر مرحله' }));
  };

  function workflowNode() {
    var t = App.result.totals;
    var steps = [
      { label: 'اطلاعات پرسنل', done: App.state.employees.length > 0,
        detail: U.int(App.state.employees.length) + ' نفر' },
      { label: 'ورود پرسشنامه', done: App.state.questionnaires.length > 0,
        detail: U.int(App.state.questionnaires.length) + ' رکورد' },
      { label: 'اعتبارسنجی', done: issueCount('err') === 0,
        detail: issueCount('err') + ' خطا' },
      { label: 'محاسبه', done: t.eligibleCount > 0, detail: U.int(t.eligibleCount) + ' نفر' },
      { label: 'تغییرات معاون بخش', done: true, detail: U.int(t.overriddenCount) + ' مورد' },
      { label: 'کنترل بودجه', done: t.budgetStatus === 'BALANCED',
        detail: t.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر' },
      { label: 'نهایی‌سازی', done: !!App.state.finalizedAt,
        detail: App.state.finalizedAt ? U.dateTime(App.state.finalizedAt) : 'انجام نشده' }
    ];
    var wrap = el('div', { style: 'display:flex;flex-wrap:wrap;gap:8px' });
    steps.forEach(function (s, i) {
      wrap.appendChild(el('div', {
        style: 'flex:1;min-width:130px;border:1px solid var(--border);border-radius:8px;' +
               'padding:9px 11px;background:' + (s.done ? 'var(--ok-soft)' : 'var(--surface-2)')
      }, [
        el('div', { class: 'small muted', text: (i + 1) + '. ' + s.label }),
        el('div', { style: 'font-weight:700;font-size:12.5px', text: s.detail }),
        el('span', { class: 'chip ' + (s.done ? 'ok' : 'warn'), text: s.done ? 'انجام شد' : 'در انتظار' })
      ]));
    });
    return wrap;
  }

  /* ======================================================================
   * VIEW — Employee master
   * ====================================================================*/
  VIEWS.employees = function (main) {
    main.appendChild(head('اطلاعات پرسنل',
      'شماره پرسنلی کلید یکتای سیستم است. رکورد تکراری بدون تأیید شما وارد نمی‌شود.',
      [
        btn('ورود فایل پرسنل', function () { pickFiles('employee'); }, 'primary'),
        btn('خروجی', function () { exportSheet('employees'); }),
        App.state.employees.length ? btn('پاک کردن', function () { clearEmployees(); }, 'danger') : null
      ].filter(Boolean)));

    if (!App.state.employees.length) {
      main.appendChild(dropzoneNode('employee',
        'فایل اطلاعات پرسنل را اینجا رها کنید',
        'ستون‌های Emp No، نام، سطح شغلی، واحد سازمانی و … به‌صورت خودکار شناسایی می‌شوند.'));
      return;
    }

    var grid = U.DataGrid({
      title: 'اطلاعات پایه پرسنل',
      rows: App.state.employees,
      sortKey: 'employeeId',
      searchFields: ['employeeId', 'fullName', 'firstName', 'lastName', 'division', 'positionTitle', 'directManager'],
      facets: [
        { key: 'division', label: 'همه واحدها' },
        { key: 'jobLevel', label: 'همه سطوح شغلی' },
        { key: 'employeeStatus', label: 'همه وضعیت‌ها' },
        { key: 'employmentType', label: 'همه انواع استخدام' }
      ],
      columns: [
        { key: 'employeeId', label: 'شماره پرسنلی', alwaysVisible: true, width: '100px' },
        { key: 'fullName', label: 'نام و نام خانوادگی', width: '160px',
          value: function (r) { return r.fullName || ((r.firstName || '') + ' ' + (r.lastName || '')).trim(); } },
        { key: 'employeeStatus', label: 'وضعیت' },
        { key: 'positionTitle', label: 'عنوان شغلی', width: '180px' },
        { key: 'jobLevel', label: 'JL', width: '50px' },
        { key: 'division', label: 'واحد سازمانی' },
        { key: 'department', label: 'دپارتمان', hidden: true },
        { key: 'employmentType', label: 'نوع استخدام' },
        { key: 'assignmentType', label: 'نوع همکاری', hidden: true },
        { key: 'workingDays', label: 'روز کارکرد', type: 'int' },
        { key: 'dateOfEmployment', label: 'تاریخ استخدام', hidden: true },
        { key: 'dateOfLeaving', label: 'تاریخ خروج', hidden: true },
        { key: 'directManager', label: 'مدیر مستقیم' },
        { key: 'managerLevel1', label: 'مدیر سطح ۱', hidden: true },
        { key: 'managerLevel2', label: 'مدیر سطح ۲', hidden: true },
        { key: 'managerLevel3', label: 'مدیر سطح ۳', hidden: true },
        { key: 'nationalId', label: 'کد ملی', hidden: true },
        { key: 'gender', label: 'جنسیت', hidden: true },
        { key: 'hasQ', label: 'پرسشنامه', calculated: true,
          render: function (r) {
            var has = App.state.questionnaires.some(function (q) {
              return q.employeeId === r.employeeId && !q.excluded;
            });
            return el('span', { class: 'chip ' + (has ? 'ok' : 'warn'), text: has ? 'دارد' : 'ندارد' });
          } },
        { key: 'sourceFile', label: 'فایل منبع', hidden: true }
      ],
      onRowClick: function (r) { showEmployeeDetail(r.employeeId); }
    });
    App.grids.employees = grid;
    main.appendChild(grid.node);
  };

  function clearEmployees() {
    U.confirm('تمام اطلاعات پرسنل حذف شود؟ پرسشنامه‌ها دست‌نخورده می‌مانند.',
      { danger: true, confirmLabel: 'حذف' }).then(function (ok) {
      if (!ok) return;
      Store.audit(App.state, {
        entity: 'employees', field: 'all',
        oldValue: App.state.employees.length + ' رکورد', newValue: '0',
        reason: 'پاک کردن دستی اطلاعات پرسنل'
      });
      App.state.employees = [];
      invalidateEmployeeIndex();
      save().then(function () { recalc(); U.toast('اطلاعات پرسنل پاک شد.', 'ok'); });
    });
  }

  /* ======================================================================
   * VIEW — Import
   * ====================================================================*/
  VIEWS.import = function (main) {
    main.appendChild(head('ورود پرسشنامه‌های تیمی',
      'چند فایل را همزمان انتخاب کنید. سیستم شیت «پرسشنامه کارانه تیمی» را پیدا می‌کند، ' +
      'ردیف عنوان را تشخیص می‌دهد، ستون‌ها را نگاشت می‌کند و همه را در یک مجموعه واحد ادغام می‌کند.',
      [btn('انتخاب فایل‌ها', function () { pickFiles('questionnaire'); }, 'primary')]));

    main.appendChild(dropzoneNode('questionnaire',
      'فایل‌های پرسشنامه را اینجا رها کنید',
      'xlsx / xls / xlsb / csv — انتخاب چندتایی پشتیبانی می‌شود.'));

    var batches = App.state.importBatches || [];
    if (batches.length) {
      var grid = U.DataGrid({
        title: 'فایل‌های وارد شده',
        rows: batches.slice().reverse(),
        searchFields: ['fileName', 'sheetName'],
        pageSize: 50,
        columns: [
          { key: 'fileName', label: 'نام فایل', width: '200px' },
          { key: 'kind', label: 'نوع', render: function (r) {
            return el('span', { class: 'chip ' + (r.kind === 'employee' ? 'info' : 'brand'),
              text: r.kind === 'employee' ? 'پرسنل' : 'پرسشنامه' });
          } },
          { key: 'sheetName', label: 'شیت' },
          { key: 'headerRow', label: 'ردیف عنوان', type: 'int' },
          { key: 'accepted', label: 'رکورد پذیرفته', type: 'int' },
          { key: 'mappedCount', label: 'ستون نگاشت‌شده', type: 'int' },
          { key: 'importedAt', label: 'زمان', render: function (r) { return U.dateTime(r.importedAt); } },
          { key: 'act', label: '', render: function (r) {
            return el('span', {}, [
              el('button', { class: 'btn sm', text: 'نگاشت', onclick: function () { showBatchMapping(r); } }),
              el('button', { class: 'btn sm danger', text: 'حذف',
                style: 'margin-inline-start:5px', onclick: function () { removeBatch(r); } })
            ]);
          } }
        ]
      });
      main.appendChild(grid.node);
    }

    var dups = Store.detectQuestionnaireDuplicates(App.state.questionnaires);
    if (dups.length) {
      var body = el('div', {});
      body.appendChild(el('p', { class: 'small muted',
        text: 'این افراد در بیش از یک فایل پرسشنامه دیده شده‌اند. تا زمانی که تعیین تکلیف نشوند، ' +
              'هیچ‌کدام از رکوردهای تکراری وارد محاسبه نمی‌شوند.' }));
      dups.forEach(function (d) {
        var unresolved = d.records.filter(function (r) { return !r.excluded; }).length > 1;
        body.appendChild(el('div', {
          style: 'border:1px solid var(--border);border-radius:8px;padding:10px 12px;margin-bottom:9px;' +
                 'background:' + (unresolved ? 'var(--err-soft)' : 'var(--ok-soft)')
        }, [
          el('div', {}, [
            el('b', { text: 'شماره پرسنلی ' + d.employeeId }),
            el('span', { class: 'muted small', text: '  ' + (d.records[0].fullName || '') }),
            el('span', { class: 'chip ' + (unresolved ? 'err' : 'ok'),
              style: 'margin-inline-start:8px',
              text: unresolved ? 'تعیین تکلیف نشده' : 'تعیین تکلیف شد' })
          ]),
          el('div', { class: 'small muted', text: 'یافت شده در: ' + d.sources.join('، ') }),
          el('div', { style: 'margin-top:7px;display:flex;gap:6px;flex-wrap:wrap' }, [
            btn('اولی بماند', function () { resolveDup(d.employeeId, 'keepFirst'); }, 'sm'),
            btn('آخری بماند', function () { resolveDup(d.employeeId, 'keepLatest'); }, 'sm'),
            btn('ادغام', function () { resolveDup(d.employeeId, 'merge'); }, 'sm'),
            btn('انتخاب دستی', function () { chooseDuplicate(d); }, 'sm primary')
          ])
        ]));
      });
      main.appendChild(U.card('کنترل رکوردهای تکراری (' + dups.length + ')', body));
    }
  };

  function dropzoneNode(kind, title, sub) {
    var dz = el('div', { class: 'dropzone' }, [
      el('div', { class: 'big', text: '📁' }),
      el('div', { class: 't', text: title }),
      el('div', { class: 's', text: sub })
    ]);
    dz.addEventListener('click', function () { pickFiles(kind); });
    dz.addEventListener('dragover', function (e) { e.preventDefault(); dz.classList.add('over'); });
    dz.addEventListener('dragleave', function () { dz.classList.remove('over'); });
    dz.addEventListener('drop', function (e) {
      e.preventDefault(); dz.classList.remove('over');
      handleFiles(Array.prototype.slice.call(e.dataTransfer.files), kind);
    });
    return dz;
  }

  function pickFiles(kind) {
    var input = el('input', {
      type: 'file', accept: '.xlsx,.xls,.xlsb,.csv',
      multiple: kind === 'questionnaire' ? 'multiple' : null,
      style: 'display:none'
    });
    input.addEventListener('change', function () {
      handleFiles(Array.prototype.slice.call(input.files), kind);
      input.value = '';
    });
    document.body.appendChild(input);
    input.click();
    setTimeout(function () { if (input.parentNode) input.parentNode.removeChild(input); }, 1000);
  }

  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(new Error('خواندن فایل «' + file.name + '» ناموفق بود.')); };
      fr.readAsArrayBuffer(file);
    });
  }

  function handleFiles(files, kind) {
    if (!files.length) return;
    var m = U.modal({
      title: 'در حال پردازش ' + files.length + ' فایل…',
      size: 'narrow', dismissable: false,
      content: el('div', {}, [el('span', { class: 'spinner' }), document.createTextNode(' لطفاً صبر کنید…')])
    });

    var parsed = [], errors = [];
    var chain = Promise.resolve();
    files.forEach(function (f) {
      chain = chain.then(function () {
        return readFile(f).then(function (buf) {
          parsed.push(Import.parseWorkbook(buf, {
            fileName: f.name, kind: kind, mappings: App.state.columnMappings
          }));
        }).catch(function (e) {
          errors.push({ file: f.name, message: e.message || String(e) });
        });
      });
    });

    chain.then(function () {
      m.close();
      if (!parsed.length) {
        U.modal({
          title: 'هیچ فایلی خوانده نشد', size: 'narrow',
          content: el('div', {}, errors.map(function (e) {
            return U.alert('err', e.file, U.esc(e.message));
          })),
          buttons: [{ label: 'بستن', kind: 'primary' }]
        });
        return;
      }
      showImportPreview(parsed, errors, kind);
    });
  }

  /* ---------------------------------------------------------------------
   * Import preview — nothing is committed until the user has seen the
   * detected mapping and any duplicate conflicts.
   * -------------------------------------------------------------------*/
  function showImportPreview(parsed, errors, kind) {
    var body = el('div', {});
    errors.forEach(function (e) { body.appendChild(U.alert('err', e.file, U.esc(e.message))); });

    var totalRecords = 0;
    parsed.forEach(function (p) { totalRecords += p.records.length; });

    var staged = null, dupsInBatch = [];
    if (kind === 'employee') {
      var incoming = [];
      parsed.forEach(function (p) { incoming = incoming.concat(p.records); });
      staged = Store.stageEmployees(App.state.employees, incoming);
      dupsInBatch = staged.duplicatesWithinBatch;
    }

    parsed.forEach(function (p) {
      var mapped = Object.keys(p.mapping);
      var det = el('div', { style: 'margin-bottom:14px' }, [
        el('div', { style: 'display:flex;gap:9px;align-items:baseline;flex-wrap:wrap' }, [
          el('b', { text: p.fileName }),
          el('span', { class: 'chip brand', text: 'شیت: ' + p.sheetName }),
          el('span', { class: 'chip', text: 'ردیف عنوان: ' + p.headerRow }),
          el('span', { class: 'chip ok', text: p.records.length + ' رکورد' }),
          el('span', { class: 'chip info', text: mapped.length + ' ستون نگاشت شد' })
        ]),
        p.warnings.length ? el('div', { class: 'small', style: 'color:var(--warn);margin-top:4px',
          text: '⚠ ' + p.warnings.join(' • ') }) : null
      ]);

      var tbl = el('table', { class: 'grid', style: 'margin-top:7px' });
      tbl.appendChild(el('thead', {}, [el('tr', {}, [
        el('th', { text: 'فیلد سیستم' }), el('th', { text: 'عنوان ستون در فایل' }),
        el('th', { text: 'ستون' }), el('th', { text: 'نمونه مقدار' })
      ])]));
      var tb = el('tbody');
      mapped.sort().forEach(function (field) {
        var idx = p.mapping[field];
        var sampleRec = p.records[0] || {};
        tb.appendChild(el('tr', {}, [
          el('td', {}, [
            el('b', { text: (App.state.columnMappings[field] || {}).label || field }),
            el('div', { class: 'small muted mono', text: field })
          ]),
          el('td', { text: String(p.headers[idx] === null || p.headers[idx] === undefined ? '' : p.headers[idx]).slice(0, 60) }),
          el('td', { class: 'num mono', text: colLetter(idx) }),
          el('td', { class: 'muted small', text: String(sampleRec[field] === null || sampleRec[field] === undefined ? '—' : sampleRec[field]).slice(0, 40) })
        ]));
      });
      tbl.appendChild(tb);
      det.appendChild(el('div', { class: 'table-wrap', style: 'max-height:230px' }, [tbl]));

      if (p.unmapped.length) {
        det.appendChild(el('div', { class: 'small muted', style: 'margin-top:5px',
          text: 'ستون‌های نادیده‌گرفته‌شده: ' + p.unmapped.map(function (u) {
            return String(u.header).slice(0, 30);
          }).join(' • ') }));
      }
      body.appendChild(det);
    });

    if (staged) {
      if (staged.conflicts.length) {
        body.appendChild(U.alert('warn', staged.conflicts.length + ' شماره پرسنلی از قبل موجود است',
          'این رکوردها با اطلاعات فعلی تفاوت دارند. تعیین کنید کدام نسخه بماند.'));
        var ct = el('table', { class: 'grid' });
        ct.appendChild(el('thead', {}, [el('tr', {}, [
          el('th', { text: 'شماره پرسنلی' }), el('th', { text: 'نام' }),
          el('th', { text: 'فیلدهای متفاوت' }), el('th', { text: 'جدیدتر' })
        ])]));
        var cb2 = el('tbody');
        staged.conflicts.slice(0, 200).forEach(function (c) {
          cb2.appendChild(el('tr', {}, [
            el('td', { text: c.employeeId }),
            el('td', { text: c.incoming.fullName || c.existing.fullName || '' }),
            el('td', { class: 'small', text: c.changedFields.map(function (f) {
              return f.field + ': «' + (f.from || '—') + '» ← «' + (f.to || '—') + '»';
            }).join(' • ').slice(0, 160) }),
            el('td', {}, [el('span', {
              class: 'chip ' + (c.newer === 'incoming' ? 'ok' : 'info'),
              text: c.newer === 'incoming' ? 'فایل جدید' : c.newer === 'existing' ? 'رکورد فعلی' : 'نامشخص'
            })])
          ]));
        });
        ct.appendChild(cb2);
        body.appendChild(el('div', { class: 'table-wrap', style: 'max-height:230px' }, [ct]));
      }
      if (dupsInBatch.length) {
        body.appendChild(U.alert('err', dupsInBatch.length + ' شماره پرسنلی تکراری داخل خود فایل',
          'تنها اولین رکورد هر شماره پرسنلی وارد می‌شود: ' +
          dupsInBatch.slice(0, 20).map(function (d) { return d.employeeId; }).join('، ')));
      }
      body.appendChild(el('div', { class: 'small', style: 'margin-top:8px' }, [
        el('b', { text: staged.fresh.length + ' رکورد جدید' }),
        document.createTextNode(' • ' + staged.conflicts.length + ' رکورد موجود با تغییرات')
      ]));
    }

    var buttons = [];
    if (kind === 'employee') {
      buttons.push({
        label: 'افزودن فقط رکوردهای جدید (' + staged.fresh.length + ')', kind: 'primary',
        onClick: function () { commitEmployees(parsed, staged, 'freshOnly'); }
      });
      if (staged.conflicts.length) {
        buttons.push({
          label: 'به‌روزرسانی موجودها + افزودن جدیدها',
          onClick: function () { commitEmployees(parsed, staged, 'overwrite'); }
        });
      }
    } else {
      buttons.push({
        label: 'ادغام ' + totalRecords + ' رکورد در مجموعه واحد', kind: 'primary',
        onClick: function () { commitQuestionnaires(parsed); }
      });
    }
    buttons.push('spacer', { label: 'انصراف' });

    U.modal({
      title: 'پیش‌نمایش ورود اطلاعات — ' + parsed.length + ' فایل',
      size: 'wide', content: body, buttons: buttons
    });
  }

  function colLetter(i) {
    var s = ''; i += 1;
    while (i > 0) { var r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }

  function commitEmployees(parsed, staged, mode) {
    var added = 0, updated = 0;
    staged.fresh.forEach(function (r) { App.state.employees.push(r); added++; });
    if (mode === 'overwrite') {
      staged.conflicts.forEach(function (c) {
        var idx = App.state.employees.indexOf(c.existing);
        if (idx === -1) return;
        c.changedFields.forEach(function (f) {
          Store.audit(App.state, {
            entity: 'employee', employeeId: c.employeeId,
            employeeName: c.incoming.fullName || c.existing.fullName,
            field: f.field, oldValue: f.from, newValue: f.to,
            reason: 'به‌روزرسانی از فایل ' + (c.incoming.sourceFile || '')
          });
        });
        App.state.employees[idx] = c.incoming;
        updated++;
      });
    }
    parsed.forEach(function (p) { recordBatch(p, 'employee', p.records.length); });
    invalidateEmployeeIndex();
    save().then(function () {
      recalc();
      U.toast(added + ' رکورد جدید اضافه و ' + updated + ' رکورد به‌روزرسانی شد.', 'ok');
      go('employees');
    });
  }

  function commitQuestionnaires(parsed) {
    var added = 0;
    parsed.forEach(function (p) {
      p.records.forEach(function (r) {
        r._key = 'q' + (App._keySeq = (App._keySeq || 0) + 1);
        App.state.questionnaires.push(r);
        added++;
      });
      recordBatch(p, 'questionnaire', p.records.length);
    });
    Store.audit(App.state, {
      entity: 'questionnaire', field: 'import',
      oldValue: '', newValue: added + ' رکورد',
      reason: 'ورود ' + parsed.length + ' فایل پرسشنامه'
    });

    /* Anything appearing twice is parked until the user decides. */
    var dups = Store.detectQuestionnaireDuplicates(App.state.questionnaires);
    dups.forEach(function (d) {
      d.records.forEach(function (r) { if (r.duplicateResolution === undefined) r.excluded = true; });
    });

    save().then(function () {
      recalc();
      if (dups.length) {
        U.toast(added + ' رکورد وارد شد — ' + dups.length +
          ' مورد تکراری تا تعیین تکلیف از محاسبات کنار گذاشته شد.', 'warn', 7000);
        go('import');
      } else {
        U.toast(added + ' رکورد پرسشنامه ادغام شد.', 'ok');
        go('questionnaires');
      }
    });
  }

  function recordBatch(p, kind, accepted) {
    App.state.importBatches.push({
      batchId: p.batchId, fileName: p.fileName, kind: kind,
      sheetName: p.sheetName, sheetNames: p.sheetNames,
      headerRow: p.headerRow, headers: p.headers, mapping: p.mapping,
      mappedCount: Object.keys(p.mapping).length,
      unmapped: p.unmapped, warnings: p.warnings,
      accepted: accepted, importedAt: p.importedAt
    });
  }

  function removeBatch(batch) {
    U.confirm('تمام رکوردهای وارد شده از فایل «' + batch.fileName + '» حذف شوند؟',
      { danger: true, confirmLabel: 'حذف' }).then(function (ok) {
      if (!ok) return;
      var before = App.state.questionnaires.length;
      App.state.questionnaires = App.state.questionnaires.filter(function (q) {
        return q.importBatchId !== batch.batchId;
      });
      if (batch.kind === 'employee') {
        App.state.employees = App.state.employees.filter(function (e) {
          return e.importBatchId !== batch.batchId;
        });
        invalidateEmployeeIndex();
      }
      App.state.importBatches = App.state.importBatches.filter(function (b) {
        return b.batchId !== batch.batchId;
      });
      Store.audit(App.state, {
        entity: 'import', field: 'remove batch',
        oldValue: batch.fileName, newValue: '', reason: 'حذف دستی فایل وارد شده'
      });
      save().then(function () {
        recalc();
        U.toast((before - App.state.questionnaires.length) + ' رکورد حذف شد.', 'ok');
      });
    });
  }

  function showBatchMapping(batch) {
    var tbl = el('table', { class: 'grid' });
    tbl.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', { text: 'ستون' }), el('th', { text: 'عنوان در فایل' }), el('th', { text: 'فیلد سیستم' })
    ])]));
    var tb = el('tbody');
    (batch.headers || []).forEach(function (h, i) {
      var field = Object.keys(batch.mapping).filter(function (f) { return batch.mapping[f] === i; })[0];
      tb.appendChild(el('tr', {}, [
        el('td', { class: 'mono', text: colLetter(i) }),
        el('td', { text: String(h === null || h === undefined ? '' : h).slice(0, 70) }),
        el('td', {}, [field
          ? el('span', { class: 'chip ok', text: (App.state.columnMappings[field] || {}).label || field })
          : el('span', { class: 'chip', text: 'نادیده گرفته شد' })])
      ]));
    });
    tbl.appendChild(tb);
    U.modal({
      title: 'نگاشت ستون‌ها — ' + batch.fileName, size: 'wide',
      content: el('div', {}, [
        el('p', { class: 'small muted', text: 'شیت «' + batch.sheetName + '» • ردیف عنوان ' + batch.headerRow }),
        el('div', { class: 'table-wrap' }, [tbl])
      ]),
      buttons: [{ label: 'بستن', kind: 'primary' }]
    });
  }

  function resolveDup(employeeId, strategy) {
    Store.resolveDuplicate(App.state.questionnaires, employeeId, strategy);
    Store.audit(App.state, {
      entity: 'questionnaire', employeeId: employeeId, field: 'duplicate',
      oldValue: 'unresolved', newValue: strategy, reason: 'تعیین تکلیف رکورد تکراری'
    });
    save().then(function () { recalc(); U.toast('رکورد تکراری تعیین تکلیف شد.', 'ok'); });
  }

  function chooseDuplicate(d) {
    var chosen = 0;
    var body = el('div', {});
    d.records.forEach(function (r, i) {
      var radio = el('input', { type: 'radio', name: 'dup', value: String(i) });
      if (i === 0) radio.checked = true;
      radio.addEventListener('change', function () { chosen = i; });
      body.appendChild(el('label', {
        class: 'checkline',
        style: 'border:1px solid var(--border);border-radius:6px;padding:8px 10px;margin-bottom:7px'
      }, [
        radio,
        el('div', {}, [
          el('b', { text: r.sourceFile || '(نامشخص)' }),
          el('div', { class: 'small muted', text:
            'Q1..Q4: ' + [r.q1, r.q2, r.q3, r.q4].join(' / ') +
            ' • ردیف ' + r.sourceRow + ' • ' + U.dateTime(r.importedAt) })
        ])
      ]));
    });
    U.modal({
      title: 'انتخاب رکورد معتبر — شماره پرسنلی ' + d.employeeId,
      size: 'narrow', content: body,
      buttons: [
        { label: 'اعمال', kind: 'primary', onClick: function () {
          Store.resolveDuplicate(App.state.questionnaires, d.employeeId, 'keepSelected', chosen);
          Store.audit(App.state, {
            entity: 'questionnaire', employeeId: d.employeeId, field: 'duplicate',
            oldValue: 'unresolved', newValue: 'keepSelected:' + d.records[chosen].sourceFile,
            reason: 'انتخاب دستی رکورد معتبر'
          });
          save().then(function () { recalc(); U.toast('اعمال شد.', 'ok'); });
        } },
        { label: 'انصراف' }
      ]
    });
  }

  /* ======================================================================
   * VIEW — Questionnaire management (editable, real-time)
   * ====================================================================*/
  VIEWS.questionnaires = function (main) {
    main.appendChild(head('مدیریت پرسشنامه کارانه تیمی',
      'پاسخ‌ها قابل ویرایش‌اند و امتیازها بلافاصله بازمحاسبه می‌شوند. ' +
      'خانه‌های زردرنگ ورودی کاربر و خانه‌های خاکستری محاسباتی هستند.',
      [
        btn('ورود فایل جدید', function () { go('import'); }, 'primary'),
        btn('خروجی', function () { exportSheet('questionnaire'); })
      ]));

    if (!App.state.questionnaires.length) {
      main.appendChild(el('div', { class: 'empty' }, [
        el('div', { class: 'big', text: '📝' }),
        el('div', { text: 'هنوز پرسشنامه‌ای وارد نشده است.' }),
        el('div', { style: 'margin-top:11px' }, [btn('ورود پرسشنامه', function () { go('import'); }, 'primary')])
      ]));
      return;
    }

    main.appendChild(el('div', { class: 'legend mb' }, [
      el('span', { html: '<i class="edit"></i> ورودی کاربر' }),
      el('span', { html: '<i class="calc"></i> محاسباتی (غیرقابل ویرایش)' }),
      el('span', { class: 'muted', text: 'حداکثر امتیاز ' + App.state.config.maxPerformanceScore +
        ' • تعداد سؤالات محاسباتی ' + App.state.config.questionCount +
        ' • حد نصاب ' + App.state.config.minPerformanceThreshold })
    ]));

    var answerOptions = Object.keys(App.state.config.answerScale);

    function answerCell(qKey) {
      return function (r) {
        var q = questionnaireByKey(r._input._key);
        var sel = el('select', { class: 'cell' });
        sel.appendChild(el('option', { value: '', text: '—' }));
        answerOptions.forEach(function (o) {
          sel.appendChild(el('option', { value: o, text: o + ' (' + App.state.config.answerScale[o] + ')' }));
        });
        var cur = q ? q[qKey] : '';
        sel.value = (cur === null || cur === undefined) ? '' : String(cur);
        if (sel.value === '' && cur !== '' && cur !== null && cur !== undefined) {
          /* wording the scale does not recognise — keep it visible rather than
             silently blanking a real answer */
          sel.appendChild(el('option', { value: String(cur), text: String(cur) + ' ⚠' }));
          sel.value = String(cur);
          sel.classList.add('invalid');
        }
        sel.addEventListener('change', function () {
          editField(q, qKey, sel.value === '' ? null : sel.value, 'ویرایش پاسخ پرسشنامه');
        });
        return sel;
      };
    }

    var grid = U.DataGrid({
      title: 'مجموعه یکپارچه پرسشنامه‌ها',
      rows: App.result.rows,
      sortKey: 'employeeId',
      searchFields: ['employeeId', 'fullName', 'division', 'positionTitle'],
      facets: [
        { key: 'division', label: 'همه واحدها' },
        { key: 'sourceFile', label: 'همه فایل‌ها', value: function (r) { return r._input.sourceFile; } },
        { key: 'status', label: 'همه وضعیت‌ها' }
      ],
      rowClass: function (r) {
        if (r.excluded) return 'row-excluded';
        if (!r.hasQuestionnaire) return 'row-err';
        if (!r.eligible) return 'row-warn';
        return '';
      },
      columns: [
        { key: 'employeeId', label: 'شماره پرسنلی', alwaysVisible: true, width: '95px', group: 'شناسایی' },
        { key: 'fullName', label: 'نام و نام خانوادگی', width: '150px', group: 'شناسایی' },
        { key: 'division', label: 'واحد سازمانی', group: 'شناسایی' },
        { key: 'positionTitle', label: 'عنوان شغلی', width: '150px', group: 'شناسایی', hidden: true },
        { key: 'jobLevel', label: 'JL', width: '48px', group: 'شناسایی' },
        { key: 'q1', label: 'Q1', group: 'پاسخ سؤالات', editable: true, width: '108px', render: answerCell('q1') },
        { key: 'q2', label: 'Q2', group: 'پاسخ سؤالات', editable: true, width: '108px', render: answerCell('q2') },
        { key: 'q3', label: 'Q3', group: 'پاسخ سؤالات', editable: true, width: '108px', render: answerCell('q3') },
        { key: 'q4', label: 'Q4', group: 'پاسخ سؤالات', editable: true, width: '108px', render: answerCell('q4') },
        { key: 'q5', label: 'Q5 (اطلاعاتی)', group: 'پاسخ سؤالات', editable: true, width: '108px',
          title: 'در محاسبه امتیاز عملکرد وارد نمی‌شود — مطابق فایل مرجع.',
          render: answerCell('q5') },
        { key: 'specialProject', label: 'اثرگذاری ویژه', group: 'اثرگذاری ویژه', editable: true,
          render: function (r) {
            var q = questionnaireByKey(r._input._key);
            var cb = el('input', { type: 'checkbox' });
            cb.checked = !!r.specialProject;
            cb.addEventListener('change', function () {
              editField(q, 'specialProject', cb.checked ? 'بله' : null, 'تغییر وضعیت اثرگذاری ویژه');
            });
            return cb;
          } },
        { key: 'specialImpactValue', label: 'امتیاز ویژه', type: 'score', decimals: 0,
          group: 'اثرگذاری ویژه', editable: true,
          render: function (r) {
            var q = questionnaireByKey(r._input._key);
            var inp = el('input', { type: 'number', class: 'cell', step: '1',
              /* The default only applies once the flag is set; showing it on a
                 disabled cell would read as a live value. */
              placeholder: r.specialProject ? String(App.state.config.specialImpactAmount) : '—' });
            inp.value = q && q.specialImpactAmount ? q.specialImpactAmount : '';
            inp.disabled = !r.specialProject;
            inp.addEventListener('change', function () {
              editField(q, 'specialImpactAmount', inp.value === '' ? null : Number(inp.value),
                'تغییر امتیاز اثرگذاری ویژه');
            });
            return inp;
          } },
        { key: 'performanceScore', label: 'امتیاز عملکرد', type: 'score', calculated: true,
          group: 'محاسبات', title: 'ستون K فایل مرجع — میانگین Q1..Q4' },
        { key: 'performanceKaraneh', label: 'عدد کارانه', type: 'score', decimals: 2, calculated: true,
          group: 'محاسبات', title: 'ستون L — امتیاز × ۱۲۰ ÷ ۴' },
        { key: 'rawCoefficient', label: 'ضریب کارانه', type: 'score', decimals: 2, calculated: true,
          group: 'محاسبات', title: 'ستون O — عدد کارانه + امتیاز ویژه' },
        { key: 'finalCoefficient', label: 'ضریب نهایی', type: 'score', decimals: 2, calculated: true,
          group: 'محاسبات', title: 'ستون R — پس از متناسب‌سازی با تعداد افراد و امتیاز کل' },
        { key: 'status', label: 'وضعیت', group: 'محاسبات',
          render: function (r) { return statusChip(r); } },
        { key: 'sourceFile', label: 'فایل منبع', hidden: true,
          value: function (r) { return r._input.sourceFile; } },
        { key: 'act', label: '', render: function (r) {
          return el('button', { class: 'btn sm', text: 'جزئیات',
            onclick: function () { showEmployeeDetail(r.employeeId); } });
        } }
      ],
      footer: function (visible) {
        var raw = 0, fin = 0, n = 0;
        visible.forEach(function (r) {
          if (!r.inScope) return;
          raw += r.rawCoefficient; fin += r.finalCoefficient; n++;
        });
        return {
          employeeId: 'جمع (' + n + ')',
          rawCoefficient: U.score(raw, 2),
          finalCoefficient: U.score(fin, 2)
        };
      },
      onRowClick: function (r) { showEmployeeDetail(r.employeeId); }
    });
    App.grids.questionnaires = grid;
    main.appendChild(grid.node);
  };

  var STATUS_LABELS = {
    'Calculated':      ['ok',    'محاسبه شد'],
    'HOD Adjusted':    ['info',  'تغییر معاون بخش'],
    'Below Threshold': ['warn',  'زیر حد نصاب'],
    'Incomplete':      ['err',   'ناقص'],
    'Pending':         ['',      'در انتظار'],
    'Exception':       ['err',   'استثنا'],
    'Finalized':       ['brand', 'نهایی شد']
  };

  function statusLabel(status) {
    return (STATUS_LABELS[status] || ['', status])[1];
  }

  function statusChip(r) {
    var m = STATUS_LABELS[r.status] || ['', r.status];
    return el('span', { class: 'chip ' + m[0], text: m[1] });
  }

  /** The one place a questionnaire field changes: audit, persist, recompute. */
  function editField(q, field, value, reason) {
    if (!q) return;
    var old = q[field];
    if (String(old === null || old === undefined ? '' : old) ===
        String(value === null || value === undefined ? '' : value)) return;
    q[field] = value;
    Store.audit(App.state, {
      entity: 'questionnaire', employeeId: q.employeeId, employeeName: q.fullName,
      field: field, oldValue: old, newValue: value, reason: reason || 'ویرایش دستی'
    });
    save();
    recalc();
  }

  /* ======================================================================
   * VIEW — روش پرداخت کارانه
   * ====================================================================*/
  VIEWS.payment = function (main) {
    var t = App.result.totals;
    main.appendChild(head('روش پرداخت کارانه',
      'معادل شیت «روش پرداخت کارانه» فایل مرجع. ترتیب و معنای ستون‌ها حفظ شده است.',
      [
        btn('خروجی', function () { exportSheet('payment'); }),
        btn('تنظیم بودجه', function () { go('settings'); })
      ]));

    if (!App.result.rows.length) {
      main.appendChild(el('div', { class: 'empty' }, [
        el('div', { class: 'big', text: '💰' }),
        el('div', { text: 'برای محاسبه، ابتدا پرسشنامه‌ها را وارد کنید.' })
      ]));
      return;
    }

    var strip = el('div', { class: 'kpi-grid' });
    strip.appendChild(U.kpi('بودجه (ریال)', U.money(t.budget), { kind: 'brand' }));
    strip.appendChild(U.kpi('امتیاز کل', U.score(t.sumTotalScore, 2),
      { kind: 'info', sub: 'مجموع ستون «امتیاز کل»' }));
    strip.appendChild(U.kpi('ضریب تأثیر گرید', String(App.state.config.gradeImpactFactor),
      { sub: App.state.config.gradeImpactFactor === 0 ? 'گرید در حال حاضر بی‌اثر است' : '' }));
    strip.appendChild(U.kpi('سرشکن زیر حد نصاب', U.score(t.ineligibleRedistribution, 4),
      { sub: 'به ازای هر فرد واجد شرایط' }));
    strip.appendChild(U.kpi('سرشکن تغییرات معاون', U.money(t.hodRedistribution),
      { kind: t.hodRedistribution < 0 ? 'warn' : '' }));
    main.appendChild(strip);

    if (App.state.config.gradeImpactFactor === 0) {
      main.appendChild(U.alert('info', 'ضریب تأثیر گرید برابر صفر است',
        'در فایل مرجع (سلول D2) نیز این مقدار صفر بوده و به همین دلیل سطح شغلی روی مبلغ کارانه اثری ندارد. ' +
        'برای فعال‌کردن اثر گرید، این ضریب را در تنظیمات تغییر دهید.',
        btn('تنظیمات', function () { go('settings'); }, 'sm')));
    }

    var grid = U.DataGrid({
      title: 'جدول پرداخت کارانه',
      rows: App.result.rows,
      sortKey: 'finalKaraneh', sortDir: 'desc',
      searchFields: ['employeeId', 'fullName', 'division', 'positionTitle'],
      facets: [
        { key: 'division', label: 'همه واحدها' },
        { key: 'jobLevel', label: 'همه سطوح شغلی' },
        { key: 'status', label: 'همه وضعیت‌ها' }
      ],
      rowClass: function (r) {
        if (r.negative) return 'row-err';
        if (r.excluded) return 'row-excluded';
        if (!r.eligible && r.inScope) return 'row-warn';
        return '';
      },
      columns: [
        { key: 'employeeId', label: 'شماره پرسنلی', alwaysVisible: true, width: '95px' },
        { key: 'fullName', label: 'نام و نام خانوادگی', width: '150px' },
        { key: 'division', label: 'واحد سازمانی' },
        { key: 'positionTitle', label: 'عنوان شغلی', width: '150px', hidden: true },
        { key: 'jobLevel', label: 'JL', width: '48px' },
        { key: 'gradeScore', label: 'عدد گرید', type: 'score', decimals: 0, calculated: true,
          title: 'ستون F — VLOOKUP از جدول گرید' },
        { key: 'gradeImpact', label: 'تأثیر گرید', type: 'score', decimals: 2, calculated: true,
          title: 'ستون G — عدد گرید × ضریب تأثیر گرید' },
        { key: 'evalScore', label: 'عدد ارزیابی', type: 'score', calculated: true, title: 'ستون H' },
        { key: 'eligibleEvalScore', label: 'ارزیابی مؤثر', type: 'score', calculated: true,
          title: 'ستون I — صفر اگر امتیاز از حد نصاب عبور نکند', hidden: true },
        { key: 'performanceContribution', label: 'امتیاز عملکردی', type: 'score', decimals: 2,
          calculated: true, title: 'ستون K — ضریب نهایی + سرشکن زیر حد نصاب' },
        { key: 'totalScore', label: 'امتیاز کل', type: 'score', decimals: 2, calculated: true,
          title: 'ستون L — تأثیر گرید + امتیاز عملکردی' },
        { key: 'initialAllocation', label: 'دریافتی قبل از تغییرات', type: 'money', calculated: true,
          title: 'ستون M — امتیاز کل ÷ مجموع امتیازها × بودجه' },
        { key: 'hodAdjustment', label: 'تغییرات معاون بخش', type: 'money', editable: true,
          title: 'ستون N',
          render: function (r) {
            if (!r.inScope || !r.eligible) return el('span', { class: 'muted', text: '—' });
            return el('button', {
              class: 'btn sm' + (r.isOverridden ? ' primary' : ''),
              text: r.isOverridden ? U.money(r.hodAdjustment) : 'تعیین',
              onclick: function () { openHodEditor(r.employeeId); }
            });
          } },
        { key: 'diff', label: 'Diff', type: 'money', calculated: true, hidden: true,
          title: 'ستون O — دریافتی اولیه منهای مبلغ معاون بخش' },
        { key: 'finalKaraneh', label: 'کارانه نهایی (ریال)', type: 'money', calculated: true,
          title: 'ستون Q — دریافتی نهایی',
          className: function (r) { return r.negative ? 'neg' : ''; } },
        { key: 'status', label: 'وضعیت', render: function (r) { return statusChip(r); } },
        { key: 'hodComment', label: 'توضیح معاون بخش', width: '160px' }
      ],
      footer: function (visible) {
        var sc = 0, alloc = 0, fin = 0, n = 0;
        visible.forEach(function (r) {
          if (!r.inScope) return;
          sc += r.totalScore; alloc += r.initialAllocation; fin += r.finalKaraneh; n++;
        });
        return {
          employeeId: 'جمع (' + n + ')',
          totalScore: U.score(sc, 2),
          initialAllocation: U.money(alloc),
          finalKaraneh: U.money(fin)
        };
      },
      onRowClick: function (r) { showEmployeeDetail(r.employeeId); }
    });
    App.grids.payment = grid;
    main.appendChild(grid.node);
  };

  /* ======================================================================
   * VIEW — HOD adjustments
   * ====================================================================*/
  VIEWS.hod = function (main) {
    var t = App.result.totals;
    main.appendChild(head('تغییرات معاون بخش',
      'معاون بخش می‌تواند مبلغ نهایی هر فرد را تعیین کند. ثبت توضیح اجباری است و ' +
      'اختلاف مبلغ بین سایر افراد سرشکن می‌شود تا مجموع پرداخت از بودجه عبور نکند.'));

    var overridden = App.result.rows.filter(function (r) { return r.isOverridden; });
    var sumOverride = 0;
    overridden.forEach(function (r) { sumOverride += r.hodAdjustment; });

    var strip = el('div', { class: 'kpi-grid' });
    strip.appendChild(U.kpi('تعداد تغییرات', U.int(overridden.length), { kind: 'info' }));
    strip.appendChild(U.kpi('مجموع مبالغ تعیین‌شده', U.moneyShort(sumOverride), { kind: 'brand' }));
    strip.appendChild(U.kpi('سرشکن بر سایرین', U.money(t.hodRedistribution),
      { kind: t.hodRedistribution < 0 ? 'warn' : 'ok', sub: 'به ازای هر نفر' }));
    strip.appendChild(U.kpi('دریافتی منفی', U.int(t.negativePayoutCount),
      { kind: t.negativePayoutCount ? 'err' : 'ok' }));
    strip.appendChild(U.kpi('بدون توضیح',
      U.int(overridden.filter(function (r) { return !r.hodComment; }).length),
      { kind: overridden.some(function (r) { return !r.hodComment; }) ? 'err' : 'ok' }));
    main.appendChild(strip);

    if (t.negativePayoutCount) {
      main.appendChild(U.alert('err', 'تغییرات فعلی سهم سایر افراد را منفی کرده است',
        t.negativePayoutCount + ' نفر با اعمال این تغییرات دریافتی منفی پیدا کرده‌اند. ' +
        'مبالغ تعیین‌شده باید کاهش یابد یا بودجه افزایش پیدا کند. تا رفع این مورد نهایی‌سازی ممکن نیست.'));
    }

    var eligible = App.result.rows.filter(function (r) { return r.inScope && r.eligible; });
    var grid = U.DataGrid({
      title: 'تعیین مبلغ توسط معاون بخش',
      rows: eligible,
      sortKey: 'initialAllocation', sortDir: 'desc',
      searchFields: ['employeeId', 'fullName', 'division', 'positionTitle'],
      facets: [
        { key: 'division', label: 'همه واحدها' },
        { key: 'adjusted', label: 'همه', value: function (r) { return r.isOverridden ? 'تغییر یافته' : 'بدون تغییر'; } }
      ],
      rowClass: function (r) {
        if (r.negative) return 'row-err';
        if (r.isOverridden && !r.hodComment) return 'row-err';
        if (r.isOverridden) return 'row-warn';
        return '';
      },
      columns: [
        { key: 'employeeId', label: 'شماره پرسنلی', alwaysVisible: true, width: '95px' },
        { key: 'fullName', label: 'نام و نام خانوادگی', width: '150px' },
        { key: 'division', label: 'واحد سازمانی' },
        { key: 'jobLevel', label: 'JL', width: '48px' },
        { key: 'totalScore', label: 'امتیاز کل', type: 'score', decimals: 2, calculated: true },
        { key: 'initialAllocation', label: 'دریافتی محاسباتی', type: 'money', calculated: true },
        { key: 'hodAdjustment', label: 'مبلغ تعیین‌شده', type: 'money', editable: true,
          render: function (r) {
            return el('button', {
              class: 'btn sm' + (r.isOverridden ? ' primary' : ''),
              text: r.isOverridden ? U.money(r.hodAdjustment) : '＋ تعیین',
              onclick: function () { openHodEditor(r.employeeId); }
            });
          } },
        { key: 'delta', label: 'اختلاف', type: 'money', calculated: true,
          value: function (r) { return r.isOverridden ? -r.diff : null; },
          className: function (r) { return (r.isOverridden && -r.diff < 0) ? 'neg' : ''; } },
        { key: 'finalKaraneh', label: 'کارانه نهایی', type: 'money', calculated: true,
          className: function (r) { return r.negative ? 'neg' : ''; } },
        { key: 'hodComment', label: 'توضیح (اجباری)', width: '200px',
          render: function (r) {
            if (!r.isOverridden) return el('span', { class: 'muted', text: '—' });
            if (!r.hodComment) return el('span', { class: 'chip err', text: 'توضیح ثبت نشده' });
            return el('span', { text: r.hodComment, title: r.hodComment });
          } },
        { key: 'act', label: '', render: function (r) {
          if (!r.isOverridden) return document.createTextNode('');
          return el('button', { class: 'btn sm danger', text: 'حذف',
            onclick: function () { clearHod(r.employeeId); } });
        } }
      ],
      footer: function (visible) {
        var alloc = 0, fin = 0;
        visible.forEach(function (r) { alloc += r.initialAllocation; fin += r.finalKaraneh; });
        return {
          employeeId: 'جمع (' + visible.length + ')',
          initialAllocation: U.money(alloc), finalKaraneh: U.money(fin)
        };
      }
    });
    main.appendChild(grid.node);
  };

  function openHodEditor(employeeId) {
    var r = resultRow(employeeId);
    if (!r || !r.eligible) { U.toast('برای این فرد امکان تعیین مبلغ وجود ندارد.', 'warn'); return; }
    var ceiling = Engine.maxAllowedAdjustment(App.result, employeeId);

    var amount = el('input', {
      type: 'number', class: 'editable', style: 'width:100%', step: '1000000',
      value: r.isOverridden ? String(r.hodAdjustment) : ''
    });
    var comment = el('textarea', {
      class: 'editable', style: 'width:100%;min-height:70px',
      placeholder: 'بر اساس چه اثرگذاری خاصی این مبلغ تعیین شده است؟'
    });
    comment.value = r.hodComment || '';

    var preview = el('div', {});
    function renderPreview() {
      U.clear(preview);
      var v = amount.value === '' ? null : Number(amount.value);
      var others = App.result.rows.filter(function (x) {
        return x.inScope && x.eligible && !x.isOverridden && x.employeeId !== employeeId;
      });
      /* How the pot shifts if this override is committed: the extra (or the
         saving) is shared equally by everyone who has no override. */
      var delta = v === null ? 0 : (r.initialAllocation - v) - (r.isOverridden ? r.diff : 0);
      var perPerson = others.length ? delta / others.length : 0;
      var minPeer = others.reduce(function (m, x) {
        return Math.min(m, x.finalKaraneh + perPerson);
      }, Infinity);

      var dl = el('dl', { class: 'kv' });
      [
        ['دریافتی محاسباتی', U.money(r.initialAllocation)],
        ['مبلغ تعیین‌شده', v === null ? '—' : U.money(v)],
        ['اثر بر هر یک از ' + others.length + ' نفر دیگر', U.money(perPerson)],
        ['سقف مجاز این مبلغ', ceiling === null ? '—' : U.money(ceiling)]
      ].forEach(function (l) {
        dl.appendChild(el('dt', { text: l[0] }));
        dl.appendChild(el('dd', { class: 'num', text: l[1] }));
      });
      preview.appendChild(dl);

      if (v !== null && others.length && minPeer < 0) {
        preview.appendChild(U.alert('err', 'این مبلغ بیش از حد مجاز است',
          'با این مبلغ، دریافتی حداقل یک نفر دیگر منفی می‌شود (' + U.money(minPeer) + '). ' +
          'حداکثر مبلغ قابل تعیین ' + U.money(ceiling) + ' ریال است.'));
      } else if (v !== null && v > r.initialAllocation) {
        preview.appendChild(U.alert('warn', 'افزایش نسبت به مبلغ محاسباتی',
          'مبلغ ' + U.money(v - r.initialAllocation) + ' ریال بیشتر از سهم محاسباتی است و ' +
          'از سهم ' + others.length + ' نفر دیگر تأمین می‌شود.'));
      } else if (v !== null && v < r.initialAllocation) {
        preview.appendChild(U.alert('info', 'کاهش نسبت به مبلغ محاسباتی',
          'مبلغ ' + U.money(r.initialAllocation - v) + ' ریال بین ' + others.length +
          ' نفر دیگر توزیع می‌شود.'));
      }
    }
    amount.addEventListener('input', renderPreview);
    renderPreview();

    U.modal({
      title: 'تعیین مبلغ — ' + (r.fullName || employeeId) + '  (' + employeeId + ')',
      content: el('div', { class: 'split' }, [
        el('div', {}, [
          el('label', { class: 'field' }, [
            el('span', { html: 'مبلغ نهایی کارانه (ریال) <b>*</b>' }), amount
          ]),
          el('label', { class: 'field' }, [
            el('span', { html: 'توضیح معاون بخش <b>*</b> — الزامی' }), comment
          ])
        ]),
        el('div', {}, [
          el('div', { class: 'small muted mb', text: 'اثر این تغییر بر بودجه' }),
          preview
        ])
      ]),
      buttons: [
        { label: 'ثبت', kind: 'primary', keepOpen: true, onClick: function (close) {
          var v = amount.value === '' ? null : Number(amount.value);
          if (v === null) { U.toast('مبلغ را وارد کنید.', 'err'); return false; }
          if (!isFinite(v) || v < 0) { U.toast('مبلغ باید عددی مثبت باشد.', 'err'); return false; }
          if (!comment.value.trim()) {
            U.toast('ثبت توضیح برای تغییر مبلغ اجباری است.', 'err', 5000); return false;
          }
          applyHod(employeeId, v, comment.value.trim());
          close();
          return false;
        } },
        r.isOverridden ? { label: 'حذف تغییر', kind: 'danger',
          onClick: function () { clearHod(employeeId); } } : null,
        'spacer',
        { label: 'انصراف' }
      ].filter(Boolean)
    });
  }

  function applyHod(employeeId, amount, comment) {
    var r = resultRow(employeeId);
    var q = questionnaireByKey(r._input._key);
    if (!q) return;
    Store.audit(App.state, {
      entity: 'hod', employeeId: employeeId, employeeName: r.fullName,
      field: 'hodAdjustment', oldValue: q.hodAdjustment, newValue: amount, reason: comment
    });
    q.hodAdjustment = amount;
    q.hodComment = comment;
    save();
    var check = Engine.validateBudget(recalc());
    if (!check.ok) U.toast(check.problems[0].message, 'err', 7000);
    else U.toast('مبلغ ثبت شد و محاسبات بازمحاسبه شد.', 'ok');
  }

  function clearHod(employeeId) {
    var r = resultRow(employeeId);
    var q = questionnaireByKey(r._input._key);
    if (!q) return;
    Store.audit(App.state, {
      entity: 'hod', employeeId: employeeId, employeeName: r.fullName,
      field: 'hodAdjustment', oldValue: q.hodAdjustment, newValue: null,
      reason: 'حذف تغییر معاون بخش'
    });
    q.hodAdjustment = null;
    q.hodComment = '';
    save();
    recalc();
    U.toast('تغییر حذف شد.', 'ok');
  }

  /* ======================================================================
   * Employee detail — the calculation breakdown
   * ====================================================================*/
  function showEmployeeDetail(employeeId) {
    var r = resultRow(employeeId);
    if (!r) { U.toast('رکورد محاسباتی برای این فرد وجود ندارد.', 'warn'); return; }
    var master = employeeById(employeeId);
    var cfg = App.state.config;
    var t = App.result.totals;

    /**
     * One row of the breakdown.
     * `formula` is arithmetic only and renders left-to-right; any Persian
     * explanation belongs in `note`, which renders right-to-left. Keeping the
     * two apart is what stops the bidi algorithm reordering the operands.
     */
    function step(name, value, formula, source, kind, note) {
      return el('div', { class: 'calc-step ' + (kind || 'is-calc') }, [
        el('div', { class: 'row' }, [
          el('span', { class: 'name', text: name }),
          el('span', { class: 'val', text: value })
        ]),
        formula ? el('div', { class: 'formula', text: formula }) : null,
        note ? el('div', { class: 'note', text: note }) : null,
        source ? el('div', { class: 'src', text: source }) : null
      ]);
    }

    function mapped(k) {
      var v = r[k];
      if (typeof v === 'number') return v;
      var m = cfg.answerScale[String(v).trim()];
      return m === undefined ? null : m;
    }

    var answers = ['q1', 'q2', 'q3', 'q4'].map(function (k, i) {
      var m = mapped(k);
      return 'Q' + (i + 1) + ' = ' + (r[k] === null || r[k] === undefined || r[k] === '' ? '—' : r[k]) +
             '  →  ' + (m === null ? '؟' : m);
    }).join('\n');

    var left = el('div', {});
    var dl = el('dl', { class: 'kv' });
    [
      ['شماره پرسنلی', employeeId],
      ['نام و نام خانوادگی', r.fullName || '—'],
      ['واحد سازمانی', r.division || '—'],
      ['عنوان شغلی', r.positionTitle || '—'],
      ['سطح شغلی (JL)', r.jobLevel || '—'],
      ['نوع استخدام', (master && master.employmentType) || '—'],
      ['روز کارکرد', (master && master.workingDays) || '—'],
      ['مدیر مستقیم', (master && master.directManager) || '—'],
      ['فایل منبع پرسشنامه', r._input.sourceFile || '—'],
      ['وضعیت', statusLabel(r.status)]
    ].forEach(function (p) {
      dl.appendChild(el('dt', { text: p[0] }));
      dl.appendChild(el('dd', { text: String(p[1]) }));
    });
    left.appendChild(U.card('مشخصات فرد', dl));

    left.appendChild(step('پاسخ سؤالات عملکرد', '', answers,
      'شیت «پرسشنامه کارانه تیمی» ستون‌های F تا I • نگاشت از جدول Data!H:I', 'is-input'));

    if (r.performanceScore !== null) {
      left.appendChild(step('امتیاز عملکرد', U.score(r.performanceScore),
        '(' + ['q1', 'q2', 'q3', 'q4'].map(mapped).join(' + ') + ') / ' + cfg.questionCount +
        ' = ' + U.score(r.performanceScore),
        'ستون K — میانگین چهار سؤال عملکردی. Q5 عمداً وارد محاسبه نمی‌شود.'));

      left.appendChild(step('عدد کارانه', U.score(r.performanceKaraneh, 2),
        U.score(r.performanceScore) + ' × ' + cfg.maxPerformanceScore + ' / ' + cfg.questionCount +
        ' = ' + U.score(r.performanceKaraneh, 2),
        'ستون L — تبدیل امتیاز ۱ تا ۵ به بازه کارانه.'));

      left.appendChild(step('اثرگذاری ویژه', U.score(r.specialImpactValue, 0),
        null, 'ستون N', r.specialProject ? 'is-input' : 'is-calc',
        r.specialProject
          ? 'اثرگذاری ویژه ثبت شده است، بنابراین ' + U.score(r.specialImpactValue, 0) +
            ' امتیاز به ضریب کارانه اضافه می‌شود.'
          : 'اثرگذاری ویژه ثبت نشده است، بنابراین امتیازی اضافه نمی‌شود.'));

      left.appendChild(step('ضریب کارانه', U.score(r.rawCoefficient, 2),
        U.score(r.performanceKaraneh, 2) + ' + ' + U.score(r.specialImpactValue, 0) +
        ' = ' + U.score(r.rawCoefficient, 2),
        'ستون O — امتیاز عملکرد به علاوه امتیاز اثرگذاری ویژه.'));

      var perPerson = -r.normalizationAdjustment;
      left.appendChild(step('ضریب نهایی کارانه', U.score(r.finalCoefficient, 2),
        U.score(r.rawCoefficient, 2) + ' − ' + U.score(perPerson, 4) + ' = ' +
        U.score(r.finalCoefficient, 2) +
        '\n\n(' + U.score(t.sumRawCoefficient, 2) + ' − ' +
        (t.inScopeCount * cfg.baselineCoefficientPerPerson) + ') / ' + t.inScopeCount +
        ' = ' + U.score(perPerson, 4),
        'ستون R — متناسب‌سازی نسبت به تعداد افراد و امتیاز کل بخش.', null,
        'مقدار متناسب‌سازی برابر است با «مجموع ضرایب منهای تعداد افراد ضربدر ' +
        cfg.baselineCoefficientPerPerson + '» تقسیم بر تعداد افراد، و از ضریب همه کسر می‌شود.'));
    } else {
      left.appendChild(U.alert('err', 'پاسخ سؤالات ناقص است',
        'تا تکمیل پاسخ‌ها امتیازی برای این فرد محاسبه نمی‌شود.'));
    }

    var right = el('div', {});
    right.appendChild(step('عدد گرید', r.gradeScore === null ? 'تعریف نشده' : U.score(r.gradeScore, 0),
      r.gradeScore === null ? null : 'getGradeScore("' + r.jobLevel + '") = ' + r.gradeScore,
      'ستون F — جایگزین VLOOKUP در جدول Data!C:D.',
      r.gradeScore === null ? 'is-input' : 'is-calc',
      r.gradeScore === null
        ? '⚠ سطح شغلی «' + r.jobLevel + '» در جدول گرید تعریف نشده است.' : null));

    right.appendChild(step('تأثیر گرید', U.score(r.gradeImpact, 2),
      U.score(r.gradeScore, 0) + ' × ' + cfg.gradeImpactFactor + ' = ' + U.score(r.gradeImpact, 2),
      'ستون G — عدد گرید × ضریب تأثیر گرید (سلول D2).'));

    right.appendChild(step('عبور از حد نصاب', r.eligible ? 'بله' : 'خیر',
      U.score(r.performanceScore) + ' ' + (cfg.thresholdMode === 'gte' ? '>=' : '>') + ' ' +
      cfg.minPerformanceThreshold + '  ->  ' + (r.eligible ? U.score(r.performanceScore) : '0'),
      'ستون I', r.eligible ? 'is-calc' : 'is-input',
      'امتیاز کمتر یا مساوی حد نصاب، کارانه فرد را صفر می‌کند.'));

    right.appendChild(step('امتیاز عملکردی', U.score(r.performanceContribution, 2),
      r.eligible
        ? U.score(r.finalCoefficient, 2) + ' + ' + U.score(t.ineligibleRedistribution, 4) +
          ' = ' + U.score(r.performanceContribution, 2) +
          '\n\n' + U.score(t.forfeitedCoefficient, 2) + ' / ' + t.eligibleCount +
          ' = ' + U.score(t.ineligibleRedistribution, 4)
        : '0',
      'ستون K', null,
      r.eligible
        ? 'سرشکن برابر است با مجموع ضرایب افراد زیر حد نصاب تقسیم بر تعداد واجدین شرایط، ' +
          'و به ضریب نهایی هر فرد واجد شرایط اضافه می‌شود.'
        : 'امتیاز این فرد از حد نصاب عبور نکرده است، بنابراین سهم عملکردی او صفر است.'));

    right.appendChild(step('امتیاز کل', U.score(r.totalScore, 2),
      U.score(r.gradeImpact, 2) + ' + ' + U.score(r.performanceContribution, 2) +
      ' = ' + U.score(r.totalScore, 2),
      'ستون L — تأثیر گرید + امتیاز عملکردی.'));

    right.appendChild(step('دریافتی قبل از تغییرات', U.money(r.initialAllocation),
      U.score(r.totalScore, 2) + ' / ' + U.score(t.sumTotalScore, 2) + ' × ' + U.money(t.budget) +
      '\n= ' + U.money(r.initialAllocation),
      'ستون M — تقسیم بودجه به نسبت امتیاز کل.'));

    if (r.isOverridden) {
      right.appendChild(step('تغییر معاون بخش', U.money(r.hodAdjustment),
        U.money(r.initialAllocation) + ' − ' + U.money(r.hodAdjustment) + ' = ' +
        U.money(r.initialAllocation - r.hodAdjustment),
        'ستون N', 'is-input',
        'اختلاف بالا بین سایر افراد بدون تغییر سرشکن می‌شود. ' +
        'توضیح ثبت‌شده: ' + (r.hodComment || '⚠ ثبت نشده')));
    } else if (t.overriddenCount) {
      right.appendChild(step('سرشکن تغییرات معاون', U.money(r.hodRedistribution),
        null, 'سلول O5', null,
        'مجموع اختلاف تغییرات معاون بخش تقسیم بر تعداد افراد بدون تغییر، ' +
        'که مجموع پرداخت را دقیقاً برابر بودجه نگه می‌دارد.'));
    }

    right.appendChild(step('کارانه نهایی', U.money(r.finalKaraneh),
      r.isOverridden ? null
        : U.money(r.initialAllocation) + ' + ' + U.money(r.hodRedistribution) +
          ' = ' + U.money(r.finalKaraneh),
      'ستون Q — دریافتی نهایی', 'is-final',
      r.isOverridden ? 'مبلغ مستقیماً توسط معاون بخش تعیین شده است.' : null));

    U.modal({
      title: 'جزئیات محاسبه — ' + (r.fullName || employeeId),
      size: 'wide',
      content: el('div', { class: 'calc-panel' }, [left, right]),
      buttons: [
        { label: 'تعیین مبلغ توسط معاون بخش', kind: 'primary', disabled: !r.eligible,
          onClick: function () { setTimeout(function () { openHodEditor(employeeId); }, 60); } },
        'spacer',
        { label: 'بستن' }
      ]
    });
  }
  App.showEmployeeDetail = showEmployeeDetail;

  /* ======================================================================
   * VIEW — Validation center
   * ====================================================================*/
  VIEWS.validation = function (main) {
    main.appendChild(head('مرکز اعتبارسنجی',
      'تمام مواردی که باید پیش از نهایی‌سازی بررسی یا برطرف شوند.',
      [
        btn('خروجی گزارش', function () { exportSheet('validation'); }),
        btn('نهایی‌سازی', function () { finalize(); }, 'primary')
      ]));

    var issues = App.validation || [];
    var byCode = {};
    issues.forEach(function (i) {
      var g = byCode[i.code] || (byCode[i.code] = { code: i.code, title: i.title, severity: i.severity, items: [] });
      g.items.push(i);
    });

    var strip = el('div', { class: 'kpi-grid' });
    strip.appendChild(U.kpi('خطا', U.int(issueCount('err')), { kind: issueCount('err') ? 'err' : 'ok' }));
    strip.appendChild(U.kpi('هشدار', U.int(issueCount('warn')), { kind: issueCount('warn') ? 'warn' : 'ok' }));
    strip.appendChild(U.kpi('اطلاع‌رسانی', U.int(issueCount('info')), { kind: 'info' }));
    strip.appendChild(U.kpi('وضعیت بودجه',
      App.result.totals.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر',
      { kind: App.result.totals.budgetStatus === 'BALANCED' ? 'ok' : 'err' }));
    main.appendChild(strip);

    if (!issues.length) {
      main.appendChild(U.alert('ok', 'هیچ مورد بازی وجود ندارد', 'سیستم آماده نهایی‌سازی است.'));
    }

    var order = { err: 0, warn: 1, info: 2 };
    Object.keys(byCode)
      .sort(function (a, b) { return order[byCode[a].severity] - order[byCode[b].severity]; })
      .forEach(function (code) {
        var g = byCode[code];
        var tbl = el('table', { class: 'grid' });
        tbl.appendChild(el('thead', {}, [el('tr', {}, [
          el('th', { text: 'شماره پرسنلی' }), el('th', { text: 'نام' }),
          el('th', { text: 'توضیح' }), el('th', { text: '' })
        ])]));
        var tb = el('tbody');
        g.items.slice(0, 500).forEach(function (i) {
          tb.appendChild(el('tr', {}, [
            el('td', { text: i.employeeId || '—' }),
            el('td', { text: i.employeeName || '—' }),
            el('td', { class: 'small', text: i.detail || '' }),
            el('td', {}, [i.employeeId
              ? el('button', { class: 'btn sm', text: 'بررسی',
                  onclick: function () { showEmployeeDetail(i.employeeId); } })
              : document.createTextNode('')])
          ]));
        });
        tbl.appendChild(tb);
        main.appendChild(U.card(
          (g.severity === 'err' ? '⛔ ' : g.severity === 'warn' ? '⚠️ ' : 'ℹ️ ') + g.title,
          el('div', { class: 'table-wrap', style: 'max-height:340px' }, [tbl]),
          { hint: g.items.length + ' مورد' +
            (g.items.length > 500 ? ' (۵۰۰ مورد اول نمایش داده شده)' : ''), tight: true }
        ));
      });
  };

  function finalize() {
    var errs = issueCount('err');
    var check = Engine.validateBudget(App.result);
    if (errs || !check.ok) {
      U.modal({
        title: 'نهایی‌سازی ممکن نیست', size: 'narrow',
        content: el('div', {}, [
          errs ? U.alert('err', errs + ' خطای باز وجود دارد',
            'تمام خطاهای مرکز اعتبارسنجی باید پیش از نهایی‌سازی برطرف شوند.') : null,
          check.ok ? null : U.alert('err', 'کنترل بودجه ناموفق',
            check.problems.map(function (p) { return p.message; }).join(' '))
        ].filter(Boolean)),
        buttons: [{ label: 'بستن', kind: 'primary' }]
      });
      return;
    }
    U.confirm('محاسبات نهایی شود؟\nمجموع پرداخت: ' + U.money(App.result.totals.sumFinalKaraneh) +
              ' ریال برای ' + App.result.totals.eligibleCount + ' نفر.',
      { confirmLabel: 'نهایی‌سازی' }).then(function (ok) {
      if (!ok) return;
      App.state.finalizedAt = new Date().toISOString();
      Store.audit(App.state, {
        entity: 'process', field: 'finalize', oldValue: '', newValue: App.state.finalizedAt,
        reason: 'نهایی‌سازی دوره ' + App.state.period
      });
      save().then(function () {
        recalc();
        U.toast('محاسبات نهایی شد.', 'ok');
        go('reports');
      });
    });
  }

  /* ======================================================================
   * VIEW — Audit log
   * ====================================================================*/
  VIEWS.audit = function (main) {
    main.appendChild(head('ردیابی تغییرات',
      'هر تغییری که می‌تواند بر مبلغ پرداختی اثر بگذارد اینجا ثبت می‌شود.',
      [btn('خروجی', function () { exportSheet('audit'); })]));

    if (!App.state.auditLog.length) {
      main.appendChild(el('div', { class: 'empty' }, [
        el('div', { class: 'big', text: '🧾' }), el('div', { text: 'هنوز تغییری ثبت نشده است.' })
      ]));
      return;
    }

    main.appendChild(U.DataGrid({
      title: 'گزارش تغییرات',
      rows: App.state.auditLog.slice().reverse(),
      searchFields: ['employeeId', 'employeeName', 'field', 'reason', 'user'],
      facets: [
        { key: 'entity', label: 'همه بخش‌ها' },
        { key: 'field', label: 'همه فیلدها' }
      ],
      columns: [
        { key: 'timestamp', label: 'زمان', width: '130px',
          render: function (r) { return U.dateTime(r.timestamp); } },
        { key: 'user', label: 'کاربر' },
        { key: 'entity', label: 'بخش' },
        { key: 'employeeId', label: 'شماره پرسنلی' },
        { key: 'employeeName', label: 'نام' },
        { key: 'field', label: 'فیلد' },
        { key: 'oldValue', label: 'مقدار قبلی', width: '130px' },
        { key: 'newValue', label: 'مقدار جدید', width: '130px' },
        { key: 'reason', label: 'دلیل', width: '220px' }
      ]
    }).node);
  };

  /* ======================================================================
   * VIEW — Settings
   * ====================================================================*/
  VIEWS.settings = function (main) {
    var cfg = App.state.config;
    main.appendChild(head('تنظیمات',
      'هیچ‌یک از این مقادیر در کد ثابت نشده است. تغییر هر کدام، کل محاسبات را بلافاصله بازمحاسبه می‌کند.'));

    function numberField(label, key, step, hint) {
      var inp = el('input', { type: 'number', class: 'editable', style: 'width:100%', step: step || 'any' });
      inp.value = cfg[key];
      inp.addEventListener('change', function () {
        var v = Number(inp.value);
        if (!isFinite(v)) { inp.value = cfg[key]; return; }
        setConfig(key, v);
      });
      return el('label', { class: 'field' }, [
        el('span', { html: U.esc(label) + (hint ? ' <span class="muted small">— ' + U.esc(hint) + '</span>' : '') }),
        inp
      ]);
    }

    var params = el('div', { class: 'form-grid' }, [
      numberField('بودجه کل (ریال)', 'budget', '1000000', 'سلول C1 فایل مرجع'),
      numberField('حداکثر امتیاز کارانه', 'maxPerformanceScore', '1', 'مقدار مرجع: 120'),
      numberField('تعداد سؤالات عملکردی', 'questionCount', '1', 'مقدار مرجع: 4'),
      numberField('حداقل امتیاز جهت دریافت', 'minPerformanceThreshold', '0.25', 'سلول D4'),
      numberField('ضریب تأثیر گرید', 'gradeImpactFactor', '0.1', 'سلول D2 — مقدار مرجع: 0'),
      numberField('امتیاز اثرگذاری ویژه', 'specialImpactAmount', '10', 'مقدار مرجع: 300'),
      numberField('ضریب پایه هر نفر', 'baselineCoefficientPerPerson', '10', 'B5 ÷ A5 — مقدار مرجع: 100')
    ]);

    var mode = el('select', { class: 'editable' }, [
      el('option', { value: 'gt', text: 'بزرگ‌تر از حد نصاب (مطابق فایل مرجع)' }),
      el('option', { value: 'gte', text: 'بزرگ‌تر یا مساوی حد نصاب' })
    ]);
    mode.value = cfg.thresholdMode;
    mode.addEventListener('change', function () { setConfig('thresholdMode', mode.value); });

    var scope = el('select', { class: 'editable' }, [
      el('option', { value: 'global', text: 'کل سازمان (مطابق فایل مرجع)' }),
      el('option', { value: 'division', text: 'به تفکیک واحد سازمانی' })
    ]);
    scope.value = cfg.normalizationScope;
    scope.addEventListener('change', function () { setConfig('normalizationScope', scope.value); });

    params.appendChild(el('label', { class: 'field' }, [
      el('span', { html: 'نحوه اعمال حد نصاب <span class="muted small">— ستون I</span>' }), mode
    ]));
    params.appendChild(el('label', { class: 'field' }, [
      el('span', { html: 'دامنه متناسب‌سازی ضرایب <span class="muted small">— «امتیاز کل بخش»</span>' }), scope
    ]));

    [['normalizeCoefficients', 'متناسب‌سازی ضرایب (ستون R)'],
     ['redistributeIneligible', 'سرشکن ضریب افراد زیر حد نصاب (سلول J5)'],
     ['redistributeHodDiff', 'سرشکن اختلاف تغییرات معاون بخش (سلول O5)']
    ].forEach(function (p) {
      var cb = el('input', { type: 'checkbox' });
      cb.checked = !!cfg[p[0]];
      cb.addEventListener('change', function () { setConfig(p[0], cb.checked); });
      params.appendChild(el('label', { class: 'checkline' }, [cb, el('span', { text: p[1] })]));
    });
    main.appendChild(U.card('بودجه و پارامترها', params));

    /* -- grade table --------------------------------------------------- */
    var gradeBody = el('div', {});
    var gt = el('table', { class: 'grid' });
    gt.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', { text: 'JL' }), el('th', { text: 'عدد گرید' }), el('th', { text: '' })
    ])]));
    var gtb = el('tbody');
    Object.keys(cfg.gradeMap).forEach(function (jl) {
      var inp = el('input', { type: 'number', class: 'cell', step: '10' });
      inp.value = cfg.gradeMap[jl];
      inp.addEventListener('change', function () {
        var old = cfg.gradeMap[jl];
        cfg.gradeMap[jl] = Number(inp.value);
        auditConfig('gradeMap.' + jl, old, cfg.gradeMap[jl]);
        save(); recalc({ repaint: false }); renderTopStats();
      });
      gtb.appendChild(el('tr', {}, [
        el('td', { class: 'mono', text: jl }),
        el('td', {}, [inp]),
        el('td', {}, [el('button', { class: 'btn sm danger', text: 'حذف', onclick: function () {
          var old = cfg.gradeMap[jl];
          delete cfg.gradeMap[jl];
          auditConfig('gradeMap.' + jl, old, '(حذف شد)');
          save(); recalc();
        } })])
      ]));
    });
    gt.appendChild(gtb);
    gradeBody.appendChild(gt);

    var newJl = el('input', { type: 'text', placeholder: 'JL مثلاً 2H', style: 'width:110px' });
    var newScore = el('input', { type: 'number', placeholder: 'عدد گرید', style: 'width:110px' });
    gradeBody.appendChild(el('div', { style: 'display:flex;gap:7px;margin-top:10px;align-items:center' }, [
      newJl, newScore,
      btn('افزودن', function () {
        var k = Engine.normalizeJobLevel(newJl.value);
        var v = Number(newScore.value);
        if (!k || !isFinite(v)) { U.toast('سطح شغلی و عدد گرید را وارد کنید.', 'err'); return; }
        cfg.gradeMap[k] = v;
        auditConfig('gradeMap.' + k, '(جدید)', v);
        newJl.value = ''; newScore.value = '';
        save(); recalc();
      }, 'primary')
    ]));

    var unmappedJl = {};
    App.result.rows.forEach(function (r) {
      if (r.inScope && r.gradeScore === null && r.jobLevel) unmappedJl[r.jobLevel] = 1;
    });
    var missingJl = Object.keys(unmappedJl);
    if (missingJl.length) {
      gradeBody.appendChild(U.alert('err', 'سطوح شغلی تعریف‌نشده',
        'این سطوح در داده‌ها وجود دارند اما در جدول گرید نیستند: ' + missingJl.join('، ') +
        ' — تا تعریف نشوند، امتیاز گرید این افراد محاسبه نمی‌شود.'));
    }
    main.appendChild(U.card('جدول گرید (JL → عدد گرید)', gradeBody,
      { hint: 'جایگزین VLOOKUP جدول Data!C:D' }));

    /* -- answer scale -------------------------------------------------- */
    var scaleBody = el('div', {});
    var st = el('table', { class: 'grid' });
    st.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', { text: 'پاسخ' }), el('th', { text: 'امتیاز' }), el('th', { text: '' })
    ])]));
    var stb = el('tbody');
    Object.keys(cfg.answerScale).forEach(function (k) {
      var inp = el('input', { type: 'number', class: 'cell', step: '1' });
      inp.value = cfg.answerScale[k];
      inp.addEventListener('change', function () {
        var old = cfg.answerScale[k];
        cfg.answerScale[k] = Number(inp.value);
        auditConfig('answerScale.' + k, old, cfg.answerScale[k]);
        save(); recalc();
      });
      stb.appendChild(el('tr', {}, [
        el('td', { text: k }), el('td', {}, [inp]),
        el('td', {}, [el('button', { class: 'btn sm danger', text: 'حذف', onclick: function () {
          var old = cfg.answerScale[k];
          delete cfg.answerScale[k];
          auditConfig('answerScale.' + k, old, '(حذف شد)');
          save(); recalc();
        } })])
      ]));
    });
    st.appendChild(stb);
    scaleBody.appendChild(st);
    var newAns = el('input', { type: 'text', placeholder: 'متن پاسخ', style: 'width:150px' });
    var newVal = el('input', { type: 'number', placeholder: 'امتیاز', style: 'width:90px' });
    scaleBody.appendChild(el('div', { style: 'display:flex;gap:7px;margin-top:10px' }, [
      newAns, newVal,
      btn('افزودن', function () {
        var k = newAns.value.trim(); var v = Number(newVal.value);
        if (!k || !isFinite(v)) { U.toast('متن پاسخ و امتیاز را وارد کنید.', 'err'); return; }
        cfg.answerScale[k] = v;
        auditConfig('answerScale.' + k, '(جدید)', v);
        newAns.value = ''; newVal.value = '';
        save(); recalc();
      }, 'primary')
    ]));
    main.appendChild(U.card('نگاشت پاسخ به امتیاز', scaleBody, { hint: 'جایگزین جدول Data!H:I' }));

    /* -- column synonyms ---------------------------------------------- */
    var mapBody = el('div', {});
    mapBody.appendChild(el('p', { class: 'small muted',
      text: 'برای هر فیلد سیستم، عناوینی که هنگام ورود فایل به آن نگاشت می‌شوند. ' +
            'با کاما جدا کنید. افزودن یک املای جدید نیازی به تغییر کد ندارد.' }));
    Object.keys(App.state.columnMappings).forEach(function (field) {
      var m = App.state.columnMappings[field];
      var inp = el('input', { type: 'text', class: 'editable', style: 'width:100%' });
      inp.value = m.synonyms.join('، ');
      inp.addEventListener('change', function () {
        var old = m.synonyms.slice();
        m.synonyms = inp.value.split(/[،,]/).map(function (s) { return s.trim(); })
          .filter(function (s) { return s; });
        auditConfig('columnMapping.' + field, old.length + ' مورد', m.synonyms.length + ' مورد');
        save();
      });
      mapBody.appendChild(el('label', { class: 'field' }, [
        el('span', { html: '<b>' + U.esc(m.label) + '</b> <span class="muted mono small">' + field + '</span>' }),
        inp
      ]));
    });
    mapBody.appendChild(btn('بازنشانی نگاشت‌ها به حالت پیش‌فرض', function () {
      U.confirm('تمام نگاشت‌های ستون به حالت پیش‌فرض بازگردد؟').then(function (ok) {
        if (!ok) return;
        App.state.columnMappings = cloneMappings();
        auditConfig('columnMappings', 'custom', 'default');
        save().then(function () { renderView(); U.toast('بازنشانی شد.', 'ok'); });
      });
    }, 'danger'));
    main.appendChild(U.card('نگاشت هوشمند ستون‌ها', mapBody, { hint: 'قابل ویرایش توسط مدیر سیستم' }));

    /* -- period & data ------------------------------------------------- */
    var periodInput = el('input', { type: 'text', class: 'editable', style: 'width:100%' });
    periodInput.value = App.state.period;
    periodInput.addEventListener('change', function () {
      var old = App.state.period;
      App.state.period = periodInput.value;
      auditConfig('period', old, App.state.period);
      var chip = document.getElementById('periodChip');
      if (chip) chip.textContent = App.state.period;
      save();
    });

    main.appendChild(U.card('دوره و داده‌ها', el('div', {}, [
      el('label', { class: 'field' }, [el('span', { text: 'عنوان دوره' }), periodInput]),
      el('div', { class: 'small muted mb',
        text: 'محل ذخیره‌سازی: ' + (Store.backend() === 'indexeddb' ? 'IndexedDB مرورگر' : 'localStorage مرورگر') +
              ' • هیچ داده‌ای به بیرون ارسال نمی‌شود.' }),
      el('div', { style: 'display:flex;gap:8px;flex-wrap:wrap' }, [
        window.SAMPLE_DATA ? btn('بارگذاری داده نمونه (۱۰۰ نفر از فایل مرجع)',
          function () { loadSampleData(); }) : null,
        btn('پشتیبان‌گیری (JSON)', function () { exportBackup(); }),
        btn('بازیابی از پشتیبان', function () { importBackup(); }),
        btn('پاک کردن کل داده‌ها', function () { wipe(); }, 'danger')
      ].filter(Boolean))
    ])));
  };

  function setConfig(key, value) {
    var old = App.state.config[key];
    if (old === value) return;
    App.state.config[key] = value;
    auditConfig(key, old, value);
    save();
    recalc();
    U.toast('تنظیمات اعمال و محاسبات بازمحاسبه شد.', 'ok', 2200);
  }

  function auditConfig(field, oldValue, newValue) {
    Store.audit(App.state, {
      entity: 'config', field: field, oldValue: oldValue, newValue: newValue,
      reason: 'تغییر تنظیمات سیستم'
    });
  }

  function wipe() {
    U.confirm('تمام اطلاعات پرسنل، پرسشنامه‌ها، تغییرات و گزارش تغییرات حذف شود؟ این عمل بازگشت‌پذیر نیست.',
      { danger: true, confirmLabel: 'حذف کامل' }).then(function (ok) {
      if (!ok) return;
      Store.clear().then(function () {
        App.state = freshState();
        App._keySeq = 0;
        invalidateEmployeeIndex();
        save().then(function () { recalc(); go('dashboard'); U.toast('همه داده‌ها پاک شد.', 'ok'); });
      });
    });
  }

  function loadSampleData() {
    if (!window.SAMPLE_DATA) { U.toast('داده نمونه در این نسخه موجود نیست.', 'warn'); return; }
    U.confirm('داده نمونه (۱۰۰ نفر از فایل مرجع Merit) بارگذاری شود؟ داده‌های فعلی جایگزین می‌شوند.')
      .then(function (ok) {
        if (!ok) return;
        var s = window.SAMPLE_DATA, now = new Date().toISOString();
        App.state = freshState();
        Object.keys(s.config).forEach(function (k) {
          if (App.state.config[k] !== undefined) App.state.config[k] = s.config[k];
        });
        App.state.questionnaires = s.employees.map(function (e, i) {
          var c = JSON.parse(JSON.stringify(e));
          c._key = 'q' + (i + 1);
          c.importedAt = now;
          c.hodComment = c.hodAdjustment !== null && c.hodAdjustment !== undefined
            ? 'مقدار وارد شده از فایل مرجع Merit.xlsb' : '';
          return c;
        });
        App.state.employees = s.employees.map(function (e) {
          return {
            employeeId: e.employeeId, fullName: e.fullName, division: e.division,
            positionTitle: e.positionTitle, jobLevel: e.jobLevel,
            employeeStatus: 'Active', sourceFile: 'Merit.xlsb (نمونه)', importedAt: now
          };
        });
        App.state.importBatches = [{
          batchId: 'sample', fileName: 'Merit.xlsb (داده نمونه)', kind: 'questionnaire',
          sheetName: 'پرسشنامه کارانه تیمی', headerRow: 6, headers: [], mapping: {},
          mappedCount: 13, unmapped: [], warnings: [], accepted: s.employees.length, importedAt: now
        }];
        App._keySeq = s.employees.length;
        Store.audit(App.state, {
          entity: 'process', field: 'sample data', oldValue: '',
          newValue: s.employees.length + ' رکورد',
          reason: 'بارگذاری داده نمونه از فایل مرجع'
        });
        invalidateEmployeeIndex();
        save().then(function () { recalc(); go('dashboard'); U.toast('داده نمونه بارگذاری شد.', 'ok'); });
      });
  }

  /* ======================================================================
   * VIEW — Reports & export
   * ====================================================================*/
  VIEWS.reports = function (main) {
    var t = App.result.totals;
    main.appendChild(head('گزارش و خروجی',
      'خروجی Excel شامل تمام شیت‌های موردنیاز است و قالب‌بندی فایل مرجع را حفظ می‌کند.'));

    if (App.state.finalizedAt) {
      main.appendChild(U.alert('ok', 'دوره نهایی شده است',
        'زمان نهایی‌سازی: ' + U.dateTime(App.state.finalizedAt)));
    } else if (issueCount('err')) {
      main.appendChild(U.alert('warn', 'دوره هنوز نهایی نشده است',
        issueCount('err') + ' خطای باز در مرکز اعتبارسنجی وجود دارد.',
        btn('بررسی', function () { go('validation'); }, 'sm')));
    }

    var strip = el('div', { class: 'kpi-grid' });
    strip.appendChild(U.kpi('پرسنل واجد شرایط', U.int(t.eligibleCount), { kind: 'ok' }));
    strip.appendChild(U.kpi('مجموع پرداخت', U.money(t.sumFinalKaraneh), { kind: 'brand', sub: 'ریال' }));
    strip.appendChild(U.kpi('میانگین پرداخت',
      U.money(t.eligibleCount ? t.sumFinalKaraneh / t.eligibleCount : 0),
      { sub: 'به ازای هر نفر واجد شرایط' }));
    strip.appendChild(U.kpi('باقیمانده بودجه', U.money(t.remainingBudget),
      { kind: Math.abs(t.remainingBudget) < 1 ? 'ok' : 'warn' }));
    main.appendChild(strip);

    var sheets = [
      ['Employee Master', 'اطلاعات پایه پرسنل', App.state.employees.length],
      ['Consolidated Questionnaire', 'مجموعه یکپارچه پرسشنامه‌ها', App.state.questionnaires.length],
      ['روش پرداخت کارانه', 'جدول محاسبه و پرداخت', App.result.rows.length],
      ['Validation Report', 'گزارش اعتبارسنجی', (App.validation || []).length],
      ['Summary', 'خلاصه مدیریتی و به تفکیک واحد', 1],
      ['Configuration', 'پارامترها و جداول مرجع', 1],
      ['Audit Log', 'ردیابی تغییرات', App.state.auditLog.length]
    ];
    var list = el('div', {});
    sheets.forEach(function (s) {
      list.appendChild(el('div', {
        style: 'display:flex;gap:10px;align-items:center;padding:7px 0;border-bottom:1px solid var(--border)'
      }, [
        el('span', { class: 'chip brand mono', text: s[0] }),
        el('span', { text: s[1] }),
        el('span', { class: 'muted small', style: 'margin-inline-start:auto',
          text: U.int(s[2]) + ' ردیف' })
      ]));
    });
    list.appendChild(el('div', { style: 'margin-top:14px;display:flex;gap:8px;flex-wrap:wrap' }, [
      btn('📊 خروجی کامل Excel', function () { exportWorkbook(); }, 'primary'),
      btn('📄 خروجی CSV — روش پرداخت کارانه', function () { exportCsv(); }),
      btn('🖨 چاپ گزارش', function () { window.print(); })
    ]));
    main.appendChild(U.card('محتوای فایل خروجی', list));

    main.appendChild(U.card('راهنمای فرآیند',
      el('ol', { style: 'margin:0;padding-inline-start:20px;line-height:2' }, [
        el('li', { text: 'ورود اطلاعات پرسنل (Employee Master)' }),
        el('li', { text: 'ورود فایل‌های پرسشنامه تیم‌ها — چند فایل به‌صورت همزمان' }),
        el('li', { text: 'تعیین تکلیف رکوردهای تکراری' }),
        el('li', { text: 'بررسی مرکز اعتبارسنجی و رفع خطاها' }),
        el('li', { text: 'بازبینی جدول روش پرداخت کارانه' }),
        el('li', { text: 'اعمال تغییرات معاون بخش همراه با توضیح' }),
        el('li', { text: 'کنترل بودجه (وضعیت باید «متوازن» باشد)' }),
        el('li', { text: 'نهایی‌سازی' }),
        el('li', { text: 'خروجی Excel' })
      ])));
  };

  /* ======================================================================
   * Excel export
   * ====================================================================*/
  /**
   * Build one worksheet, carrying the presentation intent alongside it.
   *
   * Column widths and the autofilter go in as SheetJS understands them.
   * Frozen panes, the header band and number formats are recorded on
   * `!postprocess` and applied to the written file by xlsx-postprocess.js,
   * because the community build of SheetJS emits none of the three.
   */
  function sheetFromAoa(aoa, opts) {
    opts = opts || {};
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    if (opts.cols) ws['!cols'] = opts.cols;
    var range = XLSX.utils.decode_range(ws['!ref']);
    var hr = opts.headerRow || 0;
    ws['!autofilter'] = { ref: XLSX.utils.encode_range(
      { s: { r: hr, c: 0 }, e: { r: range.e.r, c: range.e.c } }) };

    /* Number formats mirror the presentation of the reference workbook:
       thousand separators on rial amounts, two decimals on scores. */
    var numberFormats = {};
    (opts.moneyCols || []).forEach(function (c) { numberFormats[colLetter(c)] = '#,##0'; });
    (opts.intCols   || []).forEach(function (c) { numberFormats[colLetter(c)] = '#,##0'; });
    (opts.scoreCols || []).forEach(function (c) { numberFormats[colLetter(c)] = '0.00'; });

    ws['!postprocess'] = {
      xSplit: (opts.freeze && opts.freeze.xSplit) || 0,
      ySplit: (opts.freeze && opts.freeze.ySplit) || 0,
      headerRow: hr + 1,                       // 1-based, as the XML numbers rows
      numberFormats: numberFormats
    };
    return ws;
  }

  function widths(list) { return list.map(function (w) { return { wch: w }; }); }

  function buildWorkbook() {
    var wb = XLSX.utils.book_new();
    var cfg = App.state.config;
    var t = App.result.totals;

    /* ---- Employee Master ---- */
    var empAoa = [['شماره پرسنلی', 'وضعیت', 'نام', 'نام خانوادگی', 'نام و نام خانوادگی',
      'تاریخ استخدام', 'تاریخ خروج', 'عنوان شغلی', 'نوع همکاری', 'نوع استخدام', 'سطح شغلی',
      'واحد سازمانی', 'دپارتمان', 'روز کارکرد', 'وضعیت دوره آزمایشی',
      'مدیر مستقیم', 'مدیر سطح 1', 'مدیر سطح 2', 'مدیر سطح 3', 'دارای پرسشنامه', 'فایل منبع']];
    App.state.employees.forEach(function (e) {
      var has = App.state.questionnaires.some(function (q) {
        return q.employeeId === e.employeeId && !q.excluded;
      });
      empAoa.push([e.employeeId, e.employeeStatus || '', e.firstName || '', e.lastName || '',
        e.fullName || ((e.firstName || '') + ' ' + (e.lastName || '')).trim(),
        e.dateOfEmployment || '', e.dateOfLeaving || '', e.positionTitle || '',
        e.assignmentType || '', e.employmentType || '', e.jobLevel || '',
        e.division || '', e.department || '',
        (e.workingDays === null || e.workingDays === undefined) ? '' : e.workingDays,
        e.probationStatus || '', e.directManager || '', e.managerLevel1 || '',
        e.managerLevel2 || '', e.managerLevel3 || '', has ? 'بله' : 'خیر', e.sourceFile || '']);
    });
    XLSX.utils.book_append_sheet(wb, sheetFromAoa(empAoa, {
      cols: widths([13, 11, 13, 15, 20, 13, 13, 24, 12, 14, 8, 14, 14, 10, 14, 18, 18, 18, 18, 11, 20]),
      freeze: { xSplit: 1, ySplit: 1 }, intCols: [13]
    }), 'Employee Master');

    /* ---- Consolidated Questionnaire ---- */
    var qAoa = [['شماره پرسنلی', 'نام و نام خانوادگی', 'واحد سازمانی', 'عنوان شغلی', 'JL',
      'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'امتیاز عملکرد', 'عدد کارانه',
      'اثرگذاری ویژه', 'امتیاز اثرگذاری ویژه', 'ضریب کارانه', 'ضریب نهایی کارانه',
      'وضعیت', 'فایل منبع', 'کنار گذاشته شده']];
    App.result.rows.forEach(function (r) {
      qAoa.push([r.employeeId, r.fullName, r.division, r.positionTitle, r.jobLevel,
        r.q1 || '', r.q2 || '', r.q3 || '', r.q4 || '', r.q5 || '',
        r.performanceScore === null ? '' : r.performanceScore,
        r.performanceKaraneh === null ? '' : r.performanceKaraneh,
        r.specialProject ? 'بله' : 'خیر', r.specialImpactValue,
        r.rawCoefficient === null ? '' : r.rawCoefficient,
        r.finalCoefficient === null ? '' : r.finalCoefficient,
        r.status, r._input.sourceFile || '', r.excluded ? 'بله' : 'خیر']);
    });
    XLSX.utils.book_append_sheet(wb, sheetFromAoa(qAoa, {
      cols: widths([13, 20, 15, 22, 6, 11, 11, 11, 11, 11, 12, 11, 12, 14, 12, 14, 13, 20, 12]),
      freeze: { xSplit: 2, ySplit: 1 }, scoreCols: [10, 11, 14, 15], intCols: [13]
    }), 'Consolidated Questionnaire');

    /* ---- روش پرداخت کارانه — column order preserved from the reference ---- */
    var payAoa = [['شماره پرسنلی', 'نام و نام خانوادگی', 'واحد سازمانی', 'عنوان شغلی', 'JL',
      'سطح شغلی', 'عدد گرید', 'Helper', 'عدد ارزیابی عملکرد', 'Helper2', 'امتیاز عملکردی',
      'امتیاز کل', 'دریافتی قبل از تغییرات معاون بخش', 'تغییرات معاون بخش', 'Diff',
      'دریافتی', 'HOD Comment', 'Final Karaneh', 'وضعیت']];
    App.result.rows.forEach(function (r) {
      payAoa.push([r.employeeId, r.fullName, r.division, r.positionTitle, r.jobLevel,
        r.gradeScore === null ? '' : r.gradeScore, r.gradeImpact,
        r.evalScore || 0, r.eligibleEvalScore || 0, r.baseCoefficient || 0,
        r.performanceContribution, r.totalScore, r.initialAllocation,
        r.hodAdjustment === null ? '' : r.hodAdjustment,
        r.diff === null || r.diff === undefined ? '' : r.diff,
        r.finalKaraneh, r.hodComment || '', r.finalKaraneh, r.status]);
    });
    payAoa.push([]);
    payAoa.push(['جمع', '', '', '', '', '', '', '', '', '', '',
      t.sumTotalScore, t.sumInitialAllocation, '', '', t.sumFinalKaraneh, '', t.sumFinalKaraneh, '']);
    XLSX.utils.book_append_sheet(wb, sheetFromAoa(payAoa, {
      cols: widths([13, 20, 15, 22, 6, 10, 10, 9, 15, 10, 14, 12, 24, 20, 18, 20, 28, 20, 13]),
      freeze: { xSplit: 2, ySplit: 1 },
      scoreCols: [5, 6, 7, 8, 9, 10, 11], moneyCols: [12, 13, 14, 15, 17]
    }), 'روش پرداخت کارانه');

    /* ---- Validation Report ---- */
    var vAoa = [['سطح', 'کد', 'عنوان', 'شماره پرسنلی', 'نام', 'توضیح']];
    (App.validation || []).forEach(function (i) {
      vAoa.push([
        i.severity === 'err' ? 'خطا' : i.severity === 'warn' ? 'هشدار' : 'اطلاع',
        i.code, i.title, i.employeeId, i.employeeName, i.detail]);
    });
    XLSX.utils.book_append_sheet(wb, sheetFromAoa(vAoa, {
      cols: widths([9, 24, 30, 13, 20, 60]), freeze: { ySplit: 1 }
    }), 'Validation Report');

    /* ---- Summary ---- */
    var sAoa = [
      ['خلاصه مدیریتی — ' + App.state.period], [],
      ['شاخص', 'مقدار'],
      ['کل پرسنل', App.state.employees.length],
      ['رکورد پرسشنامه', App.state.questionnaires.length],
      ['در دامنه محاسبه', t.inScopeCount],
      ['واجد شرایط', t.eligibleCount],
      ['زیر حد نصاب', t.ineligibleCount],
      ['کنار گذاشته شده (تکراری/ناقص)', t.excludedCount],
      ['تغییرات معاون بخش', t.overriddenCount],
      ['خطای باز', issueCount('err')],
      ['هشدار', issueCount('warn')],
      [],
      ['بودجه (ریال)', t.budget],
      ['تخصیص‌یافته (ریال)', t.allocatedBudget],
      ['باقیمانده (ریال)', t.remainingBudget],
      ['وضعیت بودجه', t.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر'],
      ['مجموع امتیاز کل', t.sumTotalScore],
      ['متناسب‌سازی ضریب به ازای هر نفر',
        (t.sumRawCoefficient - t.inScopeCount * cfg.baselineCoefficientPerPerson) / (t.inScopeCount || 1)],
      ['سرشکن زیر حد نصاب', t.ineligibleRedistribution],
      ['سرشکن تغییرات معاون بخش (ریال)', t.hodRedistribution],
      ['زمان نهایی‌سازی', App.state.finalizedAt ? U.dateTime(App.state.finalizedAt) : 'نهایی نشده'],
      ['زمان تهیه گزارش', U.dateTime(new Date().toISOString())],
      [], ['به تفکیک واحد سازمانی'],
      ['واحد سازمانی', 'تعداد', 'واجد شرایط', 'امتیاز کل', 'کارانه نهایی (ریال)', 'سهم از بودجه']
    ];
    var byDiv = {};
    App.result.rows.forEach(function (r) {
      if (!r.inScope) return;
      var d = r.division || '—';
      var g = byDiv[d] || (byDiv[d] = { count: 0, eligible: 0, score: 0, amount: 0 });
      g.count++; if (r.eligible) g.eligible++;
      g.score += r.totalScore; g.amount += r.finalKaraneh;
    });
    Object.keys(byDiv).sort().forEach(function (d) {
      var g = byDiv[d];
      sAoa.push([d, g.count, g.eligible, g.score, g.amount, t.budget ? g.amount / t.budget : 0]);
    });
    XLSX.utils.book_append_sheet(wb, sheetFromAoa(sAoa, { cols: widths([38, 22, 14, 14, 24, 14]) }), 'Summary');

    /* ---- Configuration ---- */
    var cAoa = [
      ['پارامترهای محاسباتی'], ['پارامتر', 'مقدار', 'مرجع در فایل Excel'],
      ['بودجه کل (ریال)', cfg.budget, 'روش پرداخت کارانه!C1'],
      ['ضریب تأثیر گرید', cfg.gradeImpactFactor, 'روش پرداخت کارانه!D2'],
      ['حداقل امتیاز جهت دریافت کارانه', cfg.minPerformanceThreshold, 'روش پرداخت کارانه!D4'],
      ['نحوه اعمال حد نصاب', cfg.thresholdMode === 'gt' ? 'بزرگ‌تر از' : 'بزرگ‌تر یا مساوی', 'ستون I'],
      ['حداکثر امتیاز کارانه', cfg.maxPerformanceScore, 'پرسشنامه — ستون L'],
      ['تعداد سؤالات عملکردی', cfg.questionCount, 'پرسشنامه — ستون K'],
      ['امتیاز اثرگذاری ویژه', cfg.specialImpactAmount, 'پرسشنامه — ستون N'],
      ['ضریب پایه هر نفر', cfg.baselineCoefficientPerPerson, 'پرسشنامه!B5 ÷ A5'],
      ['دامنه متناسب‌سازی', cfg.normalizationScope === 'global' ? 'کل سازمان' : 'به تفکیک واحد', 'ستون R'],
      ['متناسب‌سازی ضرایب فعال', cfg.normalizeCoefficients ? 'بله' : 'خیر', 'ستون R'],
      ['سرشکن زیر حد نصاب فعال', cfg.redistributeIneligible ? 'بله' : 'خیر', 'سلول J5'],
      ['سرشکن تغییرات معاون فعال', cfg.redistributeHodDiff ? 'بله' : 'خیر', 'سلول O5'],
      [], ['جدول گرید'], ['JL', 'عدد گرید', 'مرجع: Data!C:D']
    ];
    Object.keys(cfg.gradeMap).forEach(function (k) { cAoa.push([k, cfg.gradeMap[k], '']); });
    cAoa.push([], ['نگاشت پاسخ به امتیاز'], ['پاسخ', 'امتیاز', 'مرجع: Data!H:I']);
    Object.keys(cfg.answerScale).forEach(function (k) { cAoa.push([k, cfg.answerScale[k], '']); });
    XLSX.utils.book_append_sheet(wb, sheetFromAoa(cAoa, { cols: widths([34, 22, 30]) }), 'Configuration');

    /* ---- Audit Log ---- */
    var aAoa = [['زمان', 'کاربر', 'بخش', 'شماره پرسنلی', 'نام', 'فیلد', 'مقدار قبلی', 'مقدار جدید', 'دلیل']];
    App.state.auditLog.forEach(function (a) {
      aAoa.push([U.dateTime(a.timestamp), a.user, a.entity, a.employeeId, a.employeeName,
        a.field, a.oldValue, a.newValue, a.reason]);
    });
    XLSX.utils.book_append_sheet(wb, sheetFromAoa(aAoa, {
      cols: widths([18, 14, 14, 13, 20, 22, 22, 22, 42]), freeze: { ySplit: 1 }
    }), 'Audit Log');

    /* Right-to-left sheet orientation, matching the reference workbook. */
    wb.SheetNames.forEach(function (n) { wb.Sheets[n]['!rtl'] = true; });
    wb.Workbook = { Views: [{ RTL: true }] };
    return wb;
  }

  /**
   * Write a workbook to the user's disk.
   *
   * SheetJS's community build has no freeze-pane support, so the archive is
   * written STORED (uncompressed) and the <pane> elements are injected
   * afterwards by xlsx-postprocess.js. The size penalty is the cost of
   * frozen headers on a sheet that can run to thousands of rows.
   */
  function writeWorkbook(wb, filename) {
    var spec = {};
    wb.SheetNames.forEach(function (n) {
      if (wb.Sheets[n]['!postprocess']) spec[n] = wb.Sheets[n]['!postprocess'];
    });
    var raw = XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: false });
    var bytes = window.XlsxPostprocess
      ? window.XlsxPostprocess.applyFormatting(raw, spec)
      : new Uint8Array(raw);
    download(bytes, filename,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }

  function exportWorkbook() {
    try {
      var name = 'Karaneh-' + App.state.period.replace(/[\s\/\\]+/g, '-') + '-' + stamp() + '.xlsx';
      writeWorkbook(buildWorkbook(), name);
      Store.audit(App.state, {
        entity: 'export', field: 'workbook', oldValue: '', newValue: name,
        reason: 'خروجی کامل Excel'
      });
      save();
      U.toast('فایل خروجی تولید شد: ' + name, 'ok', 5000);
    } catch (e) {
      console.error(e);
      U.toast('تولید خروجی ناموفق بود: ' + (e.message || e), 'err', 6000);
    }
  }

  function exportSheet(which) {
    var keep = { employees: 'Employee Master', questionnaire: 'Consolidated Questionnaire',
                 payment: 'روش پرداخت کارانه', validation: 'Validation Report',
                 audit: 'Audit Log' }[which];
    if (!keep) { exportWorkbook(); return; }
    var wb = buildWorkbook();
    var single = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(single, wb.Sheets[keep], keep);
    single.Workbook = { Views: [{ RTL: true }] };
    single.Sheets[keep]['!postprocess'] = wb.Sheets[keep]['!postprocess'];
    writeWorkbook(single, keep.replace(/[\s\/\\]+/g, '-') + '-' + stamp() + '.xlsx');
    U.toast('خروجی تولید شد.', 'ok');
  }

  function exportCsv() {
    var csv = XLSX.utils.sheet_to_csv(buildWorkbook().Sheets['روش پرداخت کارانه']);
    /* BOM so Excel reads the Persian text as UTF-8 instead of mangling it. */
    download('﻿' + csv, 'roshe-pardakht-karaneh-' + stamp() + '.csv', 'text/csv;charset=utf-8');
  }

  function exportBackup() {
    download(JSON.stringify(App.state, null, 1),
      'karaneh-backup-' + stamp() + '.json', 'application/json');
    U.toast('پشتیبان تولید شد.', 'ok');
  }

  function importBackup() {
    var input = el('input', { type: 'file', accept: '.json', style: 'display:none' });
    input.addEventListener('change', function () {
      var f = input.files[0];
      if (!f) return;
      var fr = new FileReader();
      fr.onload = function () {
        try {
          var data = JSON.parse(fr.result);
          if (!data.schemaVersion) throw new Error('ساختار فایل پشتیبان معتبر نیست.');
          U.confirm('داده‌های فعلی با محتوای این پشتیبان جایگزین شود؟', { danger: true })
            .then(function (ok) {
              if (!ok) return;
              App.state = data;
              if (!App.state.config) App.state.config = defaultConfig();
              if (!App.state.columnMappings) App.state.columnMappings = cloneMappings();
              App._keySeq = App.state.questionnaires.length;
              invalidateEmployeeIndex();
              save().then(function () { recalc(); go('dashboard'); U.toast('بازیابی انجام شد.', 'ok'); });
            });
        } catch (e) {
          U.toast('فایل پشتیبان خوانده نشد: ' + e.message, 'err', 6000);
        }
      };
      fr.readAsText(f);
    });
    document.body.appendChild(input);
    input.click();
    setTimeout(function () { if (input.parentNode) input.parentNode.removeChild(input); }, 1000);
  }

  /** Accepts a string or a Uint8Array; Blob handles both. */
  function download(content, filename, type) {
    var blob = new Blob([content], { type: type });
    var url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      if (a.parentNode) a.parentNode.removeChild(a);
      URL.revokeObjectURL(url);
    }, 500);
  }

  function stamp() {
    var d = new Date();
    function p(n) { return n < 10 ? '0' + n : String(n); }
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '-' + p(d.getHours()) + p(d.getMinutes());
  }

  /* ====================================================================== */
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
}());
