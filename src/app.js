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
  var Tpl    = window.Templates;
  var Chart  = window.Charts;
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
  /* The system is two processes, not one long menu: everything that gets the
     answers in, then everything that turns them into money. The split is what
     lets a division head be handed the second half without the first. */
  var NAV = [
    { phase: 1, label: 'ورود پرسشنامه و پاسخ سؤالات' },
    { id: 'designer',       icon: '🧩', label: 'طراحی پرسشنامه', adminOnly: true },
    { id: 'employees',      icon: '👤', label: 'اطلاعات پرسنل',   adminOnly: true },
    { id: 'import',         icon: '📥', label: 'ورود پرسشنامه‌ها', adminOnly: true },
    { id: 'questionnaires', icon: '📝', label: 'مدیریت پرسشنامه', adminOnly: true },
    { id: 'validation',     icon: '🛡', label: 'مرکز اعتبارسنجی' },
    { phase: 2, label: 'محاسبه کارانه و تغییرات معاون بخش' },
    { id: 'dashboard',      icon: '▦',  label: 'داشبورد' },
    { id: 'payment',        icon: '💰', label: 'روش پرداخت کارانه' },
    { id: 'hod',            icon: '✍️', label: 'تغییرات معاون بخش', needsPhase1: true },
    { id: 'reports',        icon: '📤', label: 'گزارش و خروجی' },
    { group: 'سیستم' },
    { id: 'audit',          icon: '🧾', label: 'ردیابی تغییرات' },
    { id: 'settings',       icon: '⚙️', label: 'تنظیمات', adminOnly: true }
  ];

  var ROLES = {
    admin: { label: 'مدیر سیستم / منابع انسانی', icon: '🛠' },
    hod:   { label: 'معاون بخش', icon: '✍️' }
  };

  function role() { return (App.state && App.state.role) || 'admin'; }
  function isAdmin() { return role() === 'admin'; }

  /** Divisions a division head may act on; empty means the whole organisation. */
  function roleScope() {
    var sc = (App.state && App.state.hodScope) || [];
    return sc.length ? sc : null;
  }

  function inScopeForRole(row) {
    if (isAdmin()) return true;
    var sc = roleScope();
    return !sc || sc.indexOf(row.division) !== -1;
  }

  function canOpen(view) {
    var item = null;
    NAV.forEach(function (n) { if (n.id === view) item = n; });
    if (!item) return true;
    if (item.adminOnly && !isAdmin()) return false;
    if (item.needsPhase1 && !phase1Ready()) return false;
    return true;
  }

  /**
   * Phase 1 is complete when nothing is left that would make a payment number
   * wrong: unresolved duplicates, unusable answers, or an unmapped job level.
   * Payment-stage problems (a missing HOD comment, a negative payout) are not
   * counted here — they belong to phase 2.
   */
  function phase1Ready() {
    if (!App.state || !App.state.questionnaires.length) return false;
    return (App.validation || []).filter(function (i) {
      return i.severity === 'err' && i.stage === 'data';
    }).length === 0;
  }

  function phase1Blockers() {
    return (App.validation || []).filter(function (i) {
      return i.severity === 'err' && i.stage === 'data';
    });
  }

  /* ======================================================================
   * Boot
   * ====================================================================*/
  function boot() {
    Store.read().then(function (saved) {
      App.state = saved || freshState();
      if (!App.state.config) App.state.config = defaultConfig();
      if (!App.state.columnMappings) App.state.columnMappings = cloneMappings();
      if (!App.state.role) App.state.role = 'admin';
      if (!App.state.hodScope) App.state.hodScope = [];
      migrateConfig(App.state.config);
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

  /**
   * Bring a stored config forward. Earlier versions listed scored questions as
   * bare ids with no text or weight; rebuild the richer shape from whatever is
   * there so a saved session keeps working after an upgrade.
   */
  function migrateConfig(cfg) {
    if (!cfg) return;
    if (!cfg.questions || !cfg.questions.length) {
      var ids = cfg.scoredQuestions || ['q1', 'q2', 'q3', 'q4'];
      cfg.questions = Engine.DEFAULT_CONFIG.questions.filter(function (q) {
        return ids.indexOf(q.id) !== -1 || q.scored === false;
      }).map(function (q) { return JSON.parse(JSON.stringify(q)); });
    }
    if (cfg.specialImpactMinScore === undefined) {
      cfg.specialImpactMinScore = Engine.DEFAULT_CONFIG.specialImpactMinScore;
    }
    if (!cfg.specialImpactQuestion) {
      cfg.specialImpactQuestion = Engine.DEFAULT_CONFIG.specialImpactQuestion;
    }
    delete cfg.scoredQuestions;
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
    var questionIds = (App.state.config.questions || []).map(function (x) { return x.id; });
    var records = App.state.questionnaires.map(function (q) {
      var master = employeeById(q.employeeId);
      /* Master data wins for organisational attributes; the questionnaire
         wins for answers. A job level typed into a team file is a fallback
         only, because HR's master file is the system of record. */
      var rec = {
        employeeId:    q.employeeId,
        fullName:      q.fullName || (master && master.fullName) || '',
        division:      (master && master.division) || q.division || '',
        positionTitle: (master && master.positionTitle) || q.positionTitle || '',
        jobLevel:      (master && master.jobLevel) || q.jobLevel || '',
        specialProject:      q.specialProject,
        specialImpactAmount: q.specialImpactAmount,
        hodAdjustment:       q.hodAdjustment,
        hodComment:          q.hodComment,
        excluded:            !!q.excluded,
        sourceFile:          q.sourceFile,
        _key:                q._key
      };
      /* Answers travel by the ids the designer defines, so adding a question
         needs no change here. */
      questionIds.forEach(function (id) { rec[id] = q[id]; });
      return rec;
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

    /* `stage` decides which half of the process an issue belongs to: 'data'
       problems must be cleared before the calculation phase opens at all,
       'payment' problems are raised during it. */
    var STAGE = {
      NO_QUESTIONNAIRE: 'data', NO_MASTER: 'data', DUPLICATE: 'data',
      INCOMPLETE_ANSWERS: 'data', UNMAPPED_JL: 'data', MISSING_JL: 'data',
      BELOW_THRESHOLD: 'data', SPECIAL_BLOCKED: 'data', SPECIAL_NO_AMOUNT: 'data',
      NO_MASTER_DATA: 'data', NO_QUESTIONNAIRES: 'data', TEMPLATE_DRIFT: 'data',
      MISSING_HOD_COMMENT: 'payment', NEGATIVE_OVERRIDE: 'payment',
      NEGATIVE_PAYOUT: 'payment', BUDGET_OVERRUN: 'payment'
    };

    function add(severity, code, title, employeeId, employeeName, detail) {
      issues.push({
        severity: severity, code: code, title: title, stage: STAGE[code] || 'data',
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
      if (r.specialImpactBlocked) {
        add('info', 'SPECIAL_BLOCKED', 'اثرگذاری ویژه ثبت شده اما امتیاز نگرفته',
            r.employeeId, r.fullName,
            'امتیاز کارانه ' + U.score(r.performanceKaraneh, 2) + ' کمتر از حد نصاب ' +
            App.state.config.specialImpactMinScore + ' است.');
      } else if (r.specialProject && !(r.specialImpactValue > 0)) {
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
    Engine.scoredQuestions(App.state.config).forEach(function (q) {
      var v = r[q.id];
      var code = q.id.toUpperCase();
      if (v === null || v === undefined || v === '') missing.push(code + ' خالی');
      else if (typeof v !== 'number' && scale[String(v).trim()] === undefined) {
        missing.push(code + ' نامعتبر («' + v + '»)');
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
      el('button', {
        class: 'role-pill', id: 'rolePill',
        title: 'تغییر سطح دسترسی',
        onclick: function () { openRolePicker(); }
      }, [document.createTextNode(ROLES[role()].icon + '  ' + ROLES[role()].label)]),
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
      if (n.phase) {
        sidebar.appendChild(el('div', { class: 'navphase' }, [
          el('span', { class: 'n', text: String(n.phase) }),
          el('span', { text: n.label })
        ]));
        return;
      }
      if (n.group) { sidebar.appendChild(el('div', { class: 'navgroup', text: n.group })); return; }
      if (n.adminOnly && !isAdmin()) return;
      var locked = n.needsPhase1 && !phase1Ready();
      var badge = null;
      if (n.id === 'validation') {
        var errs = issueCount('err'), warns = issueCount('warn');
        if (errs) badge = el('span', { class: 'badge err', text: String(errs) });
        else if (warns) badge = el('span', { class: 'badge warn', text: String(warns) });
      } else if (locked) {
        badge = el('span', { class: 'badge warn', text: '🔒' });
      } else if (n.id === 'questionnaires') {
        badge = el('span', { class: 'badge', text: String(App.state.questionnaires.length) });
      } else if (n.id === 'employees') {
        badge = el('span', { class: 'badge', text: String(App.state.employees.length) });
      } else if (n.id === 'hod') {
        var ov = App.result ? App.result.totals.overriddenCount : 0;
        if (ov) badge = el('span', { class: 'badge', text: String(ov) });
      }
      sidebar.appendChild(el('button', {
        class: 'navitem' + (App.view === n.id ? ' active' : '') + (locked ? ' locked' : ''),
        title: locked ? 'تا رفع خطاهای مرحلهٔ ۱ در دسترس نیست' : '',
        onclick: function () {
          if (locked) {
            U.toast('ابتدا باید خطاهای مرحلهٔ ۱ در مرکز اعتبارسنجی برطرف شوند.', 'warn', 5000);
            go('validation');
            return;
          }
          go(n.id);
        }
      }, [
        el('span', { class: 'ico', text: n.icon }),
        el('span', { text: n.label }),
        badge
      ]));
    });
    var pill = document.getElementById('rolePill');
    if (pill) pill.textContent = ROLES[role()].icon + '  ' + ROLES[role()].label;
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
    if (!canOpen(view)) view = isAdmin() ? 'validation' : 'dashboard';
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
    var banner = phaseBanner(App.view);
    if (banner) main.appendChild(banner);
    var fn = VIEWS[App.view] || VIEWS.dashboard;
    fn(main);
  }

  /**
   * Switching role changes what the sidebar offers and, for a division head,
   * narrows every table to their own people. It is a workflow control, not a
   * security boundary — all the data still lives in this browser.
   */
  function openRolePicker() {
    var divisions = {};
    App.state.employees.forEach(function (e) { if (e.division) divisions[e.division] = 1; });
    App.state.questionnaires.forEach(function (q) { if (q.division) divisions[q.division] = 1; });
    var list = Object.keys(divisions).sort(function (a, b) { return a.localeCompare(b, 'fa'); });

    var chosenRole = role();
    var scope = (App.state.hodScope || []).slice();

    var scopeBox = el('div', { style: 'margin-top:10px' });
    function renderScope() {
      U.clear(scopeBox);
      if (chosenRole !== 'hod') return;
      scopeBox.appendChild(el('div', { class: 'small muted mb',
        text: 'واحدهایی که این معاون بخش مسئول آن‌هاست. اگر هیچ‌کدام انتخاب نشود، کل سازمان در دسترس خواهد بود.' }));
      if (!list.length) {
        scopeBox.appendChild(el('div', { class: 'small muted', text: 'هنوز واحدی در داده‌ها وجود ندارد.' }));
        return;
      }
      list.forEach(function (d) {
        var cb = el('input', { type: 'checkbox' });
        cb.checked = scope.indexOf(d) !== -1;
        cb.addEventListener('change', function () {
          if (cb.checked) { if (scope.indexOf(d) === -1) scope.push(d); }
          else scope = scope.filter(function (x) { return x !== d; });
        });
        scopeBox.appendChild(el('label', { class: 'checkline' }, [cb, el('span', { text: d })]));
      });
    }

    var body = el('div', {});
    Object.keys(ROLES).forEach(function (key) {
      var radio = el('input', { type: 'radio', name: 'role' });
      radio.checked = key === chosenRole;
      radio.addEventListener('change', function () { chosenRole = key; renderScope(); });
      body.appendChild(el('label', {
        class: 'checkline',
        style: 'border:1px solid var(--border);border-radius:7px;padding:9px 11px;margin-bottom:8px'
      }, [
        radio,
        el('div', {}, [
          el('b', { text: ROLES[key].icon + '  ' + ROLES[key].label }),
          el('div', { class: 'small muted', text: key === 'admin'
            ? 'دسترسی کامل: طراحی پرسشنامه، ورود اطلاعات، محاسبات و تنظیمات.'
            : 'فقط مرحلهٔ ۲: مشاهدهٔ محاسبات و تعیین مبلغ برای پرسنل واحدهای خود.' })
        ])
      ]));
    });
    body.appendChild(scopeBox);
    renderScope();

    U.modal({
      title: 'سطح دسترسی', size: 'narrow', content: body,
      buttons: [
        { label: 'اعمال', kind: 'primary', onClick: function () {
          var oldRole = App.state.role;
          App.state.role = chosenRole;
          App.state.hodScope = chosenRole === 'hod' ? scope : [];
          Store.audit(App.state, {
            entity: 'access', field: 'role', oldValue: oldRole, newValue: chosenRole,
            reason: chosenRole === 'hod' && scope.length
              ? 'محدود به واحدهای: ' + scope.join('، ') : 'تغییر سطح دسترسی'
          });
          save().then(function () {
            renderShell();
            go(canOpen(App.view) ? App.view : (chosenRole === 'hod' ? 'dashboard' : 'validation'));
            U.toast('سطح دسترسی به «' + ROLES[chosenRole].label + '» تغییر کرد.', 'ok');
          });
        } },
        { label: 'انصراف' }
      ]
    });
  }

  /** Banner naming the phase the current view belongs to. */
  function phaseBanner(view) {
    var current = null, phase = null;
    NAV.forEach(function (n) {
      if (n.phase) phase = n;
      if (n.id === view) current = phase;
    });
    if (!current) return null;
    var blockers = phase1Blockers().length;
    return el('div', { class: 'phase-banner' }, [
      el('span', { class: 'num', text: String(current.phase) }),
      el('span', {}, [
        el('b', { text: 'مرحلهٔ ' + current.phase + ' — ' + current.label })
      ]),
      el('div', { class: 'spacer' }),
      current.phase === 1
        ? el('span', {
            class: 'chip ' + (phase1Ready() ? 'ok' : 'warn'),
            text: phase1Ready() ? 'آمادهٔ ورود به مرحلهٔ ۲' : blockers + ' مورد باز'
          })
        : el('span', {
            class: 'chip ' + (phase1Ready() ? 'ok' : 'err'),
            text: phase1Ready() ? 'مرحلهٔ ۱ تکمیل شده' : 'مرحلهٔ ۱ هنوز کامل نیست'
          })
    ]);
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
   * VIEW — Questionnaire designer
   * ====================================================================*/
  VIEWS.designer = function (main) {
    var cfg = App.state.config;

    main.appendChild(head('طراحی پرسشنامه',
      'متن سؤالات، وزن آن‌ها، مقیاس پاسخ و شرط سؤال اثرگذاری ویژه از اینجا تعیین می‌شود. ' +
      'تمپلیت Excel و کل محاسبات بر همین اساس ساخته می‌شوند.',
      [
        btn('دانلود تمپلیت پرسشنامه', function () { downloadQuestionnaireTemplate(); }, 'primary'),
        btn('بازنشانی به پرسشنامه مرجع', function () { resetQuestionnaire(); }, 'danger')
      ]));

    if (App.state.questionnaires.length) {
      main.appendChild(U.alert('warn', 'پرسشنامه‌های تکمیل‌شده در سیستم وجود دارد',
        'تغییر مجموعه سؤالات باعث می‌شود فایل‌های تکمیل‌شدهٔ قبلی با طراحی جدید هماهنگ نباشند. ' +
        'تغییر متن یا وزن سؤالات موجود بی‌خطر است؛ افزودن یا حذف سؤال نیازمند توزیع مجدد تمپلیت است.'));
    }

    /* ---- questions ---- */
    var listBox = el('div', {});
    function renderQuestions() {
      U.clear(listBox);
      var scoredCount = Engine.scoredQuestions(cfg).length;
      var totalWeight = 0;
      Engine.scoredQuestions(cfg).forEach(function (q) { totalWeight += Number(q.weight) || 0; });

      cfg.questions.forEach(function (q, idx) {
        var scored = q.scored !== false;
        var row = el('div', { class: 'qrow' + (scored ? '' : ' unscored') });

        var ta = el('textarea', { class: 'editable' });
        ta.value = q.text || '';
        ta.addEventListener('change', function () {
          if (ta.value.trim() === q.text) return;
          auditConfig('question.' + q.id + '.text', q.text, ta.value.trim());
          q.text = ta.value.trim();
          save(); recalc();
        });

        var weight = el('input', { type: 'number', class: 'editable', step: '0.5', min: '0' });
        weight.value = q.weight === undefined ? 1 : q.weight;
        weight.disabled = !scored;
        weight.addEventListener('change', function () {
          var v = Number(weight.value);
          if (!isFinite(v) || v < 0) { weight.value = q.weight; return; }
          auditConfig('question.' + q.id + '.weight', q.weight, v);
          q.weight = v;
          save(); recalc();
        });

        var scoredCb = el('input', { type: 'checkbox' });
        scoredCb.checked = scored;
        scoredCb.addEventListener('change', function () {
          if (!scoredCb.checked && scoredCount <= 1) {
            U.toast('حداقل یک سؤال باید در محاسبه وارد شود.', 'err');
            scoredCb.checked = true; return;
          }
          auditConfig('question.' + q.id + '.scored', scored, scoredCb.checked);
          q.scored = scoredCb.checked;
          if (q.scored && !q.weight) q.weight = 1;
          if (!q.scored) q.weight = 0;
          syncQuestionCount();
          save(); recalc();
        });

        row.appendChild(el('div', { class: 'qhead' }, [
          el('span', { class: 'code', text: q.id.toUpperCase() }),
          scored
            ? el('span', { class: 'chip ok', text: 'در محاسبه' })
            : el('span', { class: 'chip', text: 'فقط اطلاعاتی' }),
          scored && totalWeight
            ? el('span', { class: 'chip info',
                text: 'سهم ' + ((Number(q.weight) || 0) / totalWeight * 100).toFixed(0) + '٪' })
            : null,
          el('div', { class: 'spacer' }),
          el('button', { class: 'btn sm', text: '↑', title: 'انتقال به بالا',
            disabled: idx === 0 ? 'disabled' : null,
            onclick: function () { moveQuestion(idx, -1); } }),
          el('button', { class: 'btn sm', text: '↓', title: 'انتقال به پایین',
            disabled: idx === cfg.questions.length - 1 ? 'disabled' : null,
            onclick: function () { moveQuestion(idx, 1); } }),
          el('button', { class: 'btn sm danger', text: 'حذف',
            onclick: function () { removeQuestion(q); } })
        ]));
        row.appendChild(ta);
        row.appendChild(el('div', { class: 'qmeta' }, [
          el('label', {}, [scoredCb, el('span', { text: 'در محاسبهٔ امتیاز عملکرد وارد شود' })]),
          el('label', {}, [el('span', { text: 'وزن' }), weight])
        ]));
        listBox.appendChild(row);
      });

      listBox.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:4px' }, [
        btn('＋ افزودن سؤال', function () { addQuestion(); }, 'primary')
      ]));

      listBox.appendChild(el('div', { class: 'small muted', style: 'margin-top:9px' }, [
        document.createTextNode('امتیاز عملکرد = میانگین وزنی ' + scoredCount +
          ' سؤال محاسباتی، در بازهٔ ۱ تا ۵. عدد کارانه = امتیاز × ' +
          cfg.maxPerformanceScore + ' ÷ ' + cfg.questionCount + '.')
      ]));
    }
    renderQuestions();
    main.appendChild(U.card('سؤالات عملکرد', listBox,
      { hint: 'ستون‌های F تا J شیت مرجع' }));

    /* ---- special impact question ---- */
    var siText = el('textarea', { class: 'editable', style: 'width:100%;min-height:52px' });
    siText.value = cfg.specialImpactQuestion || '';
    siText.addEventListener('change', function () {
      auditConfig('specialImpactQuestion', cfg.specialImpactQuestion, siText.value.trim());
      cfg.specialImpactQuestion = siText.value.trim();
      save();
    });

    var siMin = el('input', { type: 'number', class: 'editable', step: '5', style: 'width:100%' });
    siMin.value = cfg.specialImpactMinScore;
    siMin.addEventListener('change', function () {
      var v = Number(siMin.value);
      if (!isFinite(v) || v < 0) { siMin.value = cfg.specialImpactMinScore; return; }
      setConfig('specialImpactMinScore', v);
    });

    var siAmount = el('input', { type: 'number', class: 'editable', step: '10', style: 'width:100%' });
    siAmount.value = cfg.specialImpactAmount;
    siAmount.addEventListener('change', function () {
      var v = Number(siAmount.value);
      if (!isFinite(v)) { siAmount.value = cfg.specialImpactAmount; return; }
      setConfig('specialImpactAmount', v);
    });

    /* How many people the current gate would actually let through. */
    var eligibleForSpecial = App.result.rows.filter(function (r) {
      return r.inScope && r.specialImpactUnlocked;
    }).length;
    var blocked = App.result.rows.filter(function (r) { return r.specialImpactBlocked; }).length;

    main.appendChild(U.card('سؤال اثرگذاری ویژه', el('div', {}, [
      el('label', { class: 'field' }, [el('span', { text: 'متن سؤال' }), siText]),
      el('div', { class: 'form-grid' }, [
        el('label', { class: 'field' }, [
          el('span', { html: 'حداقل امتیاز کارانه برای فعال شدن این سؤال ' +
            '<span class="muted small">— پاسخ به آن زیر این عدد امتیازی نمی‌گیرد</span>' }),
          siMin
        ]),
        el('label', { class: 'field' }, [
          el('span', { html: 'امتیاز اثرگذاری ویژه <span class="muted small">— ستون N</span>' }),
          siAmount
        ])
      ]),
      App.result.totals.inScopeCount
        ? U.alert('info', 'اثر حد نصاب فعلی',
            eligibleForSpecial + ' نفر از ' + App.result.totals.inScopeCount +
            ' نفر امتیاز کارانه‌شان به حد نصاب ' + cfg.specialImpactMinScore +
            ' می‌رسد و می‌توانند امتیاز اثرگذاری ویژه بگیرند.' +
            (blocked ? ' هم‌اکنون ' + blocked + ' نفر با وجود ثبت اثرگذاری ویژه، امتیازی دریافت نمی‌کنند.' : ''))
        : null
    ]), { hint: 'ستون‌های M و N شیت مرجع' }));

    /* ---- answer scale ---- */
    var scaleBody = el('div', {});
    function renderScale() {
      U.clear(scaleBody);
      var st = el('table', { class: 'grid' });
      st.appendChild(el('thead', {}, [el('tr', {}, [
        el('th', { text: 'متن پاسخ' }), el('th', { text: 'امتیاز' }), el('th', { text: '' })
      ])]));
      var stb = el('tbody');
      Object.keys(cfg.answerScale).forEach(function (k) {
        var inp = el('input', { type: 'number', class: 'cell', step: '1' });
        inp.value = cfg.answerScale[k];
        inp.addEventListener('change', function () {
          auditConfig('answerScale.' + k, cfg.answerScale[k], Number(inp.value));
          cfg.answerScale[k] = Number(inp.value);
          save(); recalc();
        });
        stb.appendChild(el('tr', {}, [
          el('td', { text: k }), el('td', {}, [inp]),
          el('td', {}, [el('button', { class: 'btn sm danger', text: 'حذف', onclick: function () {
            if (Object.keys(cfg.answerScale).length <= 2) {
              U.toast('حداقل دو گزینه پاسخ لازم است.', 'err'); return;
            }
            auditConfig('answerScale.' + k, cfg.answerScale[k], '(حذف شد)');
            delete cfg.answerScale[k];
            save(); recalc();
          } })])
        ]));
      });
      st.appendChild(stb);
      scaleBody.appendChild(st);

      var preview = el('div', { class: 'scale-preview' });
      Object.keys(cfg.answerScale).forEach(function (k) {
        preview.appendChild(el('span', { html: U.esc(k) + ' <b>' + cfg.answerScale[k] + '</b>' }));
      });
      scaleBody.appendChild(preview);

      var newAns = el('input', { type: 'text', placeholder: 'متن پاسخ', style: 'width:170px' });
      var newVal = el('input', { type: 'number', placeholder: 'امتیاز', style: 'width:95px' });
      scaleBody.appendChild(el('div', { style: 'display:flex;gap:7px;margin-top:11px' }, [
        newAns, newVal,
        btn('افزودن', function () {
          var k = newAns.value.trim(), v = Number(newVal.value);
          if (!k || !isFinite(v)) { U.toast('متن پاسخ و امتیاز را وارد کنید.', 'err'); return; }
          cfg.answerScale[k] = v;
          auditConfig('answerScale.' + k, '(جدید)', v);
          newAns.value = ''; newVal.value = '';
          save(); recalc(); renderScale();
        }, 'primary')
      ]));
    }
    renderScale();
    main.appendChild(U.card('نمودار ارزیابی — نگاشت پاسخ به امتیاز', scaleBody,
      { hint: 'جدول Data!H:I — این گزینه‌ها در تمپلیت Excel به فهرست کشویی تبدیل می‌شوند' }));

    /* ---- template ---- */
    main.appendChild(U.card('تمپلیت Excel', el('div', {}, [
      el('p', { class: 'small muted', style: 'margin-top:0',
        text: 'تمپلیت دقیقاً بر اساس طراحی بالا ساخته می‌شود: هر سؤال یک ستون، ' +
              'گزینه‌های پاسخ به‌صورت فهرست کشویی، و یک امضای پنهان که هنگام بازگشت فایل ' +
              'با طراحی فعلی مقایسه می‌شود.' }),
      el('dl', { class: 'kv', style: 'margin-bottom:12px' }, [
        el('dt', { text: 'امضای طراحی فعلی' }),
        el('dd', { class: 'mono', text: Tpl.signature(cfg) }),
        el('dt', { text: 'سؤالات' }),
        el('dd', { class: 'mono', text: Tpl.describe(cfg).questionIds }),
        el('dt', { text: 'گزینه‌های پاسخ' }),
        el('dd', { text: Tpl.describe(cfg).options.split('|').join('، ') })
      ]),
      el('div', { style: 'display:flex;gap:8px;flex-wrap:wrap' }, [
        btn('📄 تمپلیت خالی', function () { downloadQuestionnaireTemplate({ prefill: false }); }),
        btn('👥 تمپلیت با فهرست پرسنل', function () { downloadQuestionnaireTemplate({ prefill: true }); }, 'primary'),
        btn('🏢 تمپلیت به تفکیک واحد', function () { downloadQuestionnaireTemplatePerDivision(); })
      ])
    ])));
  };

  function syncQuestionCount() {
    var n = Engine.scoredQuestions(App.state.config).length;
    if (n && App.state.config.questionCount !== n) {
      auditConfig('questionCount', App.state.config.questionCount, n);
      App.state.config.questionCount = n;
    }
  }

  function nextQuestionId(cfg) {
    var n = 1;
    while (cfg.questions.some(function (q) { return q.id === 'q' + n; })) n++;
    return 'q' + n;
  }

  function addQuestion() {
    var cfg = App.state.config;
    var id = nextQuestionId(cfg);
    cfg.questions.push({ id: id, text: 'متن سؤال جدید', weight: 1, scored: true });
    auditConfig('question.' + id, '(جدید)', 'افزوده شد');
    syncQuestionCount();
    save().then(function () { recalc(); U.toast('سؤال ' + id.toUpperCase() + ' اضافه شد.', 'ok'); });
  }

  function removeQuestion(q) {
    var cfg = App.state.config;
    if (Engine.scoredQuestions(cfg).length <= 1 && q.scored !== false) {
      U.toast('حداقل یک سؤال محاسباتی باید باقی بماند.', 'err');
      return;
    }
    var answered = App.state.questionnaires.filter(function (r) {
      return r[q.id] !== undefined && r[q.id] !== null && r[q.id] !== '';
    }).length;
    U.confirm('سؤال ' + q.id.toUpperCase() + ' حذف شود؟' +
      (answered ? '\n' + answered + ' پاسخ ثبت‌شده برای این سؤال از محاسبات خارج می‌شود.' : ''),
      { danger: true, confirmLabel: 'حذف' }).then(function (ok) {
      if (!ok) return;
      cfg.questions = cfg.questions.filter(function (x) { return x !== q; });
      auditConfig('question.' + q.id, q.text, '(حذف شد)');
      syncQuestionCount();
      save().then(function () { recalc(); U.toast('سؤال حذف شد.', 'ok'); });
    });
  }

  function moveQuestion(idx, delta) {
    var qs = App.state.config.questions;
    var to = idx + delta;
    if (to < 0 || to >= qs.length) return;
    var tmp = qs[idx]; qs[idx] = qs[to]; qs[to] = tmp;
    auditConfig('questionOrder', idx, to);
    save().then(function () { recalc(); });
  }

  function resetQuestionnaire() {
    U.confirm('طراحی پرسشنامه به حالت فایل مرجع Merit بازگردد؟',
      { danger: true, confirmLabel: 'بازنشانی' }).then(function (ok) {
      if (!ok) return;
      var cfg = App.state.config;
      cfg.questions = JSON.parse(JSON.stringify(Engine.DEFAULT_CONFIG.questions));
      cfg.answerScale = JSON.parse(JSON.stringify(Engine.DEFAULT_CONFIG.answerScale));
      cfg.specialImpactQuestion = Engine.DEFAULT_CONFIG.specialImpactQuestion;
      cfg.specialImpactMinScore = Engine.DEFAULT_CONFIG.specialImpactMinScore;
      cfg.specialImpactAmount = Engine.DEFAULT_CONFIG.specialImpactAmount;
      cfg.questionCount = Engine.DEFAULT_CONFIG.questionCount;
      auditConfig('questionnaire', 'custom', 'reset to Merit reference');
      save().then(function () { recalc(); U.toast('طراحی بازنشانی شد.', 'ok'); });
    });
  }

  /* ======================================================================
   * VIEW — Dashboard
   * ====================================================================*/
  VIEWS.dashboard = function (main) {
    var t = App.result.totals;
    var scoped = dashboardRows();

    main.appendChild(head(
      'داشبورد — ' + App.state.period,
      isAdmin()
        ? 'دید یک‌نگاهی از کل فرآیند. نمودارها، جدول یکپارچه و فیلترها همگی به یک مجموعه داده متصل‌اند.'
        : 'دید واحدهای تحت مسئولیت شما.',
      [
        btn('📊 خروجی کارانه (قالب حقوق و دستمزد)', function () { exportPayrollFile(); }, 'primary'),
        btn('👔 خروجی به تفکیک مدیر', function () { exportByManager(); }),
        btn('گزارش کامل', function () { go('reports'); }, 'ghost')
      ]));

    if (!isAdmin() && roleScope()) {
      main.appendChild(U.alert('info', 'دامنهٔ دسترسی شما',
        'واحدهای ' + roleScope().join('، ') + ' — ' + scoped.length + ' نفر.'));
    }

    if (t.budgetStatus !== 'BALANCED') {
      main.appendChild(U.alert('err', 'بودجه در وضعیت نامعتبر است',
        t.budgetOverrun
          ? 'مجموع کارانه نهایی از بودجه تعیین‌شده بیشتر است. تا رفع این مورد، نهایی‌سازی ممکن نیست.'
          : t.negativePayoutCount + ' نفر با اعمال تغییرات معاون بخش دریافتی منفی پیدا کرده‌اند.',
        btn('بررسی', function () { go('validation'); }, 'sm')));
    }
    if (!phase1Ready()) {
      main.appendChild(U.alert('warn', 'مرحلهٔ ۱ هنوز کامل نشده است',
        phase1Blockers().length + ' مورد باید پیش از اتکا به این اعداد برطرف شود.',
        btn('مرکز اعتبارسنجی', function () { go('validation'); }, 'sm')));
    }

    /* ---- headline figures ----
       Every tile describes the same population the charts and table below
       describe. For a division head that is their own units, not the whole
       organisation — otherwise the numbers would not add up to what they see. */
    var scopeOf = function (e) {
      if (isAdmin()) return true;
      var sc = roleScope();
      return !sc || sc.indexOf(e.division) !== -1;
    };
    var staff = App.state.employees.filter(scopeOf);
    var withQ = 0, missing = 0;
    staff.forEach(function (e) {
      var has = App.state.questionnaires.some(function (q) {
        return q.employeeId === e.employeeId && !q.excluded;
      });
      if (has) withQ++;
      else if (isPayrollEligible(e)) missing++;
    });
    if (!App.state.employees.length) {
      withQ = App.state.questionnaires.filter(function (q) { return !q.excluded; }).length;
    }

    var paid = 0, eligible = 0, ineligible = 0, overridden = 0, exceptions = 0;
    scoped.forEach(function (r) {
      paid += r.finalKaraneh;
      if (r.eligible) eligible++; else ineligible++;
      if (r.isOverridden) overridden++;
      if (r.negative || !r.hasQuestionnaire || r.gradeScore === null) exceptions++;
    });

    var grid = el('div', { class: 'kpi-grid' });
    grid.appendChild(U.kpi(isAdmin() ? 'کل پرسنل' : 'پرسنل واحدهای شما',
      U.int(staff.length || scoped.length), { kind: 'brand' }));
    grid.appendChild(U.kpi('دارای پرسشنامه', U.int(withQ), { kind: 'info' }));
    grid.appendChild(U.kpi('فاقد پرسشنامه', U.int(missing),
      { kind: missing ? 'warn' : '', sub: 'از میان پرسنل فعال' }));
    grid.appendChild(U.kpi('واجد شرایط', U.int(eligible),
      { kind: 'ok', sub: 'امتیاز بالاتر از حد نصاب' }));
    grid.appendChild(U.kpi('زیر حد نصاب', U.int(ineligible), { kind: ineligible ? 'warn' : '' }));
    grid.appendChild(U.kpi('تغییرات معاون بخش', U.int(overridden), { kind: 'info' }));
    grid.appendChild(U.kpi('موارد استثنا', U.int(isAdmin() ? issueCount('err') : exceptions),
      { kind: (isAdmin() ? issueCount('err') : exceptions) ? 'err' : 'ok' }));
    grid.appendChild(U.kpi(isAdmin() ? 'کارانه نهایی' : 'کارانه واحدهای شما',
      U.moneyShort(paid), { kind: 'brand', sub: 'ریال' }));
    main.appendChild(grid);

    /* ---- budget meter ---- */
    var meterHost = el('div', {});
    var budgetBody = el('div', {}, [
      meterHost,
      el('div', { class: 'budget-legend' }, [
        el('span', { html: 'بودجه: <b class="num">' + U.money(t.budget) + '</b> ریال' }),
        el('span', { html: 'تخصیص‌یافته: <b class="num">' + U.money(t.allocatedBudget) + '</b>' }),
        el('span', { html: 'باقیمانده: <b class="num">' + U.money(t.remainingBudget) + '</b>' }),
        el('span', {}, [
          document.createTextNode('وضعیت: '),
          el('span', { class: 'chip ' + (t.budgetStatus === 'BALANCED' ? 'ok' : 'err'),
            text: t.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر' })
        ])
      ])
    ]);
    main.appendChild(U.card('کنترل بودجه', budgetBody, {
      hint: isAdmin()
        ? 'مجموع تخصیص همواره دقیقاً برابر بودجه است'
        : 'بودجه در سطح کل سازمان — سهم واحدهای شما در کاشی بالا آمده است'
    }));

    /* ---- charts ---- */
    var divisionHost = el('div', {});
    var statusHost = el('div', {});
    var statusLegendHost = el('div', {});
    var scoreHost = el('div', {});
    var levelHost = el('div', {});

    var charts = el('div', { class: 'chart-grid-2' }, [
      U.card('کارانه به تفکیک واحد سازمانی', divisionHost,
        { hint: 'مجموع پرداختی هر واحد (ریال)' }),
      U.card('توزیع امتیاز عملکرد', scoreHost,
        { hint: 'تعداد افراد در هر بازهٔ امتیاز' }),
      U.card('ترکیب وضعیت پرسنل', el('div', {}, [statusHost, statusLegendHost]),
        { hint: 'سهم هر وضعیت از جمعیت محاسبه' }),
      U.card('میانگین کارانه به تفکیک سطح شغلی', levelHost,
        { hint: 'میانگین دریافتی هر JL (ریال)' })
    ]);
    main.appendChild(charts);

    /* ---- unified table with filters ---- */
    var tableCard = el('div', {});
    main.appendChild(tableCard);
    main.appendChild(U.card('مسیر فرآیند', workflowNode(), { hint: 'وضعیت هر مرحله' }));

    /* Charts need real widths, so draw after the nodes are in the document. */
    requestAnimationFrame(function () {
      Chart.meter(meterHost, t.allocatedBudget, t.budget);
      drawDashboardCharts(scoped, {
        division: divisionHost, status: statusHost, statusLegend: statusLegendHost,
        score: scoreHost, level: levelHost
      });
      App.grids.unified = unifiedGrid(scoped);
      tableCard.appendChild(App.grids.unified.node);
    });
  };

  /** The population this dashboard is allowed to describe. */
  function dashboardRows() {
    return App.result.rows.filter(function (r) {
      return r.inScope && inScopeForRole(r);
    });
  }

  function drawDashboardCharts(rows, hosts) {
    var cfg = App.state.config;

    /* Magnitude by division — horizontal, because unit names are long. */
    var byDiv = {};
    rows.forEach(function (r) {
      var d = r.division || 'بدون واحد';
      var g = byDiv[d] || (byDiv[d] = { amount: 0, count: 0, eligible: 0 });
      g.amount += r.finalKaraneh; g.count++;
      if (r.eligible) g.eligible++;
    });
    var divRows = Object.keys(byDiv).map(function (d) {
      return {
        label: d, value: byDiv[d].amount,
        detail: U.money(byDiv[d].amount) + ' ریال<br>' + byDiv[d].count + ' نفر · ' +
                byDiv[d].eligible + ' واجد شرایط'
      };
    }).sort(function (a, b) { return b.value - a.value; }).slice(0, 12);
    Chart.horizontalBar(hosts.division, divRows, { format: U.moneyShort });

    /* Part-to-whole across workflow states. Segments are always labelled and
       separated, so the reserved status hues never carry meaning alone. */
    var seg = [
      { label: 'محاسبه شد', color: Chart.STATUS.good, value: 0 },
      { label: 'تغییر معاون بخش', color: Chart.STATUS.neutral, value: 0 },
      { label: 'زیر حد نصاب', color: Chart.STATUS.warning, value: 0 },
      { label: 'ناقص یا استثنا', color: Chart.STATUS.critical, value: 0 }
    ];
    rows.forEach(function (r) {
      if (!r.hasQuestionnaire || r.gradeScore === null || r.negative) seg[3].value++;
      else if (!r.eligible) seg[2].value++;
      else if (r.isOverridden) seg[1].value++;
      else seg[0].value++;
    });
    Chart.stackedBar(hosts.status, seg);
    U.clear(hosts.statusLegend);
    hosts.statusLegend.appendChild(Chart.legend(seg, rows.length));

    /* Distribution of performance scores across half-point bins. */
    var bins = [], binSize = 0.5, maxScore = 5;
    for (var b = 1; b < maxScore + binSize; b += binSize) bins.push({ from: b, count: 0 });
    rows.forEach(function (r) {
      if (r.performanceScore === null) return;
      var idx = Math.min(bins.length - 1, Math.max(0, Math.round((r.performanceScore - 1) / binSize)));
      bins[idx].count++;
    });
    var threshold = Number(cfg.minPerformanceThreshold);
    Chart.columns(hosts.score, bins.map(function (b) {
      var below = b.from <= threshold;
      return {
        label: b.from.toFixed(1),
        value: b.count,
        color: below ? 'var(--status-warning)' : 'var(--chart-series)',
        sub: below ? 'زیر نصاب' : '',
        detail: b.count + ' نفر با امتیاز ' + b.from.toFixed(1) +
                (below ? '<br>زیر حد نصاب — کارانه صفر' : '')
      };
    }), { height: 200 });

    /* Average payout per job level — shows whether grade matters at all. */
    var byLevel = {};
    rows.forEach(function (r) {
      var k = r.jobLevel || '—';
      var g = byLevel[k] || (byLevel[k] = { sum: 0, n: 0 });
      g.sum += r.finalKaraneh; g.n++;
    });
    var levelRows = Object.keys(byLevel).sort(function (a, b) {
      return U.naturalCompare(a, b);
    }).map(function (k) {
      return {
        label: k, value: byLevel[k].n ? byLevel[k].sum / byLevel[k].n : 0,
        detail: byLevel[k].n + ' نفر<br>میانگین ' + U.money(byLevel[k].sum / byLevel[k].n) + ' ریال'
      };
    });
    /* This series is rial, not a head count, so it carries its own formatter. */
    Chart.columns(hosts.level, levelRows, { height: 190, maxBarWidth: 54, format: U.moneyShort });
  }

  /**
   * The one table that answers "who gets what, and why" — every field a
   * manager might filter, sort or export on, in one place. It doubles as the
   * table view that makes the charts above readable without color.
   */
  function unifiedGrid(rows) {
    return U.DataGrid({
      title: 'جدول یکپارچه کارانه',
      rows: rows,
      sortKey: 'finalKaraneh', sortDir: 'desc',
      searchFields: ['employeeId', 'fullName', 'division', 'positionTitle', 'directManager'],
      facets: [
        { key: 'division', label: 'همه واحدها' },
        { key: 'jobLevel', label: 'همه سطوح شغلی' },
        { key: 'status', label: 'همه وضعیت‌ها' },
        { key: 'manager', label: 'همه مدیران', value: function (r) { return managerOf(r, 'directManager'); } },
        { key: 'managerL1', label: 'همه مدیران سطح ۱', value: function (r) { return managerOf(r, 'managerLevel1'); } },
        { key: 'special', label: 'اثرگذاری ویژه',
          value: function (r) { return r.specialImpactValue > 0 ? 'دارد' : (r.specialImpactBlocked ? 'ثبت شده ولی زیر نصاب' : 'ندارد'); } }
      ],
      rowClass: function (r) {
        if (r.negative) return 'row-err';
        if (!r.eligible) return 'row-warn';
        return '';
      },
      actions: [
        { label: '⬇ خروجی همین نما', kind: 'sm', onClick: function () { exportCurrentView(); } }
      ],
      columns: [
        { key: 'employeeId', label: 'شماره پرسنلی', alwaysVisible: true, width: '95px' },
        { key: 'fullName', label: 'نام و نام خانوادگی', width: '155px' },
        { key: 'division', label: 'واحد سازمانی' },
        { key: 'positionTitle', label: 'عنوان شغلی', width: '160px', hidden: true },
        { key: 'jobLevel', label: 'JL', width: '48px' },
        { key: 'manager', label: 'مدیر مستقیم', width: '150px',
          value: function (r) { return managerOf(r, 'directManager'); } },
        { key: 'managerL1', label: 'مدیر سطح ۱', width: '150px', hidden: true,
          value: function (r) { return managerOf(r, 'managerLevel1'); } },
        { key: 'managerL2', label: 'مدیر سطح ۲', width: '150px', hidden: true,
          value: function (r) { return managerOf(r, 'managerLevel2'); } },
        { key: 'performanceScore', label: 'امتیاز عملکرد', type: 'score', calculated: true },
        { key: 'performanceKaraneh', label: 'عدد کارانه', type: 'score', decimals: 2, calculated: true },
        { key: 'specialImpactValue', label: 'اثرگذاری ویژه', type: 'score', decimals: 0, calculated: true,
          render: function (r) {
            if (r.specialImpactBlocked) {
              return el('span', { class: 'chip warn', text: '۰ (زیر نصاب)' });
            }
            return document.createTextNode(r.specialImpactValue ? U.score(r.specialImpactValue, 0) : '—');
          } },
        { key: 'gradeScore', label: 'عدد گرید', type: 'score', decimals: 0, calculated: true, hidden: true },
        { key: 'totalScore', label: 'امتیاز کل', type: 'score', decimals: 2, calculated: true },
        { key: 'initialAllocation', label: 'دریافتی محاسباتی', type: 'money', calculated: true },
        { key: 'hodAdjustment', label: 'تغییر معاون بخش', type: 'money' },
        { key: 'finalKaraneh', label: 'کارانه نهایی (ریال)', type: 'money', calculated: true,
          className: function (r) { return r.negative ? 'neg' : ''; } },
        { key: 'status', label: 'وضعیت', render: function (r) { return statusChip(r); } },
        { key: 'hodComment', label: 'توضیح معاون بخش', width: '180px', hidden: true }
      ],
      footer: function (visible) {
        var alloc = 0, fin = 0, sc = 0;
        visible.forEach(function (r) {
          alloc += r.initialAllocation; fin += r.finalKaraneh; sc += r.totalScore;
        });
        return {
          employeeId: 'جمع (' + visible.length + ')',
          totalScore: U.score(sc, 2),
          initialAllocation: U.money(alloc),
          finalKaraneh: U.money(fin)
        };
      },
      onRowClick: function (r) { showEmployeeDetail(r.employeeId); }
    });
  }

  /** Manager names live on the master record, not the questionnaire. */
  function managerOf(row, field) {
    var m = employeeById(row.employeeId);
    return (m && m[field]) || '';
  }

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
        btn('📄 دانلود تمپلیت حقوق و دستمزد', function () { downloadEmployeeTemplate(); }),
        btn('ورود فایل پرسنل', function () { pickFiles('employee'); }, 'primary'),
        btn('خروجی', function () { exportSheet('employees'); }),
        App.state.employees.length ? btn('پاک کردن', function () { clearEmployees(); }, 'danger') : null
      ].filter(Boolean)));

    main.appendChild(U.alert('info', 'تمپلیت هماهنگ با فایل تیم حقوق و دستمزد',
      'ستون‌های تمپلیت دقیقاً همان ستون‌های فایل «کارانه بهار ۱۴۰۵-Q1» است ' +
      '(Emp No، Emp Status، Job Level، Division Alias، Working Day\'s، Manager Level 1..3 و …)، ' +
      'بنابراین خروجی تیم حقوق و دستمزد بدون تغییر قابل ورود است و خروجی نهایی نیز در همان قالب تولید می‌شود.'));

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
      'ردیف عنوان را تشخیص می‌دهد، ستون‌ها را نگاشت می‌کند، ساختار را با تمپلیت مقایسه می‌کند ' +
      'و همه را در یک مجموعه واحد ادغام می‌کند.',
      [btn('انتخاب فایل‌ها', function () { pickFiles('questionnaire'); }, 'primary')]));

    main.appendChild(U.card('۱. دانلود تمپلیت', el('div', {}, [
      el('p', { class: 'small muted', style: 'margin-top:0',
        text: 'تمپلیت بر اساس طراحی فعلی پرسشنامه ساخته می‌شود: هر سؤال یک ستون، ' +
              'گزینه‌های پاسخ به‌صورت فهرست کشویی، و فهرست پرسنل از پیش پر شده. ' +
              'فایل امضای طراحی را در خود دارد تا هنگام بازگشت، هر ناسازگاری تشخیص داده شود.' }),
      el('div', { style: 'display:flex;gap:8px;flex-wrap:wrap' }, [
        btn('👥 تمپلیت با فهرست پرسنل', function () { downloadQuestionnaireTemplate({ prefill: true }); }, 'primary'),
        btn('🏢 به تفکیک واحد سازمانی', function () { downloadQuestionnaireTemplatePerDivision(); }),
        btn('📄 تمپلیت خالی', function () { downloadQuestionnaireTemplate({ prefill: false }); }),
        btn('🧩 تغییر طراحی پرسشنامه', function () { go('designer'); }, 'ghost')
      ]),
      el('div', { class: 'small muted', style: 'margin-top:9px' }, [
        document.createTextNode('امضای طراحی فعلی: '),
        el('span', { class: 'mono', text: Tpl.signature(App.state.config) })
      ])
    ]), { hint: 'مطابق طراحی فعلی پرسشنامه' }));

    main.appendChild(el('h2', {
      style: 'font-size:14px;margin:18px 0 9px;font-weight:700'
    }, [document.createTextNode('۲. ورود فایل‌های تکمیل‌شده')]));

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
          var res = Import.parseWorkbook(buf, {
            fileName: f.name, kind: kind, mappings: App.state.columnMappings
          });
          /* A questionnaire must match the design it was cut from. A signed
             file is checked against the stored signature; an unsigned one
             falls back to checking that every scored question has a column. */
          if (kind === 'questionnaire') {
            res.template = Tpl.verifyAgainstTemplate(res.workbook, App.state.config, XLSX);
            if (res.template.level === 'unsigned') {
              var col = Tpl.verifyColumns(res.mapping, App.state.config);
              if (!col.ok) {
                res.template = { ok: false, level: 'mismatch', problems: col.problems, meta: null };
              }
            }
            if (!res.template.ok) {
              errors.push({
                file: f.name,
                message: 'ساختار فایل با تمپلیت فعلی هماهنگ نیست.',
                detail: res.template.problems
              });
              return;
            }
          }
          parsed.push(res);
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
          content: el('div', {}, [templateErrorList(errors),
            el('div', { style: 'margin-top:12px' }, [
              btn('دانلود تمپلیت صحیح', function () { downloadQuestionnaireTemplate(); }, 'primary')
            ])]),
          buttons: [{ label: 'بستن' }]
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
    if (errors.length) body.appendChild(templateErrorList(errors));
    parsed.forEach(function (p) {
      if (p.template && p.template.level === 'drift') {
        body.appendChild(U.alert('warn', p.fileName + ' — تمپلیت قدیمی است',
          U.esc(p.template.problems.join(' ')) ));
      } else if (p.template && p.template.level === 'unsigned') {
        body.appendChild(U.alert('info', p.fileName + ' — بدون امضای تمپلیت',
          'این فایل از تمپلیت این سامانه تولید نشده است. ستون‌ها بررسی شدند و ' +
          'تمام سؤالات محاسباتی پیدا شد، اما توصیه می‌شود از تمپلیت رسمی استفاده شود.'));
      }
    });

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

  /** Render import failures with the specific mismatch under each file. */
  function templateErrorList(errors) {
    var wrap = el('div', {});
    errors.forEach(function (e) {
      var detail = el('div', {});
      detail.appendChild(el('div', { text: e.message }));
      (e.detail || []).forEach(function (line) {
        detail.appendChild(el('div', {
          class: 'small mono',
          style: 'margin-top:5px;white-space:pre-wrap;direction:rtl'
        }, [document.createTextNode(line)]));
      });
      wrap.appendChild(U.alert('err', e.file, detail));
    });
    return wrap;
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
      'پاسخ‌ها را می‌توان مستقیماً در همین جدول ثبت کرد یا از فایل تکمیل‌شده وارد کرد — ' +
      'هر دو مسیر به یک مجموعهٔ واحد می‌رسند. خانه‌های زردرنگ ورودی کاربر و خانه‌های خاکستری محاسباتی هستند.',
      [
        btn('＋ افزودن فرد', function () { addManualRecord(); }, 'primary'),
        btn('📥 ورود از فایل', function () { go('import'); }),
        btn('📄 دانلود تمپلیت', function () { downloadQuestionnaireTemplate(); }),
        btn('خروجی', function () { exportSheet('questionnaire'); })
      ]));

    if (!App.state.questionnaires.length) {
      main.appendChild(el('div', { class: 'empty' }, [
        el('div', { class: 'big', text: '📝' }),
        el('div', { text: 'هنوز پرسشنامه‌ای ثبت نشده است. دو راه دارید:' }),
        el('div', { style: 'margin-top:13px;display:flex;gap:8px;justify-content:center;flex-wrap:wrap' }, [
          btn('ثبت سیستمی — افزودن فرد', function () { addManualRecord(); }, 'primary'),
          btn('ورود از فایل تکمیل‌شده', function () { go('import'); })
        ])
      ]));
      return;
    }

    var cfg = App.state.config;
    main.appendChild(el('div', { class: 'legend mb' }, [
      el('span', { html: '<i class="edit"></i> ورودی کاربر' }),
      el('span', { html: '<i class="calc"></i> محاسباتی (غیرقابل ویرایش)' }),
      el('span', { class: 'muted', text: 'حداکثر امتیاز ' + cfg.maxPerformanceScore +
        ' • ' + Engine.scoredQuestions(cfg).length + ' سؤال محاسباتی' +
        ' • حد نصاب ' + cfg.minPerformanceThreshold +
        ' • اثرگذاری ویژه از امتیاز ' + cfg.specialImpactMinScore + ' به بالا' }),
      el('span', {}, [btn('طراحی پرسشنامه', function () { go('designer'); }, 'sm ghost')])
    ]));

    var answerOptions = Object.keys(cfg.answerScale);

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
        { key: 'jobLevel', label: 'JL', width: '48px', group: 'شناسایی' }
      ].concat(cfg.questions.map(function (q, i) {
        return {
          key: q.id,
          label: q.id.toUpperCase() + (q.scored === false ? ' (اطلاعاتی)' : ''),
          group: 'پاسخ سؤالات', editable: true, width: '112px',
          title: q.text + (q.scored === false
            ? '\n\nدر محاسبهٔ امتیاز عملکرد وارد نمی‌شود.'
            : '\n\nوزن: ' + (q.weight === undefined ? 1 : q.weight)),
          render: answerCell(q.id)
        };
      })).concat([
        { key: 'specialProject', label: 'اثرگذاری ویژه', group: 'اثرگذاری ویژه', editable: true,
          title: cfg.specialImpactQuestion + '\n\nتنها از امتیاز کارانه ' +
                 cfg.specialImpactMinScore + ' به بالا قابل پاسخ است.',
          render: function (r) {
            var q = questionnaireByKey(r._input._key);
            /* The gate is enforced here as well as in the engine, so the
               control is simply unavailable rather than silently ignored. */
            if (!r.specialImpactUnlocked) {
              return el('span', {
                class: 'locked-note',
                title: 'امتیاز کارانه ' + U.score(r.performanceKaraneh, 2) +
                       ' کمتر از حد نصاب ' + cfg.specialImpactMinScore + ' است.'
              }, [document.createTextNode('🔒 زیر ' + cfg.specialImpactMinScore)]);
            }
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
            if (!r.specialImpactUnlocked) {
              return el('span', { class: 'muted', text: r.specialProject ? '۰ (اعمال نشد)' : '—' });
            }
            var inp = el('input', { type: 'number', class: 'cell', step: '1',
              /* The default only applies once the flag is set; showing it on a
                 disabled cell would read as a live value. */
              placeholder: r.specialProject ? String(cfg.specialImpactAmount) : '—' });
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
      ]),
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

  /**
   * Add a person to the questionnaire set by hand. This is the systemic
   * answering path: the same record shape the importer produces, so both
   * routes converge on one dataset.
   */
  function addManualRecord() {
    var idInput = el('input', { type: 'text', class: 'editable', style: 'width:100%',
      placeholder: 'مثلاً 1024 یا BEKI001' });
    var preview = el('div', { class: 'small muted', style: 'margin-top:6px' });

    idInput.addEventListener('input', function () {
      var id = idInput.value.trim();
      U.clear(preview);
      if (!id) return;
      var existing = App.state.questionnaires.filter(function (q) {
        return q.employeeId === id && !q.excluded;
      })[0];
      if (existing) {
        preview.appendChild(el('span', { class: 'chip err',
          text: 'برای این شماره پرسنلی از قبل پرسشنامه ثبت شده است.' }));
        return;
      }
      var master = employeeById(id);
      preview.appendChild(master
        ? el('span', { class: 'chip ok', text: master.fullName + ' — ' + (master.division || 'بدون واحد') })
        : el('span', { class: 'chip warn', text: 'در اطلاعات پرسنل یافت نشد — رکورد بدون تطبیق ثبت می‌شود.' }));
    });

    U.modal({
      title: 'افزودن فرد به پرسشنامه', size: 'narrow',
      content: el('div', {}, [
        el('label', { class: 'field' }, [
          el('span', { html: 'شماره پرسنلی <b>*</b>' }), idInput
        ]),
        preview,
        el('p', { class: 'small muted', style: 'margin-bottom:0' },
          [document.createTextNode('نام و واحد سازمانی از اطلاعات پرسنل خوانده می‌شود. ' +
            'پاسخ سؤالات را پس از افزودن، مستقیماً در جدول ثبت کنید.')])
      ]),
      buttons: [
        { label: 'افزودن', kind: 'primary', keepOpen: true, onClick: function (close) {
          var id = idInput.value.trim();
          if (!id) { U.toast('شماره پرسنلی را وارد کنید.', 'err'); return false; }
          if (App.state.questionnaires.some(function (q) {
            return q.employeeId === id && !q.excluded;
          })) { U.toast('برای این شماره پرسنلی از قبل پرسشنامه ثبت شده است.', 'err'); return false; }

          var master = employeeById(id);
          var rec = {
            employeeId: id,
            fullName: (master && master.fullName) || '',
            division: (master && master.division) || '',
            positionTitle: (master && master.positionTitle) || '',
            jobLevel: (master && master.jobLevel) || '',
            sourceFile: 'ثبت سیستمی',
            importedAt: new Date().toISOString(),
            _key: 'q' + (App._keySeq = (App._keySeq || 0) + 1)
          };
          App.state.questionnaires.push(rec);
          Store.audit(App.state, {
            entity: 'questionnaire', employeeId: id, employeeName: rec.fullName,
            field: 'record', oldValue: '', newValue: 'ثبت سیستمی',
            reason: 'افزودن دستی فرد به مجموعه پرسشنامه'
          });
          save().then(function () {
            recalc();
            U.toast('فرد اضافه شد — اکنون پاسخ سؤالات را در جدول ثبت کنید.', 'ok', 5000);
          });
          close();
          return false;
        } },
        { label: 'انصراف' }
      ]
    });
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
      rows: App.result.rows.filter(inScopeForRole),
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
    if (!phase1Ready()) {
      main.appendChild(head('تغییرات معاون بخش', ''));
      main.appendChild(U.alert('err', 'این مرحله هنوز باز نشده است',
        phase1Blockers().length + ' مورد در مرحلهٔ ۱ باز است. تعیین مبلغ روی داده‌های ناقص ' +
        'می‌تواند سهم سایر افراد را جابه‌جا کند، بنابراین تا رفع آن‌ها این صفحه قفل است.',
        btn('مشاهده موارد', function () { go('validation'); }, 'sm')));
      return;
    }
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

    var eligible = App.result.rows.filter(function (r) {
      return r.inScope && r.eligible && inScopeForRole(r);
    });
    if (!isAdmin() && roleScope()) {
      main.appendChild(U.alert('info', 'دامنهٔ دسترسی شما',
        'واحدهای ' + roleScope().join('، ') + ' — ' + eligible.length + ' نفر قابل تعیین مبلغ.'));
    }
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
    if (!phase1Ready()) {
      U.toast('تا رفع خطاهای مرحلهٔ ۱، تعیین مبلغ ممکن نیست.', 'err', 5000);
      go('validation');
      return;
    }
    if (!inScopeForRole(r)) {
      U.toast('این فرد خارج از واحدهای تحت مسئولیت شماست.', 'err', 5000);
      return;
    }
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
      var g = byCode[i.code] || (byCode[i.code] = {
        code: i.code, title: i.title, severity: i.severity, stage: i.stage, items: []
      });
      g.items.push(i);
    });

    main.appendChild(phase1Ready()
      ? U.alert('ok', 'مرحلهٔ ۱ کامل است',
          'اطلاعات پایه و پاسخ‌ها آمادهٔ محاسبه‌اند. مرحلهٔ ۲ برای معاونان بخش باز است.')
      : U.alert('err', phase1Blockers().length + ' مورد پیش از شروع فرآیند باید برطرف شود',
          'تا زمانی که این موارد باز باشند، صفحهٔ «تغییرات معاون بخش» قفل می‌ماند تا ' +
          'تصمیم‌گیری روی داده‌های ناقص انجام نشود.'));

    var strip = el('div', { class: 'kpi-grid' });
    strip.appendChild(U.kpi('خطا', U.int(issueCount('err')), { kind: issueCount('err') ? 'err' : 'ok' }));
    strip.appendChild(U.kpi('هشدار', U.int(issueCount('warn')), { kind: issueCount('warn') ? 'warn' : 'ok' }));
    strip.appendChild(U.kpi('اطلاع‌رسانی', U.int(issueCount('info')), { kind: 'info' }));
    strip.appendChild(U.kpi('وضعیت بودجه',
      App.result.totals.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر',
      { kind: App.result.totals.budgetStatus === 'BALANCED' ? 'ok' : 'err' }));
    strip.appendChild(U.kpi('مرحلهٔ ۱', phase1Ready() ? 'کامل' : phase1Blockers().length + ' مورد باز',
      { kind: phase1Ready() ? 'ok' : 'err', sub: 'کنترل ورود به مرحلهٔ ۲' }));
    main.appendChild(strip);

    if (!issues.length) {
      main.appendChild(U.alert('ok', 'هیچ مورد بازی وجود ندارد', 'سیستم آماده نهایی‌سازی است.'));
    }

    var order = { err: 0, warn: 1, info: 2 };
    var stageOrder = { data: 0, payment: 1 };
    var lastStage = null;
    Object.keys(byCode)
      .sort(function (a, b) {
        var sa = stageOrder[byCode[a].stage] - stageOrder[byCode[b].stage];
        return sa || (order[byCode[a].severity] - order[byCode[b].severity]);
      })
      .forEach(function (code) {
        var g = byCode[code];
        if (g.stage !== lastStage) {
          lastStage = g.stage;
          main.appendChild(el('h2', { style: 'font-size:14px;margin:18px 0 9px;font-weight:700' },
            [document.createTextNode(g.stage === 'data'
              ? 'مرحلهٔ ۱ — اطلاعات و پاسخ‌ها'
              : 'مرحلهٔ ۲ — محاسبه و تغییرات معاون بخش')]));
        }
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
      numberField('تعداد سؤالات عملکردی', 'questionCount', '1',
        'مخرج تبدیل امتیاز به عدد کارانه — با افزودن یا حذف سؤال خودکار به‌روز می‌شود'),
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

    /* The answer scale lives in the questionnaire designer, where it sits
       beside the questions it scores. Duplicating the editor here would give
       two places to change one thing. */
    main.appendChild(U.card('نمودار ارزیابی و سؤالات', el('div', {}, [
      el('p', { class: 'small muted', style: 'margin-top:0',
        text: 'متن و وزن سؤالات، نگاشت پاسخ به امتیاز، و شرط سؤال اثرگذاری ویژه ' +
              'در صفحهٔ «طراحی پرسشنامه» تنظیم می‌شوند.' }),
      el('div', { class: 'scale-preview' }, Object.keys(cfg.answerScale).map(function (k) {
        return el('span', { html: U.esc(k) + ' <b>' + cfg.answerScale[k] + '</b>' });
      })),
      el('div', { style: 'margin-top:12px' }, [
        btn('رفتن به طراحی پرسشنامه', function () { go('designer'); }, 'primary')
      ])
    ]), { hint: Engine.scoredQuestions(cfg).length + ' سؤال محاسباتی' }));

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
   * Template downloads
   * ====================================================================*/

  /** Employees a questionnaire template should be pre-filled with. */
  function templateRoster(filter) {
    var rows = App.state.employees.filter(function (e) {
      if (!isPayrollEligible(e)) return false;
      return !filter || filter(e);
    });
    if (!rows.length && !App.state.employees.length) {
      /* No master file yet — fall back to whoever is already in the system. */
      rows = App.state.questionnaires.filter(function (q) { return !q.excluded; });
    }
    return rows.map(function (e) {
      return {
        employeeId: e.employeeId,
        fullName: e.fullName || ((e.firstName || '') + ' ' + (e.lastName || '')).trim(),
        division: e.division || '',
        positionTitle: e.positionTitle || '',
        jobLevel: e.jobLevel || ''
      };
    }).sort(function (a, b) { return U.naturalCompare(a.employeeId, b.employeeId); });
  }

  function downloadQuestionnaireTemplate(opts) {
    opts = opts || {};
    var roster = opts.prefill === false ? [] : templateRoster(opts.filter);
    var wb = Tpl.buildQuestionnaireTemplate(App.state.config, roster, {
      XLSX: XLSX, period: App.state.period, scopeLabel: opts.scopeLabel
    });
    var name = 'Template-Questionnaire' + (opts.suffix ? '-' + opts.suffix : '') +
               '-' + stamp() + '.xlsx';
    writeWorkbook(wb, name);
    Store.audit(App.state, {
      entity: 'template', field: 'questionnaire', oldValue: '',
      newValue: name + ' (' + roster.length + ' نفر)',
      reason: 'دانلود تمپلیت پرسشنامه — امضای ' + Tpl.signature(App.state.config)
    });
    save();
    U.toast('تمپلیت با ' + roster.length + ' نفر تولید شد.', 'ok');
  }

  /** One template per division, so each team lead gets only their own people. */
  function downloadQuestionnaireTemplatePerDivision() {
    var divisions = {};
    templateRoster().forEach(function (e) {
      (divisions[e.division || 'بدون واحد'] || (divisions[e.division || 'بدون واحد'] = [])).push(e);
    });
    var names = Object.keys(divisions);
    if (!names.length) { U.toast('فهرست پرسنلی برای تفکیک وجود ندارد.', 'warn'); return; }

    var body = el('div', {});
    body.appendChild(el('p', { class: 'small muted', style: 'margin-top:0',
      text: 'برای هر واحد یک فایل جداگانه تولید می‌شود. واحدهای موردنظر را انتخاب کنید.' }));
    var chosen = {};
    names.sort(function (a, b) { return a.localeCompare(b, 'fa'); }).forEach(function (d) {
      var cb = el('input', { type: 'checkbox' });
      cb.checked = true; chosen[d] = true;
      cb.addEventListener('change', function () { chosen[d] = cb.checked; });
      body.appendChild(el('label', { class: 'checkline' }, [
        cb, el('span', { text: d + ' — ' + divisions[d].length + ' نفر' })
      ]));
    });

    U.modal({
      title: 'تمپلیت به تفکیک واحد سازمانی', size: 'narrow', content: body,
      buttons: [
        { label: 'تولید فایل‌ها', kind: 'primary', onClick: function () {
          var made = 0;
          names.forEach(function (d, i) {
            if (!chosen[d]) return;
            /* Stagger the saves: browsers drop bursts of simultaneous downloads. */
            setTimeout(function () {
              downloadQuestionnaireTemplate({
                filter: function (e) { return (e.division || 'بدون واحد') === d; },
                scopeLabel: d, suffix: safeFileName(d)
              });
            }, made * 450);
            made++;
          });
          if (!made) U.toast('هیچ واحدی انتخاب نشد.', 'warn');
        } },
        { label: 'انصراف' }
      ]
    });
  }

  function downloadEmployeeTemplate() {
    var wb = Tpl.buildEmployeeTemplate(App.state.config, App.state.employees, {
      XLSX: XLSX, period: App.state.period
    });
    var name = 'Template-Employee-Master-' + stamp() + '.xlsx';
    writeWorkbook(wb, name);
    Store.audit(App.state, {
      entity: 'template', field: 'employee', oldValue: '', newValue: name,
      reason: 'دانلود تمپلیت اطلاعات پرسنل (هماهنگ با فایل حقوق و دستمزد)'
    });
    save();
    U.toast('تمپلیت اطلاعات پرسنل تولید شد.', 'ok');
  }

  function safeFileName(s) {
    return String(s).replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '-').slice(0, 40);
  }

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

  /* ------------------------------------------------------------------------
   * Payroll-format export
   * ----------------------------------------------------------------------
   * The dashboard's headline output. Column for column the same file the
   * payroll team sends in, with Final Karaneh filled — so it goes straight
   * back to them with no reshaping, and can be handed to any management level
   * scoped to just their own people.
   * ---------------------------------------------------------------------- */

  /**
   * Whole-rial payouts that still add up to the budget.
   *
   * The engine works in floating point; a payroll file must be integers. Naive
   * rounding leaves the file a few rial off the budget, which the payroll team
   * would have to explain. Largest-remainder allocation hands those spare rial
   * to the rows that were rounded down hardest, so every group sums exactly
   * and the groups sum to the whole.
   *
   * Computed once over the entire population, so a per-manager split cannot
   * drift against the total.
   */
  function roundedPayouts() {
    var rows = App.result.rows.filter(function (r) { return r.inScope; });
    var out = {}, floors = [], sumFloor = 0, exact = 0;

    rows.forEach(function (r) {
      var v = r.finalKaraneh || 0;
      var f = Math.floor(v);
      out[r.employeeId] = f;
      sumFloor += f;
      exact += v;
      floors.push({ id: r.employeeId, frac: v - f });
    });

    var remainder = Math.round(exact) - sumFloor;
    floors.sort(function (a, b) { return b.frac - a.frac; });
    for (var i = 0; i < remainder && i < floors.length; i++) out[floors[i].id] += 1;
    return out;
  }

  function payrollSheet(rows, payouts) {
    var byId = {};
    App.result.rows.forEach(function (r) { byId[r.employeeId] = r; });
    payouts = payouts || roundedPayouts();

    var cols = Tpl.PAYROLL_COLUMNS;
    var aoa = [cols.map(function (c) { return c.label; })];
    var sum = 0, finalCol = 0, daysCol = 0;
    cols.forEach(function (c, i) {
      if (c.key === 'finalKaraneh') finalCol = i;
      if (c.key === 'workingDays') daysCol = i;
    });

    rows.forEach(function (e) {
      var r = byId[e.employeeId];
      var paid = r ? (payouts[e.employeeId] || 0) : null;
      sum += paid || 0;
      aoa.push(cols.map(function (c) {
        if (c.key === 'finalKaraneh') return paid === null ? '' : paid;
        if (c.key === 'comment') {
          /* Keep the payroll team's own note, and add ours only when there is
             something they need to know. */
          var notes = [];
          if (e.comment) notes.push(e.comment);
          if (!r) notes.push('بدون پرسشنامه');
          else if (r.isOverridden) notes.push('تغییر معاون بخش: ' + (r.hodComment || '—'));
          else if (!r.eligible && r.hasQuestionnaire) notes.push('امتیاز زیر حد نصاب');
          return notes.join(' | ');
        }
        var v = e[c.key];
        return v === null || v === undefined ? '' : v;
      }));
    });

    aoa.push([]);
    var totalRow = cols.map(function () { return ''; });
    totalRow[0] = 'جمع';
    totalRow[finalCol] = sum;
    aoa.push(totalRow);

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = cols.map(function (c) { return { wch: c.width }; });
    var fmt = {};
    fmt[colLetter(finalCol)] = '#,##0';
    fmt[colLetter(daysCol)] = '#,##0';
    ws['!postprocess'] = { xSplit: 1, ySplit: 1, headerRow: 1, numberFormats: fmt };
    return { ws: ws, total: sum, count: rows.length };
  }

  /** Master rows for the people in scope, in employee-number order. */
  function payrollRoster(filter) {
    var byId = {};
    App.state.employees.forEach(function (e) { byId[e.employeeId] = e; });
    var seen = {}, out = [];

    App.result.rows.forEach(function (r) {
      if (!r.inScope || !inScopeForRole(r)) return;
      if (filter && !filter(r, byId[r.employeeId])) return;
      seen[r.employeeId] = 1;
      out.push(byId[r.employeeId] || {
        employeeId: r.employeeId, firstName: '', lastName: r.fullName,
        division: r.division, positionTitle: r.positionTitle, jobLevel: r.jobLevel,
        employeeStatus: 'Active'
      });
    });
    /* People with no questionnaire still belong in the payroll file, with a
       blank amount, so the payroll team sees the whole population. */
    App.state.employees.forEach(function (e) {
      if (seen[e.employeeId] || !isPayrollEligible(e)) return;
      var sc = roleScope();
      if (!isAdmin() && sc && sc.indexOf(e.division) === -1) return;
      if (filter && !filter(null, e)) return;
      out.push(e);
    });
    return out.sort(function (a, b) { return U.naturalCompare(a.employeeId, b.employeeId); });
  }

  function exportPayrollFile() {
    var rows = payrollRoster();
    if (!rows.length) { U.toast('رکوردی برای خروجی وجود ندارد.', 'warn'); return; }
    var built = payrollSheet(rows);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, built.ws, 'Karaneh');
    wb.Workbook = { Views: [{ RTL: false }] };
    var name = 'Karaneh-Payroll-' + safeFileName(App.state.period) + '-' + stamp() + '.xlsx';
    writeWorkbook(wb, name);
    Store.audit(App.state, {
      entity: 'export', field: 'payroll', oldValue: '',
      newValue: name + ' — ' + built.count + ' نفر، ' + U.money(built.total) + ' ریال',
      reason: 'خروجی کارانه در قالب فایل حقوق و دستمزد'
    });
    save();
    U.toast(built.count + ' رکورد در قالب فایل حقوق و دستمزد تولید شد.', 'ok', 5000);
  }

  /**
   * One workbook, one sheet per manager at the chosen level, so each manager
   * can be sent only their own people.
   */
  function exportByManager() {
    var levels = [
      { key: 'directManager', label: 'مدیر مستقیم' },
      { key: 'managerLevel1', label: 'مدیر سطح ۱' },
      { key: 'managerLevel2', label: 'مدیر سطح ۲' },
      { key: 'managerLevel3', label: 'مدیر سطح ۳' }
    ];
    var chosen = 'directManager';
    var mode = 'sheets';

    var counts = el('div', { class: 'small muted', style: 'margin-top:8px' });
    function renderCounts() {
      var roster = payrollRoster();
      var groups = {};
      roster.forEach(function (e) {
        var k = e[chosen] || '— بدون مدیر';
        groups[k] = (groups[k] || 0) + 1;
      });
      var keys = Object.keys(groups);
      U.clear(counts);
      counts.appendChild(el('b', { text: keys.length + ' گروه' }));
      counts.appendChild(document.createTextNode(' • ' + roster.length + ' نفر' +
        (keys.length ? ' • بزرگ‌ترین گروه ' +
          Math.max.apply(null, keys.map(function (k) { return groups[k]; })) + ' نفر' : '')));
    }

    var sel = el('select', { class: 'editable', style: 'width:100%' });
    levels.forEach(function (l) { sel.appendChild(el('option', { value: l.key, text: l.label })); });
    sel.addEventListener('change', function () { chosen = sel.value; renderCounts(); });

    var modeSel = el('select', { class: 'editable', style: 'width:100%' }, [
      el('option', { value: 'sheets', text: 'یک فایل با یک شیت برای هر مدیر' }),
      el('option', { value: 'files', text: 'یک فایل جداگانه برای هر مدیر' })
    ]);
    modeSel.addEventListener('change', function () { mode = modeSel.value; });

    var body = el('div', {}, [
      el('label', { class: 'field' }, [el('span', { text: 'تفکیک بر اساس' }), sel]),
      el('label', { class: 'field' }, [el('span', { text: 'نحوهٔ تولید' }), modeSel]),
      counts,
      el('p', { class: 'small muted' }, [document.createTextNode(
        'قالب هر شیت دقیقاً همان فایل حقوق و دستمزد است، با ستون Final Karaneh تکمیل‌شده.')])
    ]);
    renderCounts();

    U.modal({
      title: 'خروجی به تفکیک سطوح مدیریتی', size: 'narrow', content: body,
      buttons: [
        { label: 'تولید خروجی', kind: 'primary', onClick: function () {
          var label = levels.filter(function (l) { return l.key === chosen; })[0].label;
          runManagerExport(chosen, label, mode);
        } },
        { label: 'انصراف' }
      ]
    });
  }

  function runManagerExport(field, label, mode) {
    var roster = payrollRoster();
    var groups = {};
    roster.forEach(function (e) {
      var k = e[field] || '— بدون مدیر';
      (groups[k] || (groups[k] = [])).push(e);
    });
    var names = Object.keys(groups).sort(function (a, b) { return a.localeCompare(b, 'fa'); });
    if (!names.length) { U.toast('گروهی برای تفکیک پیدا نشد.', 'warn'); return; }

    var used = {};
    if (mode === 'files') {
      var filePayouts = roundedPayouts();
      names.forEach(function (n, i) {
        /* Stagger the saves: browsers drop bursts of simultaneous downloads. */
        setTimeout(function () {
          var wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, payrollSheet(groups[n], filePayouts).ws, 'Karaneh');
          wb.Workbook = { Views: [{ RTL: false }] };
          writeWorkbook(wb, 'Karaneh-' + safeFileName(n) + '-' + stamp() + '.xlsx');
        }, i * 450);
      });
    } else {
      var wb = XLSX.utils.book_new();
      /* An index sheet first, so the recipient sees the whole picture. */
      var idx = [[label, 'تعداد نفرات', 'مجموع کارانه (ریال)']];
      var grand = 0, built = {};
      var payouts = roundedPayouts();
      names.forEach(function (n) {
        built[n] = payrollSheet(groups[n], payouts);
        grand += built[n].total;
        idx.push([n, built[n].count, built[n].total]);
      });
      idx.push([]);
      idx.push(['جمع کل', roster.length, grand]);
      var iws = XLSX.utils.aoa_to_sheet(idx);
      iws['!cols'] = [{ wch: 32 }, { wch: 14 }, { wch: 24 }];
      iws['!postprocess'] = { ySplit: 1, headerRow: 1, numberFormats: { B: '#,##0', C: '#,##0' } };
      XLSX.utils.book_append_sheet(wb, iws, 'فهرست');
      names.forEach(function (n, i) {
        XLSX.utils.book_append_sheet(wb, built[n].ws, safeSheetName(n, i, used));
      });
      wb.Workbook = { Views: [{ RTL: true }] };
      writeWorkbook(wb, 'Karaneh-By-' + safeFileName(label) + '-' + stamp() + '.xlsx');
    }

    Store.audit(App.state, {
      entity: 'export', field: 'byManager', oldValue: '',
      newValue: names.length + ' گروه بر اساس ' + label,
      reason: 'خروجی به تفکیک سطوح مدیریتی'
    });
    save();
    U.toast('خروجی برای ' + names.length + ' گروه تولید شد.', 'ok', 5000);
  }

  /* Exposed so the end-to-end test can drive the export without going through
     the modal, the same way `recalc` is exposed for the rest of the suite. */
  window.__runManagerExport = function (field, label, mode) {
    runManagerExport(field || 'directManager', label || 'مدیر مستقیم', mode || 'sheets');
  };

  /** Excel sheet names cap at 31 characters and reject several symbols. */
  function safeSheetName(name, i, used) {
    var base = String(name).replace(/[\\\/*?:\[\]]/g, '-').slice(0, 28) || ('گروه ' + (i + 1));
    var candidate = base, n = 2;
    while (used[candidate]) candidate = base.slice(0, 26) + '-' + (n++);
    used[candidate] = 1;
    return candidate;
  }

  /** Export exactly what the dashboard table is currently showing. */
  function exportCurrentView() {
    var grid = App.grids.unified;
    var rows = grid ? grid.getVisibleRows() : dashboardRows();
    if (!rows.length) { U.toast('ردیفی برای خروجی وجود ندارد.', 'warn'); return; }
    var byId = {};
    App.state.employees.forEach(function (e) { byId[e.employeeId] = e; });
    var roster = rows.map(function (r) {
      return byId[r.employeeId] || {
        employeeId: r.employeeId, lastName: r.fullName, division: r.division,
        positionTitle: r.positionTitle, jobLevel: r.jobLevel, employeeStatus: 'Active'
      };
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, payrollSheet(roster).ws, 'Karaneh');
    wb.Workbook = { Views: [{ RTL: false }] };
    writeWorkbook(wb, 'Karaneh-Filtered-' + stamp() + '.xlsx');
    U.toast(roster.length + ' ردیف مطابق فیلتر فعلی تولید شد.', 'ok');
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
