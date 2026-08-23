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
  /* A division head runs the same cycle HR does, only over their own people:
     they collect answers from their managers, review the numbers, set their
     budget and decide. Only the instrument itself and the system settings stay
     with HR, because those are organisation-wide. */
  var NAV = [
    /* The guide comes first: it is the screen someone opening the file for the
       first time needs, and every other screen is a step described inside it. */
    { id: 'help',           icon: '❓', label: 'راهنما' },
    { phase: 1, label: 'ورود پرسشنامه و پاسخ سؤالات' },
    { id: 'designer',       icon: '🧩', label: 'طراحی پرسشنامه', adminOnly: true },
    { id: 'employees',      icon: '👤', label: 'پرسنل' },
    { id: 'import',         icon: '📥', label: 'ورود پاسخ‌ها' },
    { id: 'questionnaires', icon: '📝', label: 'پاسخ‌ها' },
    { id: 'validation',     icon: '🛡', label: 'اعتبارسنجی' },
    { phase: 2, label: 'محاسبه کارانه و تغییرات معاون بخش' },
    { id: 'dashboard',      icon: '▦',  label: 'داشبورد' },
    { id: 'payment',        icon: '💰', label: 'پرداخت کارانه' },
    { id: 'hod',            icon: '✍️', label: 'تعیین مبلغ', needsPhase1: true },
    { id: 'reports',        icon: '📤', label: 'خروجی' },
    { group: 'سیستم' },
    { id: 'audit',          icon: '🧾', label: 'ردیابی' },
    { id: 'settings',       icon: '⚙️', label: 'تنظیمات', adminOnly: true }
  ];

  var ROLES = {
    admin: { label: 'منابع انسانی', icon: '🛠' },
    hod:   { label: 'معاون بخش', icon: '✍️' }
  };

  /** True inside a file HR generated for one division head. */
  function isHodPackage() { return !!(App.state && App.state.package); }

  /* A generated file is the same application, not a cut-down one: every screen
     is there and every setting can be changed. What differs is the data it
     carries — one group instead of the organisation — and its own storage. */

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
    /* Bind storage to this file before anything reads it, so a handover file
       never shares a slot with the HR file it came from. */
    var payload = window.__KARANEH_PACKAGE__;
    Store.setNamespace(payload ? payload.package.id : 'current');

    Store.read().then(function (saved) {
      /* A file HR generated for a division head carries its data inside it.
         It is adopted once — after that the head's own edits are what persist,
         so reopening the file does not throw their work away. */
      if (payload && (!saved || !saved.package || saved.package.id !== payload.package.id)) {
        saved = adoptPackage(payload);
      }
      App.state = saved || freshState();
      if (!App.state.config) App.state.config = defaultConfig();
      if (!App.state.columnMappings) App.state.columnMappings = cloneMappings();
      if (!App.state.role) App.state.role = 'admin';
      if (!App.state.theme) App.state.theme = 'light';
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

  /**
   * Turn an embedded package into a working state. Everything the head needs
   * travels with the file: their people, the instrument, the parameters and
   * the scope they are allowed to act on.
   */
  function adoptPackage(payload) {
    var st = freshState();
    st.package = payload.package;
    st.budgetSource = payload.budgetSource || '';
    st.period = payload.period || st.period;
    /* Full access, exactly like the file it came from: the head designs,
       configures, imports and exports without asking anyone. Only the data is
       narrowed to their own group. */
    st.role = 'admin';
    st.hodScope = [];
    st.packageScope = payload.scope || [];
    st.theme = payload.theme || 'light';
    st.employees = payload.employees || [];
    st.questionnaires = (payload.questionnaires || []).map(function (q, i) {
      var c = JSON.parse(JSON.stringify(q));
      if (!c._key) c._key = 'q' + (i + 1);
      return c;
    });
    st.importBatches = payload.importBatches || [];
    if (payload.config) {
      Object.keys(payload.config).forEach(function (k) { st.config[k] = payload.config[k]; });
    }
    if (payload.columnMappings) st.columnMappings = payload.columnMappings;
    if (payload.mail) st.mail = payload.mail;
    Store.audit(st, {
      entity: 'package', field: 'open', oldValue: '',
      newValue: payload.package.label + ' — ' + st.employees.length + ' نفر',
      reason: 'بازکردن فایل تحویلی منابع انسانی'
    });
    return st;
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
    renderRail();
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

  /**
   * Who a questionnaire is expected for: everyone on the roster.
   *
   * Non-active staff used to be left out. They are not: a person who worked
   * part of the period is still assessed, and leaving them off the template
   * meant their manager had no way to score them at all. Whether they then
   * receive money is a separate question, decided by their answers and by the
   * threshold — not by their employment status.
   */
  function isPayrollEligible(e) {
    return !!e;
  }

  function issueCount(severity) {
    return (App.validation || []).filter(function (i) { return i.severity === severity; }).length;
  }

  /* ======================================================================
   * Shell — brand bar, horizontal navigation, budget strip, progress rail
   * ====================================================================*/

  /**
   * The Irancell mark, drawn inline so it needs no asset request.
   *
   * The supplied artwork: a yellow square, a black ellipse outline, and the
   * Persian wordmark over the group mark inside it. Drawn rather than embedded
   * so it stays sharp at any size and adds nothing to the file's weight.
   */
  function logoMark() {
    var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('viewBox', '0 0 120 120');
    svg.setAttribute('width', '34');
    svg.setAttribute('height', '34');
    svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', 'ایرانسل MTN');
    svg.setAttribute('style', 'direction:ltr');
    svg.innerHTML =
      '<rect x="0" y="0" width="120" height="120" fill="#FFCC00"/>' +
      '<ellipse cx="60" cy="60" rx="49" ry="25.5" fill="none"' +
      ' stroke="#000000" stroke-width="4.8"/>' +
      /* Centred, so the anchor does not depend on the run's direction. */
      '<text x="60" y="62" text-anchor="middle" fill="#000000" direction="rtl"' +
      ' style="unicode-bidi:plaintext"' +
      ' font-family="MTN Irancell, Tahoma, sans-serif" font-weight="700"' +
      ' font-size="19">ایرانسل</text>' +
      '<text x="60" y="77" text-anchor="middle" fill="#000000"' +
      ' font-family="Arial, Helvetica, sans-serif" font-weight="700"' +
      ' font-size="12.5" letter-spacing="0.6">MTN</text>';
    return svg;
  }

  function renderShell() {
    var root = document.getElementById('app');
    U.clear(root);
    applyTheme();

    root.appendChild(el('div', { class: 'brandbar' }, [
      el('span', { class: 'logo' }, [logoMark()]),
      el('span', { class: 'title', text: 'سامانه مدیریت کارانه' }),
      el('span', { class: 'period', id: 'periodChip', text: App.state.period }),
      el('div', { class: 'spacer' }),
      el('button', {
        class: 'iconbtn', id: 'themeBtn', title: 'حالت روشن / تاریک',
        onclick: function () { toggleTheme(); }
      })
    ]));

    root.appendChild(el('nav', { class: 'navbar', id: 'navbar' }));
    root.appendChild(el('div', { class: 'strip', id: 'strip' }));

    var rail = el('aside', { class: 'rail', id: 'rail' });
    var main = el('main', { class: 'main', id: 'main' });
    root.appendChild(el('div', { class: 'shell' }, [main, rail]));

    renderNav();
    renderRail();
  }

  /* ------------------------------------------------------------------ theme */
  function applyTheme() {
    var t = (App.state && App.state.theme) || 'light';
    document.documentElement.setAttribute('data-theme', t);
    var btn = document.getElementById('themeBtn');
    if (btn) btn.textContent = t === 'dark' ? '☀' : '☾';
  }

  function toggleTheme() {
    App.state.theme = App.state.theme === 'dark' ? 'light' : 'dark';
    applyTheme();
    save();
    /* Charts read their colours from CSS custom properties at draw time, so
       they have to be redrawn rather than recoloured. */
    renderView();
  }

  /* -------------------------------------------------------------- navigation */
  function renderNav() {
    var bar = document.getElementById('navbar');
    if (!bar) return;
    U.clear(bar);

    NAV.forEach(function (n) {
      /* Phases are a separator, not a badge: a bare "1" and "2" in the bar
         told nobody anything. */
      if (n.phase || n.group) { bar.appendChild(el('span', { class: 'sep' })); return; }
      if (n.adminOnly && !isAdmin()) return;

      var locked = n.needsPhase1 && !phase1Ready();
      var badge = null;
      if (n.id === 'validation') {
        var errs = issueCount('err'), warns = issueCount('warn');
        if (errs) badge = el('span', { class: 'badge err', text: String(errs) });
        else if (warns) badge = el('span', { class: 'badge warn', text: String(warns) });
      } else if (locked) {
        badge = el('span', { class: 'badge warn', text: '🔒' });
      } else if (n.id === 'questionnaires' && App.state.questionnaires.length) {
        badge = el('span', { class: 'badge', text: String(App.state.questionnaires.length) });
      } else if (n.id === 'employees' && App.state.employees.length) {
        badge = el('span', { class: 'badge', text: String(App.state.employees.length) });
      } else if (n.id === 'hod') {
        var ov = App.result ? App.result.totals.overriddenCount : 0;
        if (ov) badge = el('span', { class: 'badge', text: String(ov) });
      }

      bar.appendChild(el('button', {
        class: 'navitem' + (App.view === n.id ? ' active' : '') + (locked ? ' locked' : ''),
        title: locked ? 'تا رفع خطاهای مرحلهٔ ۱ در دسترس نیست' : n.label,
        onclick: function () {
          if (locked) {
            U.toast('ابتدا باید خطاهای مرحلهٔ ۱ برطرف شوند.', 'warn', 5000);
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

    applyTheme();
    renderStrip();
  }

  /* ------------------------------------------------------------ budget strip */
  function renderStrip() {
    var strip = document.getElementById('strip');
    if (!strip || !App.result) return;
    var t = App.result.totals;
    U.clear(strip);

    function cell(value, label, kind) {
      strip.appendChild(el('div', { class: 'cell ' + (kind || '') }, [
        el('b', {}, [U.bidi(value)]),
        el('span', { text: label })
      ]));
      strip.appendChild(el('div', { class: 'div' }));
    }
    cell(U.money(t.budget), 'بودجه (ریال)');
    cell(U.money(t.allocatedBudget), 'تخصیص‌یافته');
    cell(U.money(t.remainingBudget), 'باقیمانده', t.remainingBudget < -1 ? 'bad' : '');
    cell(U.int(t.eligibleCount), 'واجد شرایط');
    cell(U.int(t.overriddenCount), 'تغییر معاون بخش');
    var label = t.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر';
    strip.appendChild(el('div', { class: 'cell ' + (t.budgetStatus === 'BALANCED' ? 'ok' : 'bad') }, [
      el('b', { text: label }), el('span', { text: 'وضعیت بودجه' })
    ]));
  }

  /* ======================================================================
   * Progress rail — how far the period has actually got
   * ====================================================================*/

  /**
   * The steps the process passes through, each with a completion fraction.
   * A step is only "done" when the thing it produces actually exists, so the
   * ring cannot read 100% while something is still missing.
   */
  function progressSteps() {
    var t = App.result ? App.result.totals : null;
    var staff = App.state.employees.length;
    var expected = App.state.employees.filter(isPayrollEligible).length || staff;
    var answered = App.result
      ? App.result.rows.filter(function (r) { return r.inScope && r.hasQuestionnaire; }).length
      : 0;
    var blockers = phase1Blockers().length;
    var eligible = t ? t.eligibleCount : 0;
    var reviewed = t ? t.overriddenCount : 0;

    return [
      {
        id: 'designer', label: 'طراحی پرسشنامه',
        meta: Engine.scoredQuestions(App.state.config).length + ' سؤال محاسباتی',
        done: 1
      },
      {
        id: 'employees', label: 'اطلاعات پرسنل',
        meta: staff ? U.int(staff) + ' نفر' : 'وارد نشده',
        done: staff ? 1 : 0
      },
      {
        id: 'import', label: 'دریافت پاسخ‌ها',
        meta: expected ? U.int(answered) + ' از ' + U.int(expected) : U.int(answered) + ' رکورد',
        done: expected ? Math.min(1, answered / expected) : (answered ? 1 : 0)
      },
      {
        id: 'validation', label: 'اعتبارسنجی',
        meta: blockers ? blockers + ' مورد باز' : (answered ? 'بدون خطا' : 'در انتظار داده'),
        done: answered ? (blockers ? 0 : 1) : 0,
        blocked: blockers > 0
      },
      {
        id: 'payment', label: 'محاسبه کارانه',
        meta: eligible ? U.int(eligible) + ' نفر واجد شرایط' : 'در انتظار',
        done: eligible ? 1 : 0
      },
      {
        id: 'hod', label: 'تصمیم معاون بخش',
        meta: reviewed ? U.int(reviewed) + ' مورد ثبت شده' : 'بدون تغییر',
        done: phase1Ready() && eligible ? 1 : 0
      },
      {
        id: 'validation', label: 'کنترل بودجه',
        meta: t && t.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر',
        done: t && t.budgetStatus === 'BALANCED' && eligible ? 1 : 0
      },
      {
        id: 'reports', label: 'نهایی‌سازی و ارسال',
        meta: App.state.finalizedAt ? U.dateTime(App.state.finalizedAt) : 'انجام نشده',
        done: App.state.finalizedAt ? 1 : 0
      }
    ];
  }

  function progressPercent() {
    var steps = progressSteps();
    var sum = steps.reduce(function (a, s) { return a + s.done; }, 0);
    return steps.length ? sum / steps.length : 0;
  }

  function renderRail() {
    var rail = document.getElementById('rail');
    if (!rail) return;
    U.clear(rail);

    var steps = progressSteps();
    var pct = progressPercent();

    rail.appendChild(el('h3', { text: 'پیشرفت فرآیند' }));
    rail.appendChild(el('div', { class: 'sub', text: App.state.period }));

    /* Progress ring — one number the whole period can be judged by. */
    var NS = 'http://www.w3.org/2000/svg';
    var size = 118, stroke = 11, r = (size - stroke) / 2, c = 2 * Math.PI * r;
    var svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', '0 0 ' + size + ' ' + size);
    svg.setAttribute('width', size);
    svg.setAttribute('height', size);
    svg.innerHTML =
      '<circle cx="' + size / 2 + '" cy="' + size / 2 + '" r="' + r + '" fill="none"' +
      ' stroke="var(--surface-3)" stroke-width="' + stroke + '"/>' +
      '<circle cx="' + size / 2 + '" cy="' + size / 2 + '" r="' + r + '" fill="none"' +
      ' stroke="var(--brand)" stroke-width="' + stroke + '" stroke-linecap="round"' +
      ' stroke-dasharray="' + c + '" stroke-dashoffset="' + (c * (1 - pct)) + '"' +
      ' transform="rotate(-90 ' + size / 2 + ' ' + size / 2 + ')"/>' +
      '<text class="pct" x="' + size / 2 + '" y="' + (size / 2 + 2) + '"' +
      ' text-anchor="middle" dominant-baseline="middle">' + Math.round(pct * 100) + '٪</text>' +
      '<text class="pctlabel" x="' + size / 2 + '" y="' + (size / 2 + 20) + '"' +
      ' text-anchor="middle">تکمیل شده</text>';
    rail.appendChild(el('div', { class: 'ringwrap' }, [svg]));

    var list = el('ol', { class: 'steplist' });
    var firstOpen = steps.filter(function (s) { return s.done < 1; })[0];
    steps.forEach(function (s, i) {
      var cls = s.done >= 1 ? 'done' : (s === firstOpen ? 'active' : '');
      if (s.blocked) cls += ' blocked';
      list.appendChild(el('li', {
        class: 'step ' + cls,
        title: 'رفتن به ' + s.label,
        onclick: function () { if (canOpen(s.id)) go(s.id); else go('validation'); }
      }, [
        el('span', { class: 'dot', text: s.done >= 1 ? '✓' : String(i + 1) }),
        el('span', {}, [
          el('div', { class: 'label', text: s.label }),
          el('div', { class: 'meta', text: s.meta })
        ])
      ]));
    });
    rail.appendChild(list);

    if (App.state.finalizedAt) {
      rail.appendChild(el('div', { class: 'railnote',
        text: 'دوره نهایی شده است. فایل‌های خروجی از صفحهٔ «گزارش و خروجی» قابل دریافت‌اند.' }));
    } else if (firstOpen) {
      rail.appendChild(el('div', { class: 'railnote',
        text: 'مرحلهٔ بعد: ' + firstOpen.label + ' — ' + firstOpen.meta }));
    }
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
    var fn = VIEWS[App.view] || VIEWS.dashboard;
    fn(main);
  }

  /* There is no role switch any more: HR's file and the files it generates
     are the same application, and which file you hold is the answer. `role`
     stays in the state — every file is 'admin' — so the scope helpers and the
     stored data of older files keep working unchanged. */

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
      'متن سؤالات، حوزه، لنگرهای رفتاری، وزن و شرط اثرگذاری ویژه. تمپلیت و محاسبات بر همین اساس ساخته می‌شوند.',
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

        var domain = el('input', {
          type: 'text', class: 'editable', style: 'width:190px;font-weight:700',
          placeholder: 'حوزه'
        });
        domain.value = q.domain || '';
        domain.addEventListener('change', function () {
          if (domain.value.trim() === q.domain) return;
          auditConfig('question.' + q.id + '.domain', q.domain, domain.value.trim());
          q.domain = domain.value.trim();
          save(); recalc();
        });

        row.appendChild(el('div', { class: 'qhead' }, [
          el('span', { class: 'code', text: q.id.toUpperCase() }),
          domain,
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
          q.impact
            ? el('span', { class: 'chip brand', text: 'این سؤال امتیاز ویژه را تعیین می‌کند' })
            : el('label', {}, [scoredCb, el('span', { text: 'در محاسبهٔ امتیاز عملکرد وارد شود' })]),
          q.impact ? null : el('label', {}, [el('span', { text: 'وزن' }), weight])
        ]));

        /* Behavioural anchors — the rater picks one of these, so they are the
           substance of the instrument, not decoration. */
        var anchors = el('div', { class: 'bars-grid' });
        Object.keys(cfg.answerScale).forEach(function (optName, i) {
          if (!q.anchors) q.anchors = [];
          var a = q.anchors[i] || (q.anchors[i] = { label: '', text: '' });
          var lab = el('input', {
            type: 'text', class: 'cell', placeholder: 'برچسب رفتاری',
            style: 'font-weight:700'
          });
          lab.value = a.label || '';
          lab.addEventListener('change', function () {
            auditConfig('question.' + q.id + '.anchor' + (i + 1) + '.label', a.label, lab.value.trim());
            a.label = lab.value.trim();
            save(); recalc();
          });
          var desc = el('textarea', { class: 'editable', placeholder: 'شرح رفتار در این سطح' });
          desc.value = a.text || '';
          desc.addEventListener('change', function () {
            auditConfig('question.' + q.id + '.anchor' + (i + 1), a.text, desc.value.trim());
            a.text = desc.value.trim();
            save(); recalc();
          });
          anchors.appendChild(el('div', { class: 'bars-cell' }, [
            el('div', { class: 'lvl' }, [
              el('span', { class: 'n', text: String(cfg.answerScale[optName]) }),
              el('span', { text: optName })
            ]),
            lab, desc
          ]));
        });
        row.appendChild(anchors);
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

    var step = Number(cfg.specialImpactStep) || 50;
    var siAmount = el('input', {
      type: 'number', class: 'editable', step: String(step), min: String(step),
      style: 'width:100%'
    });
    siAmount.value = cfg.specialImpactAmount;
    siAmount.addEventListener('change', function () {
      var v = Number(siAmount.value);
      if (!isFinite(v)) { siAmount.value = cfg.specialImpactAmount; return; }
      /* Snap on the way in, so the stored value is always on the scale. */
      v = Engine.snapToStep(v, cfg);
      siAmount.value = v;
      setConfig('specialImpactAmount', v);
    });

    var siStep = el('input', { type: 'number', class: 'editable', step: '5', min: '1', style: 'width:100%' });
    siStep.value = step;
    siStep.addEventListener('change', function () {
      var v = Number(siStep.value);
      if (!isFinite(v) || v < 1) { siStep.value = step; return; }
      setConfig('specialImpactStep', v);
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
          el('span', { text: 'حداقل امتیاز کارانه برای فعال شدن' }), siMin,
          el('div', { class: 'small muted', style: 'margin-top:3px',
            text: 'پاسخ به آن زیر این عدد امتیازی نمی‌گیرد' })
        ]),
        el('label', { class: 'field' }, [
          el('span', { text: 'حداکثر امتیاز اثرگذاری ویژه' }), siAmount,
          el('div', { class: 'small muted', style: 'margin-top:3px', text: 'ستون N' })
        ]),
        el('label', { class: 'field' }, [
          el('span', { text: 'گام امتیاز' }), siStep,
          el('div', { class: 'small muted', style: 'margin-top:3px',
            text: 'امتیاز فقط مضربی از این عدد می‌گیرد' })
        ])
      ]),
      (function () {
        /* The permitted bands, spelled out — this is what the rater picks. */
        var wrap = el('div', { class: 'scale-preview' });
        var st = Number(cfg.specialImpactStep) || 50;
        for (var v = st; v <= (Number(cfg.specialImpactAmount) || st); v += st) {
          wrap.appendChild(el('span', { text: U.score(v, 0) }));
        }
        return el('div', {}, [
          el('div', { class: 'small muted', text: 'امتیازهای مجاز:' }), wrap
        ]);
      }()),
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
        btn('تمپلیت خالی', function () { downloadQuestionnaireTemplate({ prefill: false }); }),
        btn('تمپلیت با فهرست پرسنل', function () { downloadQuestionnaireTemplate({ prefill: true }); }, 'primary'),
        btn('به تفکیک سطح شغلی', function () { downloadQuestionnaireTemplateByGroup('jobLevel'); }),
        btn('به تفکیک واحد سازمانی', function () { downloadQuestionnaireTemplateByGroup('division'); })
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
        btn('خروجی کارانه', function () { exportPayrollFile(); }, 'primary'),
        btn('خروجی به تفکیک سطح / مدیر', function () { exportByManager(); }),
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
      { kind: missing ? 'warn' : '', sub: 'از کل پرسنل' }));
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

    /* In a file that covers one division, a chart broken down by division is
       one bar. What the head of that division actually compares is their
       managers — the same names the questionnaires were split by. */
    var breakdownField = App.state.breakdownField || defaultBreakdownField(scoped);
    var breakdownPick = el('select', { class: 'editable', style: 'font-size:12px' },
      BREAKDOWNS.map(function (b) {
        return el('option', { value: b.key, text: b.label });
      }));
    breakdownPick.value = breakdownField;
    breakdownPick.addEventListener('change', function () {
      App.state.breakdownField = breakdownPick.value;
      save();
      renderView();
    });

    var charts = el('div', { class: 'chart-grid-2' }, [
      U.card(el('span', { style: 'display:flex;align-items:center;gap:8px' }, [
        el('span', { text: 'کارانه به تفکیک' }), breakdownPick
      ]), divisionHost, { hint: 'مجموع پرداختی هر گروه (ریال)' }),
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
      }, breakdownField);
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

  var BREAKDOWNS = [
    { key: 'division',      label: 'واحد سازمانی' },
    { key: 'directManager', label: 'مدیر مستقیم' },
    { key: 'managerLevel1', label: 'مدیر سطح ۱' },
    { key: 'managerLevel2', label: 'مدیر سطح ۲' },
    { key: 'positionTitle', label: 'عنوان شغلی' }
  ];

  /**
   * What to break the payout chart down by, before anyone has chosen.
   * One division in the file means the division chart says nothing, so the
   * manager names — the ones the questionnaires were split by — are used.
   */
  function defaultBreakdownField(rows) {
    var divisions = {};
    rows.forEach(function (r) { if (r.division) divisions[r.division] = 1; });
    if (Object.keys(divisions).length > 1) return 'division';
    var managed = rows.filter(function (r) { return managerOf(r, 'directManager'); }).length;
    return managed ? 'directManager' : 'division';
  }

  function breakdownValue(row, field) {
    if (field === 'division') return row.division || '';
    if (field === 'positionTitle') return row.positionTitle || '';
    return managerOf(row, field);
  }

  function drawDashboardCharts(rows, hosts, breakdownField) {
    var cfg = App.state.config;
    var field = breakdownField || 'division';

    /* Magnitude by group — horizontal, because unit and manager names are long. */
    var byDiv = {};
    rows.forEach(function (r) {
      var d = breakdownValue(r, field) || '— ثبت نشده';
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
    /* Whoever has the full view loads the roster here from the payroll list. */
    var loadsRoster = isAdmin();

    main.appendChild(head('اطلاعات پرسنل',
      'شماره پرسنلی کلید یکتاست. رکورد تکراری بدون تأیید شما وارد نمی‌شود.',
      (isAdmin() ? [
        btn('ورود فایل پرسنل', function () { pickFiles('employee'); }, 'primary'),
        isHodPackage() ? null : btn('تولید فایل معاون بخش', function () { openPackageBuilder(); }),
        btn('تمپلیت حقوق و دستمزد', function () { downloadEmployeeTemplate(); }),
        btn('خروجی', function () { exportSheet('employees'); }),
        App.state.employees.length ? btn('پاک کردن', function () { clearEmployees(); }, 'danger') : null
      ] : [
        btn('تفکیک پرسشنامه بین مدیران', function () { openManagerSplit(); }, 'primary'),
        btn('خروجی', function () { exportSheet('employees'); })
      ]).filter(Boolean)));

    /* Say plainly where the roster comes from: the payroll list is the file
       people arrive with, and this is the screen it belongs on. */
    if (loadsRoster) {
      main.appendChild(U.card('فایل پرسنل را از کجا بیاورم؟', el('div', {}, [
        el('p', { class: 'mb', html:
          '<b>فایلی را که از تیم حقوق و دستمزد گرفته‌اید، همین‌جا بارگذاری کنید.</b> ' +
          'همان فهرست پرسنل با شمارهٔ پرسنلی، نام، سطح شغلی، واحد سازمانی، مدیر مستقیم و روز کارکرد.' }),
        el('ul', { class: 'small muted', style: 'margin:0;padding-inline-start:18px;line-height:1.9' }, [
          el('li', { text: 'عنوان ستون‌ها هرچه باشد، سامانه خودش آن‌ها را تشخیص می‌دهد؛ فارسی و انگلیسی هر دو.' }),
          el('li', { text: 'اگر ستونی شناخته نشد، در «اعتبارسنجی» گزارش می‌شود و می‌توانید نگاشت را اصلاح کنید.' }),
          el('li', { text: 'ورود دوبارهٔ همان فایل، رکورد تکراری نمی‌سازد؛ فقط اطلاعات به‌روز می‌شود.' })
        ])
      ]), { hint: 'قالب خالی را هم می‌توانید از دکمهٔ «تمپلیت حقوق و دستمزد» بگیرید' }));
    }

    /* Splitting the questionnaire among the managers who report to this head
       is the first thing they do, so it sits on the roster screen itself. */
    if (App.state.employees.length) {
      main.appendChild(U.card(
        (isAdmin() && !isHodPackage()) ? 'تحویل فایل و تفکیک پرسشنامه' : 'تفکیک پرسشنامه بین مدیران',
        el('div', { style: 'display:flex;gap:8px;flex-wrap:wrap;align-items:center' }, [
          (isAdmin() && !isHodPackage())
            ? btn('تولید فایل معاون بخش', function () { openPackageBuilder(); }, 'primary')
            : null,
          btn('تمپلیت به تفکیک مدیر مستقیم', function () { openManagerSplit('directManager'); }),
          btn('به تفکیک مدیر سطح ۱', function () { openManagerSplit('managerLevel1'); }),
          btn('به تفکیک سطح شغلی', function () { downloadQuestionnaireTemplateByGroup('jobLevel'); }),
          btn('همهٔ حالت‌ها', function () { openManagerSplit(); }, 'ghost')
        ].filter(Boolean)),
        { hint: (isAdmin() && !isHodPackage())
          ? 'هر معاون یک فایل مستقل از سامانه، و هر مدیر یک فایل پرسشنامه با پرسنل خودش می‌گیرد'
          : 'هر مدیر یک فایل پرسشنامه با پرسنل خودش دریافت می‌کند' }));
    }



    if (!App.state.employees.length) {
      if (loadsRoster) {
        main.appendChild(dropzoneNode('employee',
          'فایل پرسنلِ دریافتی از حقوق و دستمزد را اینجا رها کنید',
          'ستون‌های Emp No، نام، سطح شغلی، واحد سازمانی و … به‌صورت خودکار شناسایی می‌شوند.'));
      } else {
        main.appendChild(U.card('پرسنلی در این فایل نیست', el('div', { class: 'small muted',
          text: 'این فایل را منابع انسانی برای شما ساخته است و فهرست پرسنل باید داخل آن باشد. ' +
                'اگر خالی است، به منابع انسانی اطلاع دهید تا فایل را دوباره تولید کند.' })));
      }
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

  /**
   * Split the questionnaire template among managers.
   *
   * A division head does not fill four hundred questionnaires themselves —
   * they hand each of their managers the people who report to that manager.
   * Any of the manager levels works, because org charts differ in depth.
   */
  function openManagerSplit(preset) {
    var fields = [
      { key: 'directManager', label: 'مدیر مستقیم' },
      { key: 'managerLevel1', label: 'مدیر سطح ۱' },
      { key: 'managerLevel2', label: 'مدیر سطح ۲' },
      { key: 'managerLevel3', label: 'مدیر سطح ۳' },
      { key: 'jobLevel',      label: 'سطح شغلی' },
      { key: 'division',      label: 'واحد سازمانی' }
    ];
    var field = preset || 'directManager';
    var chosen = {};

    var listBox = el('div', { style: 'max-height:280px;overflow-y:auto;margin-top:10px' });
    var summary = el('div', { class: 'small muted', style: 'margin-top:8px' });

    function roster() {
      return templateRoster(function (e) { return inScopeForRole({ division: e.division }); });
    }

    function groupsFor(f) {
      var g = {}, source = App.state.employees.filter(function (e) {
        return isPayrollEligible(e) && inScopeForRole({ division: e.division });
      });
      source.forEach(function (e) {
        var k = e[f];
        if (!k) { g['— ثبت نشده'] = (g['— ثبت نشده'] || 0) + 1; return; }
        g[k] = (g[k] || 0) + 1;
      });
      return g;
    }

    function renderGroups() {
      var g = groupsFor(field);
      var names = Object.keys(g).sort(function (a, b) {
        return field === 'jobLevel' ? U.naturalCompare(b, a) : a.localeCompare(b, 'fa');
      });
      chosen = {};
      U.clear(listBox);
      if (!names.length) {
        listBox.appendChild(el('div', { class: 'small muted',
          text: 'برای این تفکیک، مقداری در اطلاعات پرسنل ثبت نشده است.' }));
        U.clear(summary);
        return;
      }
      names.forEach(function (n) {
        chosen[n] = n.indexOf('— ثبت نشده') !== 0;
        var cb = el('input', { type: 'checkbox' });
        cb.checked = chosen[n];
        cb.addEventListener('change', function () { chosen[n] = cb.checked; updateSummary(); });
        listBox.appendChild(el('label', { class: 'checkline' }, [
          cb, el('span', { text: n + ' — ' + g[n] + ' نفر' })
        ]));
      });
      updateSummary();
    }

    function updateSummary() {
      var g = groupsFor(field);
      var picked = Object.keys(chosen).filter(function (k) { return chosen[k]; });
      var people = picked.reduce(function (a, k) { return a + (g[k] || 0); }, 0);
      U.clear(summary);
      summary.appendChild(el('b', { text: picked.length + ' فایل پرسشنامه' }));
      summary.appendChild(document.createTextNode(' • ' + people + ' نفر'));
    }

    var sel = el('select', { class: 'editable', style: 'width:100%' });
    fields.forEach(function (f) { sel.appendChild(el('option', { value: f.key, text: f.label })); });
    sel.value = field;
    sel.addEventListener('change', function () { field = sel.value; renderGroups(); });

    var body = el('div', {}, [
      el('label', { class: 'field' }, [el('span', { text: 'تفکیک بر اساس' }), sel]),
      listBox, summary,
      el('div', { class: 'small muted', style: 'margin-top:10px' },
        [document.createTextNode(
          'هر فایل شامل شیت BARS، فهرست کشویی پاسخ‌ها و فهرست پرسنل همان گروه است.')])
    ]);
    renderGroups();

    U.modal({
      title: 'تفکیک پرسشنامه', size: 'narrow', content: body,
      buttons: [
        { label: 'تولید فایل‌ها', kind: 'primary', onClick: function () {
          var picked = Object.keys(chosen).filter(function (k) { return chosen[k]; });
          if (!picked.length) { U.toast('هیچ گروهی انتخاب نشد.', 'warn'); return; }
          var meta = fields.filter(function (f) { return f.key === field; })[0];
          picked.forEach(function (name, i) {
            setTimeout(function () {
              downloadQuestionnaireTemplate({
                filter: function (e) {
                  var v = e[field] || '— ثبت نشده';
                  return String(v) === name && inScopeForRole({ division: e.division });
                },
                scopeLabel: meta.label + ': ' + name,
                suffix: safeFileNameOr(name, 'group'),
                quiet: i < picked.length - 1
              });
            }, i * 520);
          });
          if (picked.length > 1) {
            U.toast(picked.length + ' فایل در حال تولید است — همه در صفحهٔ دانلود می‌مانند.',
              'warn', 8000);
          }
        } },
        { label: 'انصراف' }
      ]
    });
  }

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
      'چند فایل تکمیل‌شده را همزمان انتخاب کنید؛ ساختار با تمپلیت مقایسه و همه در یک مجموعه ادغام می‌شود.',
      [btn('انتخاب فایل‌ها', function () { pickFiles('questionnaire'); }, 'primary')]));

    main.appendChild(U.card('دانلود تمپلیت', el('div', {
      style: 'display:flex;gap:8px;flex-wrap:wrap'
    }, [
      btn('تمپلیت با فهرست پرسنل', function () { downloadQuestionnaireTemplate({ prefill: true }); }, 'primary'),
      btn('به تفکیک سطح شغلی', function () { downloadQuestionnaireTemplateByGroup('jobLevel'); }),
      btn('به تفکیک واحد سازمانی', function () { downloadQuestionnaireTemplateByGroup('division'); }),
      btn('تمپلیت خالی', function () { downloadQuestionnaireTemplate({ prefill: false }); }),
      btn('طراحی پرسشنامه', function () { go('designer'); }, 'ghost')
    ])));

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
      'پاسخ‌ها را در همین جدول ثبت کنید یا از فایل تکمیل‌شده وارد کنید.',
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

    /** Anchor text for one question at one answer, for tooltips and the panel. */
    function anchorFor(question, answer) {
      if (!question || !question.anchors) return null;
      var score = cfg.answerScale[String(answer).trim()];
      if (score === undefined) return null;
      var idx = answerOptions.indexOf(String(answer).trim());
      return question.anchors[idx] || null;
    }

    function answerCell(qKey) {
      var question = cfg.questions.filter(function (x) { return x.id === qKey; })[0];
      return function (r) {
        var q = questionnaireByKey(r._input._key);
        var sel = el('select', { class: 'cell' });
        sel.appendChild(el('option', { value: '', text: '—' }));
        answerOptions.forEach(function (o, i) {
          var a = question && question.anchors ? question.anchors[i] : null;
          sel.appendChild(el('option', {
            value: o,
            /* The behavioural label is what the rater is actually choosing;
               the wording alone does not say what it means. */
            text: (a && a.label ? a.label + ' — ' : '') + o,
            title: a ? a.text : ''
          }));
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
        function syncTitle() {
          var a = anchorFor(question, sel.value);
          sel.title = a ? (a.label + ' — ' + a.text) : (question ? question.text : '');
        }
        syncTitle();
        sel.addEventListener('change', function () {
          syncTitle();
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
      ].concat(cfg.questions.filter(function (q) { return !q.impact; }).map(function (q) {
        return {
          key: q.id,
          label: q.domain || q.id.toUpperCase(),
          group: 'پاسخ سؤالات', editable: true, width: '130px',
          title: q.id.toUpperCase() + ' — ' + q.text +
            '\n\nوزن: ' + (q.weight === undefined ? 1 : q.weight),
          render: answerCell(q.id)
        };
      })).concat([
        { key: 'specialProject',
          label: (cfg.questions.filter(function (q) { return q.impact; })[0] || {}).domain || 'اثرگذاری ویژه',
          group: 'اثرگذاری ویژه', editable: true,
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
            /* The scale moves in fixed steps, so this is a list of the
               permitted bands rather than a free number. */
            var step = Number(cfg.specialImpactStep) || 50;
            var max = Number(cfg.specialImpactAmount) || step;
            var sel = el('select', { class: 'cell' });
            sel.appendChild(el('option', { value: '', text: '—' }));
            for (var v = step; v <= max; v += step) {
              sel.appendChild(el('option', { value: String(v), text: U.score(v, 0) }));
            }
            var cur = q && q.specialImpactAmount
              ? String(Engine.snapToStep(Number(q.specialImpactAmount), cfg)) : '';
            sel.value = cur;
            sel.disabled = !r.specialProject;
            sel.addEventListener('change', function () {
              editField(q, 'specialImpactAmount', sel.value === '' ? null : Number(sel.value),
                'تغییر امتیاز اثرگذاری ویژه');
            });
            return sel;
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

  /**
   * Budget and the parameters that shape it, on the payment screen.
   *
   * These used to live in Settings, which a division head cannot open. They
   * belong next to the numbers they move: the head types a budget and watches
   * the table underneath change, rather than setting a value elsewhere and
   * navigating back to see what it did.
   */
  function budgetPanel() {
    var cfg = App.state.config;
    var box = el('div', {});

    function liveNumber(label, key, step, hint, format) {
      var isMoney = !!format;
      var inp = isMoney
        ? U.moneyInput({ style: 'width:100%;font-weight:700', value: cfg[key] })
        : el('input', {
            type: 'number', class: 'editable', step: step || 'any',
            style: 'width:100%;font-weight:700'
          });
      if (!isMoney) inp.value = cfg[key];
      function read() { return isMoney ? inp.getNumber() : Number(inp.value); }
      function reset() {
        if (isMoney) inp.setNumber(cfg[key]); else inp.value = cfg[key];
      }
      /* The hint already sits next to the label; the echo underneath is only
         for a live readback of the typed number, so it stays empty without a
         formatter rather than repeating the hint. */
      var echo = el('div', { class: 'small muted', style: 'margin-top:3px' });
      function renderEcho() {
        var v = read();
        echo.textContent = format && v !== null ? format(Number(v)) : '';
      }
      renderEcho();
      inp.addEventListener('input', renderEcho);
      inp.addEventListener('change', function () {
        var v = read();
        if (v === null || !isFinite(v) || v < 0) { reset(); renderEcho(); return; }
        if (Number(cfg[key]) === v) return;
        setConfig(key, v);
        renderView();
      });
      /* The hint goes under the control, not into the label. A hint long
         enough to wrap made its label two lines tall and pushed that one
         field's input below the others in the row. */
      return el('label', { class: 'field' }, [
        el('span', { text: label }),
        inp,
        hint ? el('div', { class: 'small muted', style: 'margin-top:3px', text: hint }) : null,
        echo
      ].filter(Boolean));
    }

    box.appendChild(el('div', { class: 'form-grid' }, [
      liveNumber('بودجه کل (ریال)', 'budget', '1000000',
        App.state.budgetSource || null, function () { return ''; }),
      liveNumber('حداقل امتیاز جهت دریافت کارانه', 'minPerformanceThreshold', '0.25',
        'امتیاز کمتر یا مساوی این عدد، کارانهٔ صفر'),
      liveNumber('حداکثر امتیاز کارانه', 'maxPerformanceScore', '10',
        'مقیاس تبدیل امتیاز ۱ تا ۵')
    ]));

    var t = App.result.totals;
    box.appendChild(el('div', { class: 'budget-bar' }, [
      el('i', {
        class: t.budgetOverrun ? 'over' : 'used',
        style: 'width:' + (t.budget
          ? Math.min(100, t.allocatedBudget / t.budget * 100).toFixed(2) : 0) + '%'
      })
    ]));
    box.appendChild(el('div', { class: 'budget-legend' }, [
      el('span', { html: 'تخصیص‌یافته: <b class="num">' + U.money(t.allocatedBudget) + '</b>' }),
      el('span', { html: 'باقیمانده: <b class="num">' + U.money(t.remainingBudget) + '</b>' }),
      el('span', {}, [
        document.createTextNode('وضعیت: '),
        el('span', {
          class: 'chip ' + (t.budgetStatus === 'BALANCED' ? 'ok' : 'err'),
          text: t.budgetStatus === 'BALANCED' ? 'متوازن' : 'نامعتبر'
        })
      ])
    ]));
    return box;
  }

  /**
   * What the budget actually bought, by job level.
   *
   * A division head reviews band by band, so this is the summary they need
   * before touching an individual: how many people at each level, what they
   * cost in total, and what the average and range look like inside the band.
   */
  function levelSummary(rows) {
    var cfg = App.state.config;
    var byLevel = {};
    rows.forEach(function (r) {
      if (!r.inScope) return;
      var k = r.jobLevel || 'نامشخص';
      var g = byLevel[k] || (byLevel[k] = {
        level: k, grade: r.gradeScore, count: 0, eligible: 0, overridden: 0,
        score: 0, amount: 0, min: Infinity, max: -Infinity, perf: 0, perfN: 0
      });
      g.count++;
      if (r.eligible) g.eligible++;
      if (r.isOverridden) g.overridden++;
      g.score += r.totalScore;
      g.amount += r.finalKaraneh;
      if (r.performanceScore !== null) { g.perf += r.performanceScore; g.perfN++; }
      if (r.eligible) {
        g.min = Math.min(g.min, r.finalKaraneh);
        g.max = Math.max(g.max, r.finalKaraneh);
      }
    });

    var levels = Object.keys(byLevel).map(function (k) { return byLevel[k]; })
      .sort(function (a, b) { return (b.grade || 0) - (a.grade || 0) || U.naturalCompare(b.level, a.level); });
    var total = levels.reduce(function (a, g) { return a + g.amount; }, 0);

    var host = el('div', {});
    var chartHost = el('div', {});
    host.appendChild(chartHost);

    var tbl = el('table', { class: 'grid' });
    tbl.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', { text: 'سطح شغلی' }), el('th', { text: 'عدد گرید' }),
      el('th', { text: 'نفرات' }), el('th', { text: 'واجد شرایط' }),
      el('th', { text: 'میانگین امتیاز' }), el('th', { text: 'مجموع دریافتی (ریال)' }),
      el('th', { text: 'میانگین هر نفر' }), el('th', { text: 'کمترین' }),
      el('th', { text: 'بیشترین' }), el('th', { text: 'سهم از بودجه' })
    ])]));
    var tb = el('tbody');
    levels.forEach(function (g) {
      tb.appendChild(el('tr', {}, [
        el('td', { class: 'mono', text: g.level }),
        el('td', { class: 'num', text: g.grade === null ? '—' : U.score(g.grade, 0) }),
        el('td', { class: 'num', text: U.int(g.count) }),
        el('td', { class: 'num', text: U.int(g.eligible) }),
        el('td', { class: 'num', text: g.perfN ? U.score(g.perf / g.perfN) : '—' }),
        el('td', { class: 'num', text: U.money(g.amount) }),
        el('td', { class: 'num', text: U.money(g.eligible ? g.amount / g.eligible : 0) }),
        el('td', { class: 'num', text: g.eligible ? U.money(g.min) : '—' }),
        el('td', { class: 'num', text: g.eligible ? U.money(g.max) : '—' }),
        el('td', { class: 'num', text: U.percent(total ? g.amount / total : 0) })
      ]));
    });
    tbl.appendChild(tb);
    tbl.appendChild(el('tfoot', {}, [el('tr', {}, [
      el('td', { text: 'جمع' }), el('td', {}),
      el('td', { class: 'num', text: U.int(levels.reduce(function (a, g) { return a + g.count; }, 0)) }),
      el('td', { class: 'num', text: U.int(levels.reduce(function (a, g) { return a + g.eligible; }, 0)) }),
      el('td', {}),
      el('td', { class: 'num', text: U.money(total) }),
      el('td', {}), el('td', {}), el('td', {}),
      el('td', { class: 'num', text: U.percent(total ? 1 : 0) })
    ])]));
    host.appendChild(el('div', { class: 'table-wrap', style: 'max-height:none;margin-top:12px' }, [tbl]));

    /* Draw once the node has a width to measure. */
    requestAnimationFrame(function () {
      Chart.horizontalBar(chartHost, levels.map(function (g) {
        return {
          label: 'JL ' + g.level, value: g.amount,
          detail: U.money(g.amount) + ' ریال<br>' + g.count + ' نفر · میانگین ' +
                  U.money(g.eligible ? g.amount / g.eligible : 0)
        };
      }), { format: U.moneyShort, labelWidth: 92 });
    });
    return host;
  }

  /* ======================================================================
   * VIEW — روش پرداخت کارانه
   * ====================================================================*/
  VIEWS.payment = function (main) {
    var t = App.result.totals;
    main.appendChild(head('روش پرداخت کارانه',
      'معادل شیت «روش پرداخت کارانه» فایل مرجع.',
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

    var scoped = App.result.rows.filter(inScopeForRole);
    var scopedPaid = 0, scopedPeople = 0, scopedEligible = 0;
    scoped.forEach(function (r) {
      if (!r.inScope) return;
      scopedPeople++;
      scopedPaid += r.finalKaraneh;
      if (r.eligible) scopedEligible++;
    });

    var strip = el('div', { class: 'kpi-grid' });
    strip.appendChild(U.kpi('بودجه (ریال)', U.money(t.budget), { kind: 'brand' }));
    strip.appendChild(U.kpi('مجموع دریافتی', U.money(scopedPaid),
      { kind: 'ok', sub: scopedPeople + ' نفر' }));
    strip.appendChild(U.kpi('میانگین هر نفر',
      U.money(scopedEligible ? scopedPaid / scopedEligible : 0),
      { sub: scopedEligible + ' نفر واجد شرایط' }));
    strip.appendChild(U.kpi('امتیاز کل', U.score(t.sumTotalScore, 2), { kind: 'info' }));
    strip.appendChild(U.kpi('سرشکن زیر حد نصاب', U.score(t.ineligibleRedistribution, 4),
      { sub: 'به ازای هر فرد واجد شرایط' }));
    strip.appendChild(U.kpi('سرشکن تغییرات معاون', U.money(t.hodRedistribution),
      { kind: t.hodRedistribution < 0 ? 'warn' : '' }));
    main.appendChild(strip);

    /* The budget first, then what it produced, then the roster: the number is
       the input to this screen, so it sits at the top of it. */
    main.appendChild(U.card('بودجه و پارامترها', budgetPanel(),
      { hint: 'عدد را همین‌جا وارد کنید؛ گزارش و جدول پایین لحظه‌ای بازمحاسبه می‌شوند' }));
    main.appendChild(U.card('مجموع دریافتی به تفکیک سطح شغلی', levelSummary(scoped),
      { hint: 'با هر تغییر بودجه یا امتیاز، بلافاصله به‌روز می‌شود' }));


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
      'تعیین مبلغ نهایی هر فرد. ثبت توضیح اجباری است و اختلاف بین سایرین سرشکن می‌شود.'));

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

    var amount = U.moneyInput({
      style: 'width:100%',
      value: r.isOverridden ? r.hodAdjustment : ''
    });
    var comment = el('textarea', {
      class: 'editable', style: 'width:100%;min-height:70px',
      placeholder: 'بر اساس چه اثرگذاری خاصی این مبلغ تعیین شده است؟'
    });
    comment.value = r.hodComment || '';

    var preview = el('div', {});
    function renderPreview() {
      U.clear(preview);
      var v = amount.getNumber();
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
          var v = amount.getNumber();
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
        /* Finalising is the moment the two teams need their files. */
        setTimeout(function () { sendFinalPackage(); }, 400);
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
   * VIEW — Help
   * ----------------------------------------------------------------------
   * Written for the division head, because they are the one who receives a
   * file and has to work out what to do with it without anyone beside them.
   * The steps mirror the progress rail exactly, so the two never disagree.
   * ====================================================================*/
  VIEWS.help = function (main) {
    var cfg = App.state.config;
    var scope = roleScope();

    main.appendChild(head('راهنما',
      isHodPackage() || !isAdmin()
        ? 'آنچه باید انجام دهید، به ترتیب.'
        : 'مرور فرآیند، از طراحی پرسشنامه تا ارسال فایل نهایی.'));

    if (isHodPackage()) {
      main.appendChild(U.alert('info', 'این فایل مخصوص شماست',
        'صادرشده برای «' + App.state.package.label + '» در تاریخ ' +
        U.dateTime(App.state.package.issuedAt) + ' — ' +
        U.int(App.state.employees.length) + ' نفر' +
        (scope ? ' • واحدهای ' + scope.join('، ') : '') +
        '. کار شما در همین مرورگر ذخیره می‌شود؛ فایل را جای امنی نگه دارید.'));
    }

    function stepCard(n, title, lines, action) {
      return el('div', {
        style: 'display:flex;gap:12px;padding:13px 15px;border:1px solid var(--border);' +
               'border-radius:9px;margin-bottom:10px;background:var(--surface)'
      }, [
        el('span', {
          style: 'width:26px;height:26px;border-radius:50%;flex:0 0 26px;display:grid;' +
                 'place-items:center;background:var(--brand);color:var(--brand-ink);' +
                 'font-weight:700;font-size:12px'
        }, [document.createTextNode(String(n))]),
        el('div', { style: 'flex:1' }, [
          el('b', { text: title }),
          el('ul', { style: 'margin:6px 0 0;padding-inline-start:18px;line-height:1.9;font-size:12.5px' },
            lines.map(function (t) { return el('li', { text: t }); })),
          action ? el('div', { style: 'margin-top:9px' }, [action]) : null
        ])
      ]);
    }

    var steps = el('div', {});

    if (isAdmin()) {
      /* HR's own sequence. The guide is the first tab now, so it has to open
         on what HR actually does, not on what a division head does. */
      steps.appendChild(stepCard(1, 'پرسشنامه را طراحی کنید', [
        'سؤال‌ها، وزن هر سؤال، مقیاس پاسخ و شرح رفتاری هر سطح (BARS) در «طراحی پرسشنامه» تعیین می‌شود.',
        'همین طراحی مبنای تمپلیت Excel و همهٔ فایل‌های معاونان است.'
      ], btn('طراحی پرسشنامه', function () { go('designer'); }, 'sm')));

      steps.appendChild(stepCard(2, 'فایل پرسنل را بارگذاری کنید', [
        'فایلی را که از تیم حقوق و دستمزد گرفته‌اید، در بخش «پرسنل» بارگذاری کنید.',
        'همان فهرست پرسنل با شمارهٔ پرسنلی، نام، سطح شغلی، واحد سازمانی، مدیر مستقیم و روز کارکرد.',
        'عنوان ستون‌ها هرچه باشد شناسایی می‌شود؛ قالب خالی هم از همان صفحه قابل دریافت است.'
      ], btn('رفتن به پرسنل', function () { go('employees'); }, 'sm primary')));

      /* Issuing handover files is HR's job. Inside one of those files the step
         is noise — its holder splits the questionnaire among their managers
         instead, which is the step below. */
      steps.appendChild(isHodPackage()
        ? stepCard(3, 'پرسشنامه را بین مدیران خود تقسیم کنید', [
            'در «پرسنل» گزینهٔ «به تفکیک مدیر مستقیم» را بزنید.',
            'برای هر مدیر یک فایل Excel جداگانه با پرسنل خودش تولید می‌شود.',
            'هر فایل شیت BARS را دارد: شرح رفتاری هر سطح، برای اینکه همه یکسان امتیاز بدهند.'
          ], btn('تفکیک پرسشنامه', function () { go('employees'); }, 'sm'))
        : stepCard(3, 'برای هر معاون یک فایل بسازید', [
            'در «پرسنل» گزینهٔ «تولید فایل معاون بخش» را بزنید و مبنای تفکیک را انتخاب کنید.',
            'هر معاون یک فایل مستقل با پرسنل خودش می‌گیرد.',
            'اگر می‌خواهید خودتان پاسخ‌ها را جمع کنید، به‌جای آن تمپلیت پرسشنامه را تفکیک کنید.'
          ], btn('تولید فایل معاون بخش', function () { go('employees'); }, 'sm')));

      steps.appendChild(stepCard(4, 'پاسخ‌ها را وارد و بررسی کنید', [
        'فایل‌های برگشتی را در «ورود پاسخ‌ها» انتخاب کنید — چندتایی هم می‌شود.',
        '«اعتبارسنجی» هر چیزی را که مانع محاسبه است فهرست می‌کند.',
        'تا وقتی خطای قرمز باز باشد، صفحهٔ «تعیین مبلغ» قفل می‌ماند.'
      ], btn('اعتبارسنجی', function () { go('validation'); }, 'sm')));

      steps.appendChild(stepCard(5, 'بودجه را وارد کنید', [
        'در «پرداخت کارانه»، بودجهٔ کل را در همان صفحه وارد کنید.',
        'بالای آن، مجموع دریافتی به تفکیک سطح شغلی را می‌بینید.',
        'با هر تغییر بودجه، همهٔ اعداد بلافاصله بازمحاسبه می‌شوند.'
      ], btn('پرداخت کارانه', function () { go('payment'); }, 'sm')));

      steps.appendChild(stepCard(6, 'تغییرات معاونان را ثبت کنید', [
        'در «تعیین مبلغ»، مبلغ نهایی هر فرد قابل تغییر است و ثبت توضیح اجباری است.',
        'اختلاف روی بقیه سرشکن می‌شود؛ مجموع پرداخت همیشه برابر بودجه می‌ماند.'
      ], btn('تعیین مبلغ', function () { go('hod'); }, 'sm')));

      steps.appendChild(stepCard(7, 'نهایی و ارسال کنید', [
        'دو فایل تولید می‌شود: یکی برای تیم عملکرد (بدون مبلغ) و یکی برای جبران خدمات.',
        'نشانی‌های ایمیل در «تنظیمات → گیرندگان ارسال نهایی» ثبت می‌شوند.',
        'فایل‌ها در صفحهٔ دانلود گوشهٔ پایین هم می‌مانند؛ اگر دانلود خودکار شروع نشد، آنجا کلیک کنید.'
      ], btn('خروجی', function () { go('reports'); }, 'sm')));

    } else {

    steps.appendChild(stepCard(1, 'پرسنل خود را ببینید', [
          'فهرست افرادی که در این فایل هستند در بخش «پرسنل» آمده است.',
      'این همان فهرستی است که منابع انسانی از فایل حقوق و دستمزد گرفته است.',
      'اگر کسی جا افتاده یا اضافه است، می‌توانید فایل پرسنل را همان‌جا دوباره بارگذاری کنید.'
    ], btn('رفتن به پرسنل', function () { go('employees'); }, 'sm')));

    steps.appendChild(stepCard(2, 'پرسشنامه را بین مدیران خود تقسیم کنید', [
      'در بخش «پرسنل» گزینهٔ «تفکیک پرسشنامه بین مدیران» را بزنید.',
      'برای هر مدیر یک فایل Excel جداگانه با پرسنل خودش تولید می‌شود.',
      'هر فایل شیت BARS را دارد: شرح رفتاری هر سطح، برای اینکه همه یکسان امتیاز بدهند.'
    ], btn('تفکیک پرسشنامه', function () { openManagerSplit(); }, 'sm primary')));

    steps.appendChild(stepCard(3, 'فایل‌های تکمیل‌شده را برگردانید', [
      'فایل‌هایی که مدیران پر کرده‌اند را در «ورود پاسخ‌ها» انتخاب کنید — چندتایی هم می‌شود.',
      'اگر ساختار فایل با تمپلیت یکی نباشد، وارد نمی‌شود و علت دقیق گفته می‌شود.',
      'می‌توانید پاسخ‌ها را مستقیماً در بخش «پاسخ‌ها» هم ثبت یا اصلاح کنید.'
    ], btn('ورود پاسخ‌ها', function () { go('import'); }, 'sm')));

    steps.appendChild(stepCard(4, 'خطاها را برطرف کنید', [
      'بخش «اعتبارسنجی» هر چیزی که مانع محاسبه است را فهرست می‌کند.',
      'تا وقتی خطای قرمز باز باشد، صفحهٔ «تعیین مبلغ» قفل می‌ماند.',
      'دلیلش این است که تخصیص نسبی است: محاسبه روی جمعیت ناقص، سهم بقیه را جابه‌جا می‌کند.'
    ], btn('اعتبارسنجی', function () { go('validation'); }, 'sm')));

    steps.appendChild(stepCard(5, 'بودجه را وارد کنید', [
      'در «پرداخت کارانه»، بودجهٔ خود را در همان صفحه وارد کنید.',
      'بالای آن، مجموع دریافتی به تفکیک سطح شغلی را می‌بینید.',
      'با هر تغییر بودجه، همهٔ اعداد بلافاصله بازمحاسبه می‌شوند.'
    ], btn('پرداخت کارانه', function () { go('payment'); }, 'sm')));

    steps.appendChild(stepCard(6, 'مبلغ افراد را تعیین کنید', [
      'در «تعیین مبلغ» می‌توانید مبلغ نهایی هر فرد را دستی وارد کنید.',
      'ثبت توضیح اجباری است.',
      'پیش از ثبت می‌بینید این تغییر چقدر از سهم بقیه کم می‌کند و سقف مجاز چقدر است.',
      'مجموع پرداخت همیشه دقیقاً برابر بودجه می‌ماند.'
    ], btn('تعیین مبلغ', function () { go('hod'); }, 'sm')));

    steps.appendChild(stepCard(7, 'نهایی و ارسال کنید', [
      'وقتی وضعیت بودجه «متوازن» بود و خطایی نماند، نهایی‌سازی کنید.',
      'دو فایل تولید می‌شود: یکی برای تیم عملکرد (بدون مبلغ) و یکی برای جبران خدمات.',
      'فایل‌ها در صفحهٔ دانلود گوشهٔ پایین هم می‌مانند؛ اگر دانلود خودکار شروع نشد، آنجا کلیک کنید.'
    ], btn('خروجی', function () { go('reports'); }, 'sm')));

    }

    main.appendChild(U.card('گام‌به‌گام', steps));

    /* The instrument, so the head can answer "what does 3 mean?" */
    var barsBody = el('div', {});
    barsBody.appendChild(el('p', { class: 'small muted', style: 'margin-top:0',
      text: 'برای هر حوزه، رفتاری را انتخاب کنید که بیشترین شباهت را به عملکرد واقعی فرد ' +
            'در این دوره دارد — نه بهترین یا بدترین روز او.' }));
    var options = Object.keys(cfg.answerScale);
    var bt = el('table', { class: 'grid bars-scale-table' });
    bt.appendChild(el('thead', {}, [el('tr', {},
      [el('th', { text: 'حوزه' })].concat(options.map(function (o) {
        return el('th', { text: cfg.answerScale[o] + ' — ' + o });
      })))]));
    var btb = el('tbody');
    cfg.questions.forEach(function (q) {
      btb.appendChild(el('tr', {},
        [el('td', {}, [
          el('b', { text: q.domain || q.id.toUpperCase() }),
          el('div', { class: 'small muted', style: 'white-space:normal', text: q.text })
        ])].concat(options.map(function (o, i) {
          var a = (q.anchors || [])[i];
          return el('td', { class: 'anchor' }, [
            a ? el('span', { class: 'anchor-label', text: a.label }) : null,
            a ? document.createTextNode(a.text) : document.createTextNode('—')
          ].filter(Boolean));
        }))));
    });
    bt.appendChild(btb);
    barsBody.appendChild(el('div', { class: 'table-wrap', style: 'max-height:none' }, [bt]));
    main.appendChild(U.card('مقیاس رفتاری (BARS)', barsBody, { tight: false }));

    /* The handful of rules that decide someone's money. */
    var rules = [
      ['امتیاز عملکرد', 'میانگین وزنی ' + Engine.scoredQuestions(cfg).length +
        ' حوزهٔ امتیازدهی، در بازهٔ ۱ تا ۵.'],
      ['عدد کارانه', 'امتیاز عملکرد × ' + cfg.maxPerformanceScore + ' ÷ ' + cfg.questionCount + '.'],
      ['اثرگذاری ویژه', 'تنها از عدد کارانهٔ ' + cfg.specialImpactMinScore +
        ' به بالا قابل ثبت است و امتیاز آن مضربی از ' + (cfg.specialImpactStep || 50) + ' است.'],
      ['حد نصاب', 'امتیاز کمتر یا مساوی ' + cfg.minPerformanceThreshold +
        ' یعنی کارانهٔ صفر؛ سهم آن بین بقیه سرشکن می‌شود.'],
      ['سطح شغلی', cfg.gradeImpactFactor
        ? 'با ضریب ' + cfg.gradeImpactFactor + ' در امتیاز کل اثر دارد.'
        : 'با ضریب صفر، در حال حاضر بر مبلغ اثری ندارد.'],
      ['تقسیم بودجه', 'بودجه به نسبت «امتیاز کل» هر فرد تقسیم می‌شود.'],
      ['تغییر دستی', 'هر ریالی که به یک نفر اضافه شود، از سهم بقیه به‌طور مساوی کم می‌شود.']
    ];
    var dl = el('dl', { class: 'kv' });
    rules.forEach(function (r) {
      dl.appendChild(el('dt', { text: r[0] }));
      dl.appendChild(el('dd', { style: 'font-weight:400', text: r[1] }));
    });
    main.appendChild(U.card('قواعد محاسبه', dl));

    main.appendChild(U.card('سؤالات پرتکرار', el('dl', { class: 'kv' }, [
      el('dt', { text: 'کارم ذخیره می‌شود؟' }),
      el('dd', { style: 'font-weight:400',
        text: 'بله، در همین مرورگر. با بستن صفحه از بین نمی‌رود، اما پاک کردن داده‌های مرورگر آن را حذف می‌کند.' }),
      el('dt', { text: 'دکمهٔ دانلود کار نمی‌کند' }),
      el('dd', { style: 'font-weight:400',
        text: 'فایل ساخته شده است. صفحهٔ دانلود گوشهٔ پایین را ببینید و روی نام فایل کلیک کنید.' }),
      el('dt', { text: 'چرا نمی‌توانم اثرگذاری ویژه ثبت کنم؟' }),
      el('dd', { style: 'font-weight:400',
        text: 'عدد کارانهٔ آن فرد به حد نصاب ' + cfg.specialImpactMinScore + ' نرسیده است.' }),
      el('dt', { text: 'چرا صفحهٔ تعیین مبلغ باز نمی‌شود؟' }),
      el('dd', { style: 'font-weight:400',
        text: 'هنوز خطایی در اعتبارسنجی باز است. آن را برطرف کنید.' }),
      el('dt', { text: 'مجموع پرداخت از بودجه بیشتر می‌شود؟' }),
      el('dd', { style: 'font-weight:400',
        text: 'نه. مجموع همیشه برابر بودجه می‌ماند. اگر مبلغ یک نفر را خیلی بالا ببرید، ' +
              'دریافتی دیگران منفی می‌شود و سامانه جلوی نهایی‌سازی را می‌گیرد.' })
    ])));
  };

  /* ======================================================================
   * VIEW — Settings
   * ====================================================================*/
  VIEWS.settings = function (main) {
    var cfg = App.state.config;
    main.appendChild(head('تنظیمات',
      'هیچ‌یک از این مقادیر در کد ثابت نشده است. تغییر هر کدام، کل محاسبات را بلافاصله بازمحاسبه می‌کند.'));

    function numberField(label, key, step, hint) {
      var isMoney = key === 'budget';
      var inp = isMoney
        ? U.moneyInput({ style: 'width:100%', value: cfg[key] })
        : el('input', { type: 'number', class: 'editable', style: 'width:100%', step: step || 'any' });
      if (!isMoney) inp.value = cfg[key];
      inp.addEventListener('change', function () {
        var v = isMoney ? inp.getNumber() : Number(inp.value);
        if (v === null || !isFinite(v)) {
          if (isMoney) inp.setNumber(cfg[key]); else inp.value = cfg[key];
          return;
        }
        setConfig(key, v);
      });
      return el('label', { class: 'field' }, [
        el('span', { text: label }),
        inp,
        hint ? el('div', { class: 'small muted', style: 'margin-top:3px', text: hint }) : null
      ].filter(Boolean));
    }

    var params = el('div', { class: 'form-grid' }, [
      numberField('بودجه کل (ریال)', 'budget', '1000000', 'سلول C1 فایل مرجع'),
      numberField('حداکثر امتیاز کارانه', 'maxPerformanceScore', '1', 'مقدار مرجع: 120'),
      numberField('تعداد سؤالات عملکردی', 'questionCount', '1',
        'مخرج تبدیل امتیاز به عدد کارانه — با افزودن یا حذف سؤال خودکار به‌روز می‌شود'),
      numberField('حداقل امتیاز جهت دریافت', 'minPerformanceThreshold', '0.25', 'سلول D4'),
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
      el('span', { text: 'نحوه اعمال حد نصاب' }), mode,
      el('div', { class: 'small muted', style: 'margin-top:3px', text: 'ستون I' })
    ]));
    params.appendChild(el('label', { class: 'field' }, [
      el('span', { text: 'دامنه متناسب‌سازی ضرایب' }), scope,
      el('div', { class: 'small muted', style: 'margin-top:3px', text: '«امتیاز کل بخش»' })
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


    /* -- grade impact ---------------------------------------------------
       Its own panel rather than one field among many: turning grade on moves
       money between people, so the number needs its effect shown beside it
       before it is committed. */
    var gradeBox = el('div', {});
    var previewBox = el('div', {});

    var factorInput = el('input', {
      type: 'number', class: 'editable', step: '0.05', min: '0',
      style: 'width:120px;font-size:15px;font-weight:700'
    });
    factorInput.value = cfg.gradeImpactFactor;

    var slider = el('input', {
      type: 'range', min: '0', max: '2', step: '0.05', style: 'flex:1;min-width:180px'
    });
    slider.value = cfg.gradeImpactFactor;

    /* Pending value: the preview follows the control live, but nothing is
       written until the user commits. */
    var pending = Number(cfg.gradeImpactFactor);

    function setPending(v) {
      pending = isFinite(v) && v >= 0 ? v : 0;
      factorInput.value = pending;
      slider.value = Math.min(2, pending);
      renderGradePreview();
    }
    factorInput.addEventListener('input', function () { setPending(Number(factorInput.value)); });
    slider.addEventListener('input', function () { setPending(Number(slider.value)); });

    function commitFactor() {
      if (Number(cfg.gradeImpactFactor) === pending) {
        U.toast('این مقدار هم‌اکنون اعمال شده است.', 'warn', 2500);
        return;
      }
      setConfig('gradeImpactFactor', pending);
      renderView();
    }

    /**
     * Run the engine at the pending factor and diff it against the live one,
     * so the panel can answer the only question that matters: who gains, who
     * loses, and by how much.
     */
    function renderGradePreview() {
      U.clear(previewBox);
      var rows = App.result.rows.filter(function (r) { return r.inScope; });
      if (!rows.length) {
        previewBox.appendChild(el('div', { class: 'small muted',
          text: 'برای پیش‌نمایش اثر، ابتدا پرسشنامه‌ها را وارد کنید.' }));
        return;
      }

      var inputs = rows.map(function (r) { return r._input; });
      var trial = Engine.calculate(inputs, mergeCfg(cfg, { gradeImpactFactor: pending }));
      var current = Engine.calculate(inputs, mergeCfg(cfg, { gradeImpactFactor: cfg.gradeImpactFactor }));

      var before = {}, after = {};
      current.rows.forEach(function (r) { before[r.employeeId] = r.finalKaraneh; });
      trial.rows.forEach(function (r) { after[r.employeeId] = r.finalKaraneh; });

      /* Share of the score pool that grade would control. */
      var gradeShare = trial.totals.sumTotalScore
        ? trial.rows.reduce(function (a, r) { return a + r.gradeImpact; }, 0) / trial.totals.sumTotalScore
        : 0;

      var moved = 0, movers = [];
      trial.rows.forEach(function (r) {
        if (!r.inScope) return;
        var d = (after[r.employeeId] || 0) - (before[r.employeeId] || 0);
        if (Math.abs(d) > 0.5) moved += Math.abs(d);
        movers.push({ row: r, delta: d });
      });
      movers.sort(function (a, b) { return b.delta - a.delta; });

      var byLevel = {};
      movers.forEach(function (m) {
        var k = m.row.jobLevel || '—';
        var g = byLevel[k] || (byLevel[k] = { n: 0, delta: 0, grade: m.row.gradeScore });
        g.n++; g.delta += m.delta;
      });

      var dl = el('dl', { class: 'kv', style: 'margin-bottom:12px' });
      [
        ['ضریب فعلی', String(cfg.gradeImpactFactor)],
        ['ضریب پیش‌نمایش', String(pending)],
        ['سهم گرید از امتیاز کل', U.percent(gradeShare, 1)],
        ['مبلغ جابه‌جاشده', U.money(moved / 2) + ' ریال'],
        ['مجموع پرداخت', U.money(trial.totals.sumFinalKaraneh) + ' ریال']
      ].forEach(function (l) {
        dl.appendChild(el('dt', { text: l[0] }));
        dl.appendChild(el('dd', { class: 'num', text: l[1] }));
      });
      previewBox.appendChild(dl);

      if (pending === Number(cfg.gradeImpactFactor)) {
        previewBox.appendChild(el('div', { class: 'small muted',
          text: 'برای دیدن اثر، ضریب را تغییر دهید.' }));
      } else {
        previewBox.appendChild(U.alert(pending > 0 ? 'warn' : 'info',
          pending > 0 ? 'اثر این تغییر' : 'حذف اثر گرید',
          pending > 0
            ? 'با این ضریب، ' + U.money(moved / 2) + ' ریال از افراد با سطح شغلی پایین‌تر ' +
              'به افراد با سطح شغلی بالاتر منتقل می‌شود. مجموع پرداخت تغییر نمی‌کند.'
            : 'سطح شغلی دیگر بر مبلغ کارانه اثری نخواهد داشت — همان وضعیت فایل مرجع.'));
      }

      /* Per job level: the honest summary, since grade acts by level. */
      var levelTbl = el('table', { class: 'grid' });
      levelTbl.appendChild(el('thead', {}, [el('tr', {}, [
        el('th', { text: 'JL' }), el('th', { text: 'عدد گرید' }), el('th', { text: 'نفرات' }),
        el('th', { text: 'تغییر مجموع (ریال)' }), el('th', { text: 'به ازای هر نفر' })
      ])]));
      var ltb = el('tbody');
      Object.keys(byLevel).sort(function (a, b) {
        return (byLevel[b].grade || 0) - (byLevel[a].grade || 0);
      }).forEach(function (k) {
        var g = byLevel[k];
        ltb.appendChild(el('tr', {}, [
          el('td', { class: 'mono', text: k }),
          el('td', { class: 'num', text: g.grade === null ? '⚠ تعریف نشده' : U.score(g.grade, 0) }),
          el('td', { class: 'num', text: U.int(g.n) }),
          el('td', { class: 'num' + (g.delta < 0 ? ' neg' : ''), text: U.money(g.delta) }),
          el('td', { class: 'num' + (g.delta < 0 ? ' neg' : ''), text: U.money(g.n ? g.delta / g.n : 0) })
        ]));
      });
      levelTbl.appendChild(ltb);
      previewBox.appendChild(el('div', { class: 'table-wrap', style: 'max-height:230px' }, [levelTbl]));
    }

    gradeBox.appendChild(el('p', { class: 'small muted', style: 'margin-top:0' },
      [document.createTextNode(
        'تأثیر گرید = عدد گرید × این ضریب، و به امتیاز عملکردی اضافه می‌شود (ستون G و L). ' +
        'در فایل مرجع این ضریب صفر است، بنابراین سطح شغلی هیچ اثری بر مبلغ ندارد. ' +
        'با افزایش آن، بودجه از سطوح پایین‌تر به سطوح بالاتر منتقل می‌شود — مجموع پرداخت ثابت می‌ماند.')]));

    gradeBox.appendChild(el('div', {
      style: 'display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:6px'
    }, [
      el('span', { class: 'small', text: 'ضریب تأثیر گرید' }),
      factorInput,
      slider,
      btn('اعمال', function () { commitFactor(); }, 'primary'),
      btn('صفر (مطابق فایل مرجع)', function () { setPending(0); }, 'sm')
    ]));
    gradeBox.appendChild(el('div', { class: 'small muted', style: 'margin-bottom:12px' },
      [document.createTextNode('مقادیر متداول: ۰ (بی‌اثر) · ۰٫۲۵ (اثر ملایم) · ۰٫۵ (اثر متوسط) · ۱ (اثر کامل)')]));
    gradeBox.appendChild(previewBox);
    renderGradePreview();

    main.appendChild(U.card('ضریب تأثیر گرید', gradeBox,
      { hint: 'سلول D2 فایل مرجع — پیش‌نمایش پیش از اعمال' }));

    /* -- grade table ---------------------------------------------------
       Sorted by score rather than insertion order, with the step to the next
       level shown: an out-of-place value is then visible at a glance instead
       of hiding in a list. */
    var gradeBody = el('div', {});
    function renderGradeTable() {
      U.clear(gradeBody);
      var keys = Object.keys(cfg.gradeMap).sort(function (a, b) {
        return cfg.gradeMap[a] - cfg.gradeMap[b];
      });

      var counts = {};
      App.result.rows.forEach(function (r) {
        if (r.inScope && r.jobLevel) counts[r.jobLevel] = (counts[r.jobLevel] || 0) + 1;
      });
      App.state.employees.forEach(function (e) {
        var jl = Engine.normalizeJobLevel(e.jobLevel);
        if (jl && counts[jl] === undefined) counts[jl] = 0;
      });

      var gt = el('table', { class: 'grid', id: 'gradeTable' });
      gt.appendChild(el('thead', {}, [el('tr', {}, [
        el('th', { text: 'JL' }), el('th', { text: 'عدد گرید' }),
        el('th', { text: 'فاصله تا سطح قبل' }), el('th', { text: 'نفرات' }), el('th', { text: '' })
      ])]));
      var gtb = el('tbody');

      keys.forEach(function (jl, i) {
        var inp = el('input', { type: 'number', class: 'cell', step: '5' });
        inp.value = cfg.gradeMap[jl];
        inp.addEventListener('change', function () {
          var v = Number(inp.value);
          if (!isFinite(v)) { inp.value = cfg.gradeMap[jl]; return; }
          var old = cfg.gradeMap[jl];
          cfg.gradeMap[jl] = v;
          auditConfig('gradeMap.' + jl, old, v);
          save(); recalc();
        });

        var step = i === 0 ? null : cfg.gradeMap[jl] - cfg.gradeMap[keys[i - 1]];
        var used = counts[jl];

        gtb.appendChild(el('tr', {}, [
          el('td', { class: 'mono', text: jl }),
          el('td', {}, [inp]),
          el('td', { class: 'num muted small',
            text: step === null ? '—' : (step === 0 ? '⚠ برابر سطح قبل' : '+' + step) }),
          el('td', { class: 'num' }, [
            used === undefined
              ? el('span', { class: 'muted', text: '—' })
              : el('span', { class: used ? '' : 'muted', text: U.int(used) })
          ]),
          el('td', {}, [el('button', {
            class: 'btn sm danger', text: 'حذف',
            title: used ? used + ' نفر از این سطح استفاده می‌کنند' : '',
            onclick: function () { removeGradeLevel(jl, used); }
          })])
        ]));
      });
      gt.appendChild(gtb);
      gradeBody.appendChild(gt);

      var newJl = el('input', { type: 'text', placeholder: 'JL مثلاً 5 یا 4H', style: 'width:120px' });
      var newScore = el('input', { type: 'number', placeholder: 'عدد گرید', step: '5', style: 'width:120px' });
      gradeBody.appendChild(el('div', { style: 'display:flex;gap:7px;margin-top:11px;align-items:center' }, [
        newJl, newScore,
        btn('افزودن', function () {
          var k = Engine.normalizeJobLevel(newJl.value);
          var v = Number(newScore.value);
          if (!k || !isFinite(v)) { U.toast('سطح شغلی و عدد گرید را وارد کنید.', 'err'); return; }
          if (cfg.gradeMap[k] !== undefined) { U.toast('این سطح از قبل تعریف شده است.', 'err'); return; }
          cfg.gradeMap[k] = v;
          auditConfig('gradeMap.' + k, '(جدید)', v);
          newJl.value = ''; newScore.value = '';
          save(); recalc();
        }, 'primary')
      ]));

      /* Levels present in the data but absent from the table. */
      var unmapped = {};
      App.result.rows.forEach(function (r) {
        if (r.inScope && r.gradeScore === null && r.jobLevel) unmapped[r.jobLevel] = 1;
      });
      App.state.employees.forEach(function (e) {
        var jl = Engine.normalizeJobLevel(e.jobLevel);
        if (jl && cfg.gradeMap[jl] === undefined) unmapped[jl] = 1;
      });
      var missingJl = Object.keys(unmapped);
      if (missingJl.length) {
        var fix = el('div', { style: 'display:flex;gap:6px;flex-wrap:wrap;margin-top:7px' });
        missingJl.forEach(function (jl) {
          fix.appendChild(btn('افزودن ' + jl, function () {
            /* Seat it between its neighbours so no existing grade moves. */
            var suggested = suggestGradeScore(jl);
            cfg.gradeMap[jl] = suggested;
            auditConfig('gradeMap.' + jl, '(جدید)', suggested);
            save(); recalc();
          }, 'sm primary'));
        });
        gradeBody.appendChild(U.alert('err', 'سطوح شغلی تعریف‌نشده',
          'این سطوح در داده‌ها وجود دارند اما در جدول گرید نیستند: ' + missingJl.join('، ') +
          ' — تا تعریف نشوند، امتیاز گرید این افراد محاسبه نمی‌شود.'));
        gradeBody.appendChild(fix);
      }

      gradeBody.appendChild(el('div', { class: 'small muted', style: 'margin-top:10px' },
        [document.createTextNode(
          'تا وقتی ضریب تأثیر گرید صفر باشد، این اعداد بر مبلغ کارانه اثری ندارند.')]));
    }

    function removeGradeLevel(jl, used) {
      var doIt = function () {
        var old = cfg.gradeMap[jl];
        delete cfg.gradeMap[jl];
        auditConfig('gradeMap.' + jl, old, '(حذف شد)');
        save(); recalc();
      };
      if (used) {
        U.confirm(used + ' نفر سطح شغلی «' + jl + '» دارند. با حذف آن، امتیاز گرید این افراد ' +
          'محاسبه نمی‌شود و در مرکز اعتبارسنجی خطا ثبت می‌شود. ادامه می‌دهید؟',
          { danger: true, confirmLabel: 'حذف' }).then(function (ok) { if (ok) doIt(); });
      } else doIt();
    }

    /**
     * A starting figure for a level the table does not cover: the midpoint of
     * its neighbours in the existing ladder, so nothing already agreed moves.
     * It is a suggestion the user then confirms, never a silent decision.
     */
    function suggestGradeScore(jl) {
      var base = parseFloat(jl);
      var entries = Object.keys(cfg.gradeMap).map(function (k) {
        return { key: k, rank: parseFloat(k), score: cfg.gradeMap[k], high: /H$/i.test(k) };
      }).filter(function (e) { return isFinite(e.rank); })
        .sort(function (a, b) { return a.score - b.score; });
      if (!entries.length || !isFinite(base)) return 100;

      var isHigh = /H$/i.test(jl);
      var lower = null, upper = null;
      entries.forEach(function (e) {
        var r = e.rank + (e.high ? 0.5 : 0);
        var mine = base + (isHigh ? 0.5 : 0);
        if (r < mine && (!lower || e.score > lower.score)) lower = e;
        if (r > mine && (!upper || e.score < upper.score)) upper = e;
      });
      if (lower && upper) return Math.round((lower.score + upper.score) / 2 / 5) * 5;
      if (lower) return lower.score + 50;
      if (upper) return Math.max(0, upper.score - 50);
      return 100;
    }

    renderGradeTable();
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

    /* -- delivery recipients ------------------------------------------- */
    var m = mailConfig();
    var mailBody = el('div', {});

    function addressRow(who, key, hint) {
      var inp = el('input', { type: 'text', class: 'editable', placeholder: 'name@mtnirancell.ir' });
      inp.value = m[key] || '';
      var note = el('span', { class: 'small' });
      function validate() {
        var bad = invalidAddresses(inp.value);
        U.clear(note);
        if (!inp.value.trim()) {
          note.appendChild(el('span', { class: 'chip', text: 'ثبت نشده' }));
        } else if (bad.length) {
          note.appendChild(el('span', { class: 'chip err', text: 'نامعتبر: ' + bad.join('، ') }));
        } else {
          note.appendChild(el('span', { class: 'chip ok',
            text: parseAddresses(inp.value).length + ' نشانی' }));
        }
        inp.classList.toggle('invalid', bad.length > 0);
      }
      validate();
      inp.addEventListener('input', validate);
      inp.addEventListener('change', function () {
        if (m[key] === inp.value.trim()) return;
        auditConfig('mail.' + key, m[key], inp.value.trim());
        m[key] = inp.value.trim();
        save();
      });
      mailBody.appendChild(el('div', { class: 'mailrow' }, [
        el('span', { class: 'who', text: who }), inp, note
      ]));
      if (hint) mailBody.appendChild(el('div', { class: 'small muted', style: 'margin:-4px 0 10px 158px' },
        [document.createTextNode(hint)]));
    }

    addressRow('تیم عملکرد', 'performance', 'نتایج ارزیابی — بدون هیچ مبلغی');
    addressRow('تیم جبران خدمات', 'compensation', 'مبالغ نهایی در قالب فایل حقوق و دستمزد');
    addressRow('رونوشت (اختیاری)', 'cc', '');

    mailBody.appendChild(el('div', { style: 'margin-top:10px;display:flex;gap:8px;flex-wrap:wrap' }, [
      btn('پیش‌نمایش ارسال', function () {
        if (!App.result || !App.result.totals.inScopeCount) {
          U.toast('ابتدا پرسشنامه‌ها را وارد کنید.', 'warn'); return;
        }
        sendFinalPackage();
      }, 'primary')
    ]));
    mailBody.appendChild(el('div', { class: 'small muted', style: 'margin-top:9px' },
      [document.createTextNode(
        'صفحه در مرورگر اجرا می‌شود و خودش ایمیل نمی‌فرستد: فایل‌ها را می‌سازد و ' +
        'پیش‌نویس آدرس‌دهی‌شده را باز می‌کند تا پیوست کنید.')]));

    main.appendChild(U.card('گیرندگان ارسال نهایی', mailBody));

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

  /** Shallow copy of a config with a few keys overridden, for previews. */
  function mergeCfg(base, overrides) {
    var out = {}, k;
    for (k in base) if (base.hasOwnProperty(k)) out[k] = base[k];
    for (k in overrides) if (overrides.hasOwnProperty(k)) out[k] = overrides[k];
    return out;
  }

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
          if (App.state.config[k] === undefined) return;
          /* Merge the grade table rather than replacing it: the sample carries
             the workbook's six levels, but an organisation may have added its
             own (2H, for one) and loading a demo must not delete them. */
          if (k === 'gradeMap') {
            Object.keys(s.config.gradeMap).forEach(function (jl) {
              App.state.config.gradeMap[jl] = s.config.gradeMap[jl];
            });
            return;
          }
          App.state.config[k] = s.config[k];
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
      'دریافت فایل‌های خروجی و ارسال به تیم‌های عملکرد و جبران خدمات.'));

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

    /* Delivery first: this is where the period ends and the files go out. */
    var mail = mailConfig();
    var deliverBody = el('div', {}, [
      el('div', { style: 'display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px' }, [
        btn('دریافت و ارسال بستهٔ نهایی', function () { sendFinalPackage(); }, 'primary'),
        btn('فقط فایل عملکرد', function () {
          writeWorkbook(buildPerformanceWorkbook(),
            'Karaneh-Performance-' + periodSlug() + '-' + stamp() + '.xlsx');
          U.toast('فایل عملکرد تولید شد.', 'ok');
        }),
        btn('فقط فایل مبالغ', function () {
          writeWorkbook(buildCompensationWorkbook(),
            'Karaneh-Compensation-' + periodSlug() + '-' + stamp() + '.xlsx');
          U.toast('فایل مبالغ تولید شد.', 'ok');
        }),
        btn('گیرندگان', function () { go('settings'); }, 'ghost')
      ]),
      el('div', { class: 'small' }, [
        el('span', { class: 'chip ' + (parseAddresses(mail.performance).length ? 'ok' : 'err'),
          text: 'تیم عملکرد: ' + (parseAddresses(mail.performance).join('، ') || 'ثبت نشده') }),
        document.createTextNode('  '),
        el('span', { class: 'chip ' + (parseAddresses(mail.compensation).length ? 'ok' : 'err'),
          text: 'تیم جبران خدمات: ' + (parseAddresses(mail.compensation).join('، ') || 'ثبت نشده') })
      ])
    ]);
    main.appendChild(U.card('تحویل نهایی', deliverBody,
      { hint: 'فایل عملکرد بدون مبلغ · فایل مبالغ در قالب حقوق و دستمزد' }));

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
      btn('خروجی کامل Excel', function () { exportWorkbook(); }, 'primary'),
      btn('خروجی CSV', function () { exportCsv(); }),
      btn('چاپ', function () { window.print(); })
    ]));
    main.appendChild(U.card('گزارش تحلیلی', list));

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
    }).sort(function (a, b) {
      /* Highest job level first, then by employee number — the order a
         division head wants to review in. */
      var byLevel = U.naturalCompare(b.jobLevel || '', a.jobLevel || '');
      return byLevel || U.naturalCompare(a.employeeId, b.employeeId);
    });
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
    if (!opts.quiet) U.toast('تمپلیت با ' + roster.length + ' نفر تولید شد.', 'ok');
  }

  var GROUPINGS = {
    jobLevel: { label: 'سطح شغلی', prefix: 'JL-' },
    division: { label: 'واحد سازمانی', prefix: '' }
  };

  /**
   * One template per group. Job level is the default because a division head
   * reviews their people band by band — a level-4 conversation is a different
   * conversation from a level-3 one — so the file arrives already sorted that
   * way rather than mixed together by unit.
   */
  function downloadQuestionnaireTemplateByGroup(field) {
    field = field || 'jobLevel';
    var meta = GROUPINGS[field] || GROUPINGS.jobLevel;
    var groups = {};
    templateRoster().forEach(function (e) {
      var k = e[field] || 'نامشخص';
      (groups[k] || (groups[k] = [])).push(e);
    });
    var names = Object.keys(groups).sort(function (a, b) {
      return field === 'jobLevel' ? U.naturalCompare(b, a) : a.localeCompare(b, 'fa');
    });
    if (!names.length) { U.toast('فهرست پرسنلی برای تفکیک وجود ندارد.', 'warn'); return; }

    var chosen = {};
    var body = el('div', {});
    names.forEach(function (d) {
      var cb = el('input', { type: 'checkbox' });
      cb.checked = true; chosen[d] = true;
      cb.addEventListener('change', function () { chosen[d] = cb.checked; });
      body.appendChild(el('label', { class: 'checkline' }, [
        cb, el('span', { text: meta.label + ' ' + d + ' — ' + groups[d].length + ' نفر' })
      ]));
    });

    U.modal({
      title: 'تمپلیت به تفکیک ' + meta.label, size: 'narrow', content: body,
      buttons: [
        { label: 'تولید فایل‌ها', kind: 'primary', onClick: function () {
          var picked = names.filter(function (d) { return chosen[d]; });
          if (!picked.length) { U.toast('هیچ گروهی انتخاب نشد.', 'warn'); return; }
          picked.forEach(function (d, i) {
            /* Stagger the saves: browsers drop bursts of simultaneous downloads. */
            setTimeout(function () {
              downloadQuestionnaireTemplate({
                filter: function (e) { return (e[field] || 'نامشخص') === d; },
                scopeLabel: meta.label + ' ' + d,
                suffix: safeFileNameOr(meta.prefix + d, 'group'),
                quiet: i < picked.length - 1
              });
            }, i * 500);
          });
          if (picked.length > 1) {
            U.toast(picked.length + ' فایل در حال تولید است. اگر مرورگر اجازه خواست، ' +
              'دانلود چندگانه را تأیید کنید.', 'warn', 8000);
          }
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

  /* Persian letters, mapped to the closest Latin spelling. Browsers drop a
     `download` filename that contains non-ASCII — Chromium saves it as
     "download" with no extension, which looks exactly like a broken button —
     so every generated name is transliterated before it is used. */
  var TRANSLIT = {
    'ا': 'a', 'آ': 'a', 'أ': 'a', 'إ': 'e', 'ب': 'b', 'پ': 'p', 'ت': 't', 'ث': 's',
    'ج': 'j', 'چ': 'ch', 'ح': 'h', 'خ': 'kh', 'د': 'd', 'ذ': 'z', 'ر': 'r', 'ز': 'z',
    'ژ': 'zh', 'س': 's', 'ش': 'sh', 'ص': 's', 'ض': 'z', 'ط': 't', 'ظ': 'z', 'ع': 'a',
    'غ': 'gh', 'ف': 'f', 'ق': 'gh', 'ک': 'k', 'ك': 'k', 'گ': 'g', 'ل': 'l', 'م': 'm',
    'ن': 'n', 'و': 'v', 'ه': 'h', 'ة': 'h', 'ی': 'y', 'ي': 'y', 'ئ': 'y', 'ؤ': 'v',
    '\u200c': '-', '\u064b': '', '\u064c': '', '\u064d': '', '\u064e': '',
    '\u064f': '', '\u0650': '', '\u0651': '', '\u0652': '',
    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
    '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9',
    '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
    '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'
  };

  /**
   * An ASCII filename a browser will actually honour.
   * Returns '' when nothing survives, so callers can fall back to a label of
   * their own rather than producing a file called "-".
   */
  function safeFileName(s) {
    var out = String(s === null || s === undefined ? '' : s)
      .split('').map(function (ch) {
        if (TRANSLIT[ch] !== undefined) return TRANSLIT[ch];
        return /[A-Za-z0-9]/.test(ch) ? ch : '-';
      }).join('');
    return out.replace(/-+/g, '-').replace(/^-|-$/g, '').slice(0, 48);
  }

  /** Same, but never empty — used where a name is required. */
  function safeFileNameOr(s, fallback) {
    return safeFileName(s) || safeFileName(fallback) || 'file';
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
  function writeWorkbook(wb, filename, label) {
    var spec = {};
    wb.SheetNames.forEach(function (n) {
      if (wb.Sheets[n]['!postprocess']) spec[n] = wb.Sheets[n]['!postprocess'];
    });
    var raw = XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: false });
    var bytes = window.XlsxPostprocess
      ? window.XlsxPostprocess.applyFormatting(raw, spec)
      : new Uint8Array(raw);
    download(bytes, filename,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', label);
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
    var name = 'Karaneh-Payroll-' + periodSlug() + '-' + stamp() + '.xlsx';
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
      { key: 'jobLevel', label: 'سطح شغلی' },
      { key: 'directManager', label: 'مدیر مستقیم' },
      { key: 'managerLevel1', label: 'مدیر سطح ۱' },
      { key: 'managerLevel2', label: 'مدیر سطح ۲' },
      { key: 'managerLevel3', label: 'مدیر سطح ۳' }
    ];
    var chosen = 'jobLevel';
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
      title: 'خروجی به تفکیک گروه', size: 'narrow', content: body,
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
          writeWorkbook(wb, 'Karaneh-' + safeFileNameOr(n, 'group') + '-' + stamp() + '.xlsx', n);
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
      writeWorkbook(wb, 'Karaneh-By-' + safeFileNameOr(label, 'group') + '-' + stamp() + '.xlsx');
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
  /* Test hook: the roster a questionnaire template is built from. */
  window.__templateRoster = function (filter) { return templateRoster(filter); };

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

  /* ======================================================================
   * Handover — building a division head's own copy of the system
   * ----------------------------------------------------------------------
   * HR keeps one file. Each division head gets their own: the same
   * application, carrying only their people, opening in their role, and
   * remembering their work separately.
   *
   * The page is a single self-contained file, so it can produce that copy by
   * cloning its own document and injecting a payload script ahead of the
   * application code. No server and no build step involved.
   * ---------------------------------------------------------------------- */

  var PACKAGE_MARK = 'karaneh-package-payload';

  /** Which employees and answers belong to one handover. */
  function packageSlice(field, value) {
    var employees = App.state.employees.filter(function (e) {
      return String(e[field] || '') === String(value);
    });
    var ids = {};
    employees.forEach(function (e) { ids[e.employeeId] = 1; });
    var questionnaires = App.state.questionnaires.filter(function (q) {
      return ids[q.employeeId];
    });
    var divisions = {};
    employees.forEach(function (e) { if (e.division) divisions[e.division] = 1; });
    return {
      employees: employees,
      questionnaires: questionnaires,
      divisions: Object.keys(divisions)
    };
  }

  /**
   * Clone the running page into a standalone file.
   *
   * `document.documentElement.outerHTML` is the whole application, inline
   * scripts included. The clone is reset to its pre-boot state — the rendered
   * interface stripped out, any earlier payload removed — before the new
   * payload is inserted into <head>, ahead of the application script that
   * reads it.
   */
  function buildPackageHtml(payload) {
    var source = '<!DOCTYPE html>\n' + document.documentElement.outerHTML;
    var doc = new DOMParser().parseFromString(source, 'text/html');

    /* Reset anything the running session drew or attached. */
    var app = doc.getElementById('app');
    if (app) {
      app.innerHTML = '<div style="padding:40px;text-align:center;' +
        'font-family:Tahoma,sans-serif">در حال بارگذاری…</div>';
    }
    ['dlTray'].forEach(function (id) {
      var n = doc.getElementById(id);
      if (n && n.parentNode) n.parentNode.removeChild(n);
    });
    Array.prototype.slice.call(doc.querySelectorAll('.toast-host, .modal-backdrop'))
      .forEach(function (n) { n.parentNode.removeChild(n); });
    doc.documentElement.setAttribute('data-theme', payload.theme || 'light');

    /* A file generated from a generated file must not stack payloads. */
    Array.prototype.slice.call(doc.querySelectorAll('script[data-' + PACKAGE_MARK + ']'))
      .forEach(function (n) { n.parentNode.removeChild(n); });

    var script = doc.createElement('script');
    script.setAttribute('data-' + PACKAGE_MARK, '1');
    /* JSON is embedded rather than assigned as a literal so that any "</script>"
       inside the data cannot close the tag early. */
    script.textContent = 'window.__KARANEH_PACKAGE__ = JSON.parse(' +
      JSON.stringify(JSON.stringify(payload)) + ');';
    doc.head.appendChild(script);

    var title = doc.querySelector('title');
    if (title) title.textContent = 'کارانه — ' + payload.package.label;

    return '<!DOCTYPE html>\n' + doc.documentElement.outerHTML;
  }

  /* Exposed for the end-to-end test, which builds a package without going
     through the modal. */
  window.__buildPkg = function (label, slice, scope) {
    return buildPackageHtml(packagePayload(label, slice, scope));
  };

  /**
   * The share of HR's own allocation that this group currently holds.
   *
   * Handing the head the organisation-wide budget would be misleading — ten
   * people would appear to be sharing a hundred billion rial. The group's
   * present allocation is a defensible opening number, and the head is free to
   * overwrite it on the payment screen.
   */
  function sliceBudget(slice) {
    var ids = {};
    slice.employees.forEach(function (e) { ids[e.employeeId] = 1; });
    var sum = 0;
    App.result.rows.forEach(function (r) {
      if (ids[r.employeeId] && r.inScope) sum += r.finalKaraneh || 0;
    });
    return Math.round(sum);
  }

  function packagePayload(label, slice, scope) {
    var config = JSON.parse(JSON.stringify(App.state.config));
    var share = sliceBudget(slice);
    if (share > 0) config.budget = share;
    return {
      package: {
        id: 'pkg-' + Date.now() + '-' + Math.random().toString(16).slice(2, 8),
        label: label,
        issuedAt: new Date().toISOString(),
        issuedBy: ROLES[role()].label,
        period: App.state.period
      },
      period: App.state.period,
      theme: App.state.theme || 'light',
      scope: scope || slice.divisions,
      config: config,
      budgetSource: share > 0 ? 'سهم این گروه از تخصیص منابع انسانی' : '',
      columnMappings: JSON.parse(JSON.stringify(App.state.columnMappings)),
      mail: JSON.parse(JSON.stringify(App.state.mail || {})),
      employees: JSON.parse(JSON.stringify(slice.employees)),
      questionnaires: JSON.parse(JSON.stringify(slice.questionnaires)),
      importBatches: []
    };
  }

  function emitPackage(label, slice, scope, quiet) {
    var payload = packagePayload(label, slice, scope);
    var html = buildPackageHtml(payload);
    var name = 'Karaneh-' + safeFileNameOr(label, 'group') + '-' + periodSlug() +
               '-' + stamp() + '.html';
    download(html, name, 'text/html;charset=utf-8', label);
    Store.audit(App.state, {
      entity: 'package', field: 'issue', oldValue: '',
      newValue: name + ' — ' + slice.employees.length + ' نفر',
      reason: 'تولید فایل معاون بخش: ' + label
    });
    if (!quiet) {
      U.toast('فایل «' + label + '» با ' + slice.employees.length + ' نفر تولید شد.', 'ok', 5000);
    }
    return { name: name, size: html.length, count: slice.employees.length };
  }

  /**
   * The HR-side screen for issuing handover files. HR picks how to carve up
   * the organisation, sees the resulting groups with their headcounts, and
   * gets one self-contained file per group.
   */
  function openPackageBuilder() {
    if (!App.state.employees.length) {
      U.toast('ابتدا اطلاعات پرسنل را وارد کنید.', 'warn');
      return;
    }

    var fields = [
      { key: 'division',      label: 'واحد سازمانی' },
      { key: 'directManager', label: 'مدیر مستقیم' },
      { key: 'managerLevel1', label: 'مدیر سطح ۱' },
      { key: 'managerLevel2', label: 'مدیر سطح ۲' },
      { key: 'managerLevel3', label: 'مدیر سطح ۳' },
      { key: 'jobLevel',      label: 'سطح شغلی' }
    ];
    var field = 'division';
    var chosen = {};

    var listBox = el('div', { style: 'max-height:300px;overflow-y:auto;margin-top:10px' });
    var summary = el('div', { class: 'small muted', style: 'margin-top:8px' });

    function groupsFor(f) {
      var g = {};
      App.state.employees.forEach(function (e) {
        if (!isPayrollEligible(e)) return;
        var k = e[f];
        if (!k) return;
        g[k] = (g[k] || 0) + 1;
      });
      return g;
    }

    function renderGroups() {
      var g = groupsFor(field);
      var names = Object.keys(g).sort(function (a, b) {
        return field === 'jobLevel' ? U.naturalCompare(b, a) : a.localeCompare(b, 'fa');
      });
      chosen = {};
      U.clear(listBox);
      if (!names.length) {
        listBox.appendChild(el('div', { class: 'small muted',
          text: 'برای این تفکیک، مقداری در اطلاعات پرسنل ثبت نشده است.' }));
        U.clear(summary);
        return;
      }
      names.forEach(function (n) {
        chosen[n] = true;
        var cb = el('input', { type: 'checkbox' });
        cb.checked = true;
        cb.addEventListener('change', function () { chosen[n] = cb.checked; updateSummary(); });
        listBox.appendChild(el('label', { class: 'checkline' }, [
          cb, el('span', { text: n + ' — ' + g[n] + ' نفر' })
        ]));
      });
      updateSummary();
    }

    function updateSummary() {
      var g = groupsFor(field);
      var picked = Object.keys(chosen).filter(function (k) { return chosen[k]; });
      var people = picked.reduce(function (a, k) { return a + (g[k] || 0); }, 0);
      U.clear(summary);
      summary.appendChild(el('b', { text: picked.length + ' فایل' }));
      summary.appendChild(document.createTextNode(' • ' + people + ' نفر در مجموع'));
    }

    var sel = el('select', { class: 'editable', style: 'width:100%' });
    fields.forEach(function (f) { sel.appendChild(el('option', { value: f.key, text: f.label })); });
    sel.addEventListener('change', function () { field = sel.value; renderGroups(); });

    var body = el('div', {}, [
      el('label', { class: 'field' }, [
        el('span', { text: 'تفکیک بر اساس' }), sel
      ]),
      listBox,
      summary,
      el('div', { class: 'small muted', style: 'margin-top:10px' },
        [document.createTextNode(
          'هر فایل یک نسخهٔ کامل و مستقل از سامانه است که فقط پرسنل همان گروه را دارد ' +
          'و در نقش «معاون بخش» باز می‌شود.')])
    ]);
    renderGroups();

    U.modal({
      title: 'تولید فایل معاون بخش', size: 'narrow', content: body,
      buttons: [
        { label: 'تولید فایل‌ها', kind: 'primary', onClick: function () {
          var picked = Object.keys(chosen).filter(function (k) { return chosen[k]; });
          if (!picked.length) { U.toast('هیچ گروهی انتخاب نشد.', 'warn'); return; }
          picked.forEach(function (name, i) {
            /* Stagger the saves: browsers drop bursts of simultaneous downloads. */
            setTimeout(function () {
              var slice = packageSlice(field, name);
              var scope = field === 'division' ? [name] : slice.divisions;
              emitPackage(name, slice, scope, i < picked.length - 1);
            }, i * 550);
          });
          save();
          if (picked.length > 1) {
            U.toast(picked.length + ' فایل در حال تولید است. اگر مرورگر اجازه خواست، ' +
              'دانلود چندگانه را تأیید کنید — همه در صفحهٔ دانلود هم می‌مانند.', 'warn', 9000);
          }
        } },
        { label: 'انصراف' }
      ]
    });
  }

  /* ======================================================================
   * Delivery — the two files a finalised period produces, and the handover
   * ----------------------------------------------------------------------
   * A division head finishes; two different teams need two different things.
   * The performance team needs the ratings and never the money; C&B needs the
   * amounts in payroll's own layout. Splitting them here means neither team
   * has to be sent data it should not hold.
   *
   * The page runs from the browser with no server, so it cannot put a file
   * into an email itself. What it does instead: build both files, hand them
   * over, and open an addressed draft naming exactly what to attach.
   * ---------------------------------------------------------------------- */

  function mailConfig() {
    var m = App.state.mail || (App.state.mail = {});
    if (m.performance === undefined) m.performance = '';
    if (m.compensation === undefined) m.compensation = '';
    if (m.cc === undefined) m.cc = '';
    return m;
  }

  /** Split and tidy a comma or semicolon separated address list. */
  function parseAddresses(text) {
    return String(text || '').split(/[,;،\s]+/)
      .map(function (a) { return a.trim(); })
      .filter(function (a) { return a; });
  }

  var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

  function invalidAddresses(text) {
    return parseAddresses(text).filter(function (a) { return !EMAIL_RE.test(a); });
  }

  /**
   * The performance package: ratings, scores and the behavioural anchor each
   * answer corresponds to. Deliberately carries no rial figure.
   */
  function buildPerformanceWorkbook() {
    var cfg = App.state.config;
    var questions = cfg.questions.filter(function (q) { return !q.impact; });
    var options = Object.keys(cfg.answerScale);

    var header = ['شماره پرسنلی', 'نام و نام خانوادگی', 'واحد سازمانی', 'عنوان شغلی', 'سطح شغلی'];
    questions.forEach(function (q) {
      header.push((q.domain || q.id.toUpperCase()) + ' — پاسخ');
      header.push((q.domain || q.id.toUpperCase()) + ' — سطح');
    });
    header.push('امتیاز عملکرد', 'عدد کارانه', 'اثرگذاری ویژه', 'امتیاز ویژه',
                'ضریب نهایی کارانه', 'وضعیت', 'توضیح معاون بخش');

    var aoa = [header];
    App.result.rows.forEach(function (r) {
      if (!r.inScope || !inScopeForRole(r)) return;
      var row = [r.employeeId, r.fullName, r.division, r.positionTitle, r.jobLevel];
      questions.forEach(function (q) {
        var answer = r[q.id];
        var idx = options.indexOf(String(answer).trim());
        var anchor = (q.anchors || [])[idx];
        row.push(answer === null || answer === undefined ? '' : answer);
        row.push(anchor ? anchor.label : '');
      });
      row.push(r.performanceScore === null ? '' : r.performanceScore,
               r.performanceKaraneh === null ? '' : r.performanceKaraneh,
               r.specialProject ? 'بله' : 'خیر',
               r.specialImpactValue || '',
               r.finalCoefficient === null ? '' : r.finalCoefficient,
               statusLabel(r.status),
               r.hodComment || '');
      aoa.push(row);
    });

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 13 }, { wch: 22 }, { wch: 16 }, { wch: 24 }, { wch: 8 }]
      .concat(questions.reduce(function (a) { return a.concat([{ wch: 14 }, { wch: 16 }]); }, []))
      .concat([{ wch: 12 }, { wch: 12 }, { wch: 13 }, { wch: 11 }, { wch: 14 }, { wch: 14 }, { wch: 32 }]);
    var scoreStart = 5 + questions.length * 2;
    ws['!postprocess'] = {
      xSplit: 2, ySplit: 1, headerRow: 1,
      numberFormats: (function () {
        var f = {};
        [scoreStart, scoreStart + 1, scoreStart + 3, scoreStart + 4].forEach(function (c) {
          f[colLetter(c)] = '0.00';
        });
        return f;
      }())
    };

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Performance');
    XLSX.utils.book_append_sheet(wb, Tpl.buildBarsSheet(cfg, XLSX), 'BARS');
    wb.SheetNames.forEach(function (n) { wb.Sheets[n]['!rtl'] = true; });
    wb.Workbook = { Views: [{ RTL: true }] };
    return wb;
  }

  /** The compensation package: payroll's own layout with the amounts filled. */
  function buildCompensationWorkbook() {
    var rows = payrollRoster();
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, payrollSheet(rows).ws, 'Karaneh');
    wb.Workbook = { Views: [{ RTL: false }] };
    return wb;
  }

  function periodSlug() { return safeFileNameOr(App.state.period, 'period'); }

  /**
   * Produce both files and open the two drafts. Returns what was produced so
   * the caller can report it rather than guessing.
   */
  function sendFinalPackage(opts) {
    opts = opts || {};
    var m = mailConfig();
    var stampNow = stamp();
    var perfName = 'Karaneh-Performance-' + periodSlug() + '-' + stampNow + '.xlsx';
    var compName = 'Karaneh-Compensation-' + periodSlug() + '-' + stampNow + '.xlsx';

    writeWorkbook(buildPerformanceWorkbook(), perfName);
    setTimeout(function () { writeWorkbook(buildCompensationWorkbook(), compName); }, 500);

    var scope = isAdmin() ? 'کل سازمان' : (roleScope() ? roleScope().join('، ') : 'کل سازمان');
    var count = App.result.rows.filter(function (r) { return r.inScope && inScopeForRole(r); }).length;

    var drafts = [
      {
        to: m.performance, file: perfName, team: 'تیم عملکرد',
        subject: 'کارانه ' + App.state.period + ' — نتایج ارزیابی عملکرد (' + scope + ')',
        body: 'با سلام،\n\nنتایج ارزیابی عملکرد دورهٔ ' + App.state.period +
              ' برای ' + count + ' نفر (' + scope + ') نهایی شد.\n\n' +
              'فایل پیوست: ' + perfName + '\n\nبا احترام'
      },
      {
        to: m.compensation, file: compName, team: 'تیم جبران خدمات',
        subject: 'کارانه ' + App.state.period + ' — مبالغ نهایی (' + scope + ')',
        body: 'با سلام،\n\nمبالغ نهایی کارانهٔ دورهٔ ' + App.state.period +
              ' برای ' + count + ' نفر (' + scope + ') نهایی شد.\n' +
              'مجموع پرداخت: ' + U.money(App.result.totals.sumFinalKaraneh) + ' ریال\n\n' +
              'فایل پیوست: ' + compName + '\n\nبا احترام'
      }
    ];

    Store.audit(App.state, {
      entity: 'delivery', field: 'send', oldValue: '',
      newValue: perfName + ' → ' + (m.performance || '—') + ' | ' +
                compName + ' → ' + (m.compensation || '—'),
      reason: 'آماده‌سازی ارسال بستهٔ نهایی دوره ' + App.state.period
    });
    save();

    if (opts.silent) return drafts;
    showDeliveryPanel(drafts, m);
    return drafts;
  }

  function mailtoUrl(draft, cc) {
    var parts = [];
    parts.push('subject=' + encodeURIComponent(draft.subject));
    parts.push('body=' + encodeURIComponent(draft.body));
    if (cc) parts.push('cc=' + encodeURIComponent(parseAddresses(cc).join(',')));
    return 'mailto:' + encodeURIComponent(parseAddresses(draft.to).join(',')).replace(/%40/g, '@') +
           '?' + parts.join('&');
  }

  function showDeliveryPanel(drafts, m) {
    var body = el('div', {});
    body.appendChild(el('p', { class: 'small muted', style: 'margin-top:0' },
      [document.createTextNode(
        'هر دو فایل تولید و در صفحهٔ دانلود قرار گرفتند. با کلیک روی «بازکردن ایمیل»، ' +
        'پیش‌نویس با گیرنده و موضوع آماده می‌شود؛ فایل را از صفحهٔ دانلود پیوست کنید.')]));

    drafts.forEach(function (d) {
      var missing = !parseAddresses(d.to).length;
      body.appendChild(el('div', {
        style: 'border:1px solid var(--border);border-radius:9px;padding:11px 13px;margin-bottom:10px'
      }, [
        el('div', { style: 'display:flex;align-items:center;gap:9px;flex-wrap:wrap' }, [
          el('b', { text: d.team }),
          missing
            ? el('span', { class: 'chip err', text: 'نشانی ثبت نشده' })
            : el('span', { class: 'chip ok mono', text: parseAddresses(d.to).join('، ') })
        ]),
        el('div', { class: 'small muted mono', style: 'margin-top:5px', text: d.file }),
        el('div', { style: 'margin-top:9px;display:flex;gap:7px' }, [
          el('a', {
            class: 'btn sm primary' + (missing ? ' disabled' : ''),
            href: missing ? '#' : mailtoUrl(d, m.cc),
            text: 'بازکردن ایمیل',
            onclick: function (e) {
              if (missing) { e.preventDefault(); go('settings'); }
            }
          }),
          missing ? btn('ثبت نشانی', function () { go('settings'); }, 'sm') : null
        ].filter(Boolean))
      ]));
    });

    U.modal({
      title: 'ارسال بستهٔ نهایی', content: body,
      buttons: [
        { label: 'نمایش فایل‌ها', onClick: function () { showDownloadTray(); } },
        'spacer',
        { label: 'بستن', kind: 'primary' }
      ]
    });
  }

  function exportWorkbook() {
    try {
      var name = 'Karaneh-' + periodSlug() + '-' + stamp() + '.xlsx';
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
    writeWorkbook(single, safeFileNameOr(keep, 'sheet') + '-' + stamp() + '.xlsx', keep);
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

  /**
   * Hand a file to the user.
   *
   * A programmatic click is the normal path, but it is not reliable
   * everywhere: some browsers refuse synthetic downloads from a file:// page,
   * and every browser blocks the second and later files of a batch until the
   * user grants permission. Silently failing there looks exactly like a broken
   * button, so the last file handed out is also kept as a real link the user
   * can click, and the panel that shows it stays until dismissed.
   *
   * Accepts a string or a Uint8Array; Blob handles both.
   */
  function download(content, filename, type, label) {
    var blob = new Blob([content], { type: type });
    var url = URL.createObjectURL(blob);

    var a = el('a', { href: url, download: filename });
    a.style.display = 'none';
    document.body.appendChild(a);
    try { a.click(); } catch (e) { /* fall through to the manual link */ }
    setTimeout(function () { if (a.parentNode) a.parentNode.removeChild(a); }, 1000);

    rememberDownload(filename, url, type, label);
  }

  /* The files produced in this session, newest first, each still clickable. */
  App.downloads = [];

  function rememberDownload(filename, url, type, label) {
    App.downloads.unshift({
      filename: filename, url: url, type: type, label: label || '', at: new Date()
    });
    /* Object URLs hold the blob in memory; keep a bounded number alive. */
    while (App.downloads.length > 12) {
      var old = App.downloads.pop();
      try { URL.revokeObjectURL(old.url); } catch (e) { /* already gone */ }
    }
    showDownloadTray();
  }

  /**
   * A persistent tray listing what this session produced. It is the answer to
   * "the download button does nothing": whatever the browser did, the file is
   * here and one click away.
   */
  function showDownloadTray() {
    var host = document.getElementById('dlTray');
    if (!host) {
      host = el('div', { id: 'dlTray', class: 'dl-tray' });
      document.body.appendChild(host);
    }
    U.clear(host);

    host.appendChild(el('div', { class: 'dl-head' }, [
      el('b', { text: 'فایل‌های آمادهٔ دریافت' }),
      el('span', { class: 'muted small', text: ' (' + App.downloads.length + ')' }),
      el('div', { style: 'flex:1' }),
      el('button', {
        class: 'btn sm ghost', text: '✕',
        title: 'بستن',
        onclick: function () { if (host.parentNode) host.parentNode.removeChild(host); }
      })
    ]));

    App.downloads.forEach(function (d) {
      host.appendChild(el('a', {
        class: 'dl-item', href: d.url, download: d.filename,
        title: (d.label ? d.label + '\n' : '') + d.filename
      }, [
        el('span', { class: 'ico', text: '⬇' }),
        el('span', {}, [
          d.label ? el('div', { class: 'label', text: d.label }) : null,
          el('div', { class: 'name', text: d.filename })
        ].filter(Boolean))
      ]));
    });

    host.appendChild(el('div', { class: 'dl-note',
      text: 'اگر دانلود خودکار شروع نشد، روی نام فایل کلیک کنید.' }));
  }
  App.showDownloadTray = showDownloadTray;

  function stamp() {
    var d = new Date();
    function p(n) { return n < 10 ? '0' + n : String(n); }
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '-' + p(d.getHours()) + p(d.getMinutes());
  }

  /* ====================================================================== */
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
}());
