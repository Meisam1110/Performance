/* ============================================================================
 * templates.js — downloadable Excel templates and the checks that pair with them
 * ----------------------------------------------------------------------------
 * Two templates ship from here:
 *
 *   1. «پرسشنامه کارانه تیمی» — generated from whatever the questionnaire
 *      designer currently holds. Question text, order, weights and the answer
 *      scale all come from config, so redesigning the instrument reissues the
 *      template with no code change.
 *
 *   2. Employee master — the column layout the payroll team's own file uses,
 *      so their export drops straight in.
 *
 * Every generated file carries a hidden `_Template` sheet describing the
 * design it was cut from. On import that signature is compared against the
 * current design, which is what lets the system say "this file does not match
 * the template" instead of silently importing the wrong shape.
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Templates = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var META_SHEET = '_Template';
  var QUESTIONNAIRE_SHEET = 'پرسشنامه کارانه تیمی';
  var TEMPLATE_VERSION = 1;

  /* Identity columns every questionnaire template opens with. */
  var IDENTITY = [
    { key: 'employeeId',    label: 'شماره پرسنلی',      width: 14, locked: true },
    { key: 'fullName',      label: 'نام و نام خانوادگی', width: 22, locked: true },
    { key: 'division',      label: 'واحد سازمانی',      width: 16, locked: true },
    { key: 'positionTitle', label: 'عنوان شغلی',        width: 24, locked: true },
    { key: 'jobLevel',      label: 'JL',                width: 7,  locked: true }
  ];

  /* The payroll team's file, column for column. */
  var PAYROLL_COLUMNS = [
    { key: 'employeeId',       label: 'Emp No',                   width: 13 },
    { key: 'employeeStatus',   label: 'Emp Status',               width: 12 },
    { key: 'firstName',        label: 'First Name',               width: 14 },
    { key: 'lastName',         label: 'Last Name',                width: 18 },
    { key: 'dateOfEmployment', label: 'Date of Employment',       width: 18 },
    { key: 'dateOfLeaving',    label: 'Date of Leaving',          width: 18 },
    { key: 'nationalId',       label: 'National ID',              width: 14 },
    { key: 'gender',           label: 'Gender',                   width: 9 },
    { key: 'positionTitle',    label: 'Pos Title',                width: 28 },
    { key: 'assignmentType',   label: 'Assignment Type',          width: 15 },
    { key: 'employmentType',   label: 'Employment Type',          width: 18 },
    { key: 'jobLevel',         label: 'Job Level',                width: 10 },
    { key: 'division',         label: 'Division Alias',           width: 14 },
    { key: 'workingDays',      label: "Working Day's Q1",         width: 15 },
    { key: 'probationStatus',  label: 'Probation Checking',       width: 17 },
    { key: 'checking',         label: 'Checking',                 width: 14 },
    { key: 'comment',          label: 'Comment',                  width: 24 },
    { key: 'finalKaraneh',     label: 'Final Karaneh',            width: 20, output: true },
    { key: 'directManager',    label: 'Direct Manager',           width: 20 },
    { key: 'managerLevel1',    label: 'Manager Level 1',          width: 20 },
    { key: 'managerLevel2',    label: 'Manager Level 2',          width: 20 },
    { key: 'managerLevel3',    label: 'Manager Level 3',          width: 20 }
  ];

  /* ------------------------------------------------------------------------
   * Template signature
   * ---------------------------------------------------------------------- */

  /** The question that drives the special-impact amount, if one is flagged. */
  function impactQuestion(config) {
    return (config.questions || []).filter(function (q) { return q.impact; })[0] || null;
  }

  /** Questions that appear as answer columns on the template. */
  function answerQuestions(config) {
    return (config.questions || []).filter(function (q) { return !q.impact; });
  }

  /** Stable, order-sensitive fingerprint of the questionnaire design. */
  function describe(config) {
    var questions = (config.questions || []).map(function (q) {
      return q.id + ':' + (q.scored === false ? 0 : 1) + ':' + (q.weight === undefined ? 1 : q.weight);
    });
    return {
      version: TEMPLATE_VERSION,
      questions: questions.join('|'),
      questionIds: (config.questions || []).map(function (q) { return q.id; }).join('|'),
      options: Object.keys(config.answerScale || {}).join('|'),
      specialImpactMinScore: String(config.specialImpactMinScore === undefined ? '' : config.specialImpactMinScore)
    };
  }

  /** Small non-cryptographic checksum; enough to notice an edited template. */
  function signature(config) {
    var d = describe(config);
    var text = [d.version, d.questions, d.options, d.specialImpactMinScore].join('#');
    var h = 5381;
    for (var i = 0; i < text.length; i++) h = ((h * 33) ^ text.charCodeAt(i)) >>> 0;
    return ('0000000' + h.toString(16)).slice(-8);
  }

  function metaSheetRows(config, period, kind) {
    var d = describe(config);
    return [
      ['KARANEH_TEMPLATE'],
      ['kind', kind],
      ['version', d.version],
      ['period', period || ''],
      ['generatedAt', new Date().toISOString()],
      ['questions', d.questions],
      ['questionIds', d.questionIds],
      ['options', d.options],
      ['specialImpactMinScore', d.specialImpactMinScore],
      ['signature', signature(config)],
      [],
      ['این شیت برای اعتبارسنجی فایل هنگام بازگشت به سامانه است. آن را حذف یا ویرایش نکنید.']
    ];
  }

  /** Read the signature back out of an uploaded workbook, if it carries one. */
  function readMeta(workbook, XLSX) {
    if (!workbook.Sheets[META_SHEET]) return null;
    var rows = XLSX.utils.sheet_to_json(workbook.Sheets[META_SHEET], { header: 1, defval: null });
    var meta = {};
    rows.forEach(function (r) {
      if (r && r.length >= 2 && r[0]) meta[String(r[0]).trim()] = r[1] === null ? '' : String(r[1]).trim();
    });
    return meta.signature ? meta : null;
  }

  /**
   * Compare an uploaded workbook against the current questionnaire design.
   *
   * @returns {{ok:boolean, level:'match'|'drift'|'mismatch'|'unsigned', problems:string[], meta:Object}}
   *
   * `mismatch` means the file was cut from a different question set and must
   * not be imported. `drift` means only weights or the gate moved — the
   * answers are still usable, so it is a warning. `unsigned` means the file
   * did not come from this system at all; column matching decides.
   */
  function verifyAgainstTemplate(workbook, config, XLSX) {
    var meta = readMeta(workbook, XLSX);
    var current = describe(config);
    if (!meta) {
      return { ok: true, level: 'unsigned', problems: [], meta: null };
    }
    if (meta.signature === signature(config)) {
      return { ok: true, level: 'match', problems: [], meta: meta };
    }

    var problems = [];
    if (meta.questionIds !== current.questionIds) {
      problems.push('مجموعه سؤالات فایل با طراحی فعلی پرسشنامه یکسان نیست.\n' +
        'فایل: ' + (meta.questionIds || '—') + '\n' +
        'طراحی فعلی: ' + current.questionIds);
    }
    if (meta.options !== current.options) {
      problems.push('گزینه‌های پاسخ فایل با طراحی فعلی یکسان نیست.\n' +
        'فایل: ' + (meta.options || '—') + '\n' +
        'طراحی فعلی: ' + current.options);
    }
    var structural = problems.length > 0;
    if (meta.questions !== current.questions && !structural) {
      problems.push('وزن سؤالات از زمان تولید این تمپلیت تغییر کرده است. ' +
        'پاسخ‌ها قابل استفاده‌اند اما امتیازها با وزن‌های جدید محاسبه می‌شوند.');
    }
    if (meta.specialImpactMinScore !== current.specialImpactMinScore && !structural) {
      problems.push('حد نصاب سؤال اثرگذاری ویژه از ' + (meta.specialImpactMinScore || '—') +
        ' به ' + current.specialImpactMinScore + ' تغییر کرده است.');
    }
    return {
      ok: !structural,
      level: structural ? 'mismatch' : 'drift',
      problems: problems,
      meta: meta
    };
  }

  /**
   * Column-level check for files that carry no signature: does the detected
   * mapping cover every scored question?
   */
  function verifyColumns(mapping, config) {
    var missing = [], questions = config.questions || [];
    questions.forEach(function (q) {
      if (q.scored === false) return;
      if (mapping[q.id] === undefined) missing.push(q.id);
    });
    var problems = [];
    if (missing.length) {
      problems.push('ستون این سؤالات در فایل پیدا نشد: ' + missing.join('، ') +
        ' — ساختار فایل با تمپلیت فعلی هماهنگ نیست.');
    }
    if (mapping.employeeId === undefined) {
      problems.push('ستون «شماره پرسنلی» در فایل پیدا نشد.');
    }
    return { ok: problems.length === 0, problems: problems, missing: missing };
  }

  /* ------------------------------------------------------------------------
   * Questionnaire template
   * ---------------------------------------------------------------------- */

  /**
   * @param config    the live configuration (questions, answer scale, gate)
   * @param employees rows to pre-fill; pass [] for a blank template
   * @param opts      { period, XLSX, scopeLabel }
   * @returns a SheetJS workbook, with `!postprocess` set on each sheet
   */
  function buildQuestionnaireTemplate(config, employees, opts) {
    opts = opts || {};
    var XLSX = opts.XLSX || (typeof window !== 'undefined' ? window.XLSX : null);
    var questions = config.questions || [];
    var options = Object.keys(config.answerScale || {});

    var impact = impactQuestion(config);
    var columns = IDENTITY.slice();
    answerQuestions(config).forEach(function (q) {
      columns.push({
        key: q.id,
        label: q.id.toUpperCase(),
        description: (q.domain ? q.domain + ' — ' : '') + q.text,
        width: 17,
        answer: true,
        scored: q.scored !== false,
        weight: q.weight === undefined ? 1 : q.weight
      });
    });
    columns.push({
      key: 'specialProject', label: 'اثرگذاری ویژه',
      description: (impact && impact.domain ? impact.domain + ' — ' : '') +
                   (impact ? impact.text : (config.specialImpactQuestion || '')),
      width: 16, choice: true
    });
    columns.push({
      key: 'specialImpactAmount', label: 'امتیاز اثرگذاری ویژه', width: 20, amount: true
    });
    columns.push({
      key: 'specialImpactComment', label: 'توضیح اثرگذاری ویژه', width: 34
    });

    var aoa = [];
    aoa.push(['پرسشنامه کارانه تیمی — ' + (opts.period || '')]);
    aoa.push([opts.scopeLabel ? 'دامنه: ' + opts.scopeLabel : '']);
    aoa.push(['پاسخ‌ها را از فهرست کشویی هر خانه انتخاب کنید. شرح رفتاری هر سطح در شیت BARS آمده است.']);
    aoa.push([]);
    /* Descriptive band: the full question text sits above the short code. */
    aoa.push(columns.map(function (c) { return c.description || ''; }));
    aoa.push(columns.map(function (c) { return c.label; }));

    employees.forEach(function (e) {
      aoa.push(columns.map(function (c) {
        if (c.answer || c.choice || c.key === 'specialImpactComment') return '';
        var v = e[c.key];
        return v === null || v === undefined ? '' : v;
      }));
    });
    /* Blank rows so a manager can add someone the master file missed. */
    for (var b = 0; b < 20; b++) aoa.push(columns.map(function () { return ''; }));

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = columns.map(function (c) { return { wch: c.width }; });

    var headerRow = 6;                       // 1-based row holding the codes
    var firstData = headerRow + 1;
    var lastData = headerRow + employees.length + 20;

    var validations = [];
    columns.forEach(function (c, i) {
      if (c.answer) validations.push({ col: i, options: options });
      if (c.choice) validations.push({ col: i, options: ['بله', 'خیر'] });
      if (c.amount) {
        /* The special score moves in fixed steps, so the template offers the
           permitted values rather than a free number the system would reject. */
        var step = config.specialImpactStep || 50;
        var max = config.specialImpactAmount || 300;
        var choices = [];
        for (var v = step; v <= max; v += step) choices.push(String(v));
        validations.push({ col: i, options: choices });
      }
    });

    ws['!postprocess'] = {
      xSplit: 2, ySplit: headerRow, headerRow: headerRow,
      numberFormats: {},
      validations: validations.map(function (v) {
        return { col: v.col, firstRow: firstData, lastRow: lastData, options: v.options };
      })
    };

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, QUESTIONNAIRE_SHEET);

    XLSX.utils.book_append_sheet(wb, buildBarsSheet(config, XLSX), 'BARS');

    var mws = XLSX.utils.aoa_to_sheet(metaSheetRows(config, opts.period, 'questionnaire'));
    mws['!cols'] = [{ wch: 24 }, { wch: 60 }];
    XLSX.utils.book_append_sheet(wb, mws, META_SHEET);

    wb.SheetNames.forEach(function (n) { wb.Sheets[n]['!rtl'] = true; });
    wb.Workbook = { Views: [{ RTL: true }], Sheets: [{}, {}, { Hidden: 1 }] };
    return wb;
  }

  /**
   * The rating scale, on one sheet.
   *
   * One row per question, one column per level, each cell carrying the anchor
   * label and the behaviour it describes. This is what a rater reads while
   * filling the questionnaire, so it ships inside the same workbook rather
   * than as a separate document.
   */
  function buildBarsSheet(config, XLSX) {
    var levels = Object.keys(config.answerScale || {});
    var aoa = [];
    aoa.push(['مقیاس رفتاری ارزیابی — BARS']);
    aoa.push(['برای هر حوزه، رفتاری را انتخاب کنید که بیشترین شباهت را به عملکرد واقعی فرد در این دوره دارد.']);
    aoa.push([]);

    var header = ['حوزه', 'سؤال'];
    levels.forEach(function (name, i) {
      header.push((config.answerScale[name]) + ' — ' + name);
    });
    aoa.push(header);

    (config.questions || []).forEach(function (q) {
      var row = [q.domain || '', q.text || ''];
      for (var i = 0; i < levels.length; i++) {
        var a = (q.anchors || [])[i];
        row.push(a ? ((a.label ? '(' + a.label + ')\n\n' : '') + (a.text || '')) : '');
      }
      aoa.push(row);
    });

    aoa.push([]);
    aoa.push(['وزن هر سؤال در امتیاز عملکرد']);
    aoa.push(['کد', 'حوزه', 'وزن', 'در محاسبه']);
    (config.questions || []).forEach(function (q) {
      aoa.push([q.id.toUpperCase(), q.domain || '', q.weight === undefined ? 1 : q.weight,
                q.impact ? 'امتیاز ویژه' : (q.scored === false ? 'خیر' : 'بله')]);
    });
    aoa.push([]);
    aoa.push(['سؤال اثرگذاری ویژه تنها زمانی امتیاز می‌گیرد که امتیاز کارانهٔ حاصل از سایر سؤالات ' +
              'حداقل ' + (config.specialImpactMinScore || 0) + ' باشد.']);
    aoa.push(['امتیاز اثرگذاری ویژه باید مضربی از ' + (config.specialImpactStep || 50) + ' باشد.']);

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 24 }, { wch: 60 }].concat(levels.map(function () { return { wch: 40 }; }));
    /* Anchor cells hold a paragraph each; without wrapping and tall rows the
       sheet is unreadable. Rows are sized for the longest anchor. */
    ws['!rows'] = aoa.map(function (row, i) {
      return (i >= 4 && i < 4 + (config.questions || []).length) ? { hpt: 108 } : null;
    });
    ws['!postprocess'] = {
      xSplit: 2, ySplit: 4, headerRow: 4, numberFormats: {},
      wrapRows: { from: 5, to: 4 + (config.questions || []).length,
                  fromCol: 0, toCol: 1 + levels.length }
    };
    return ws;
  }

  /* ------------------------------------------------------------------------
   * Employee master template — the payroll team's own layout
   * ---------------------------------------------------------------------- */
  function buildEmployeeTemplate(config, employees, opts) {
    opts = opts || {};
    var XLSX = opts.XLSX || (typeof window !== 'undefined' ? window.XLSX : null);

    var aoa = [PAYROLL_COLUMNS.map(function (c) { return c.label; })];
    (employees || []).forEach(function (e) {
      aoa.push(PAYROLL_COLUMNS.map(function (c) {
        if (c.output) return '';
        var v = e[c.key];
        return v === null || v === undefined ? '' : v;
      }));
    });
    for (var b = 0; b < (employees && employees.length ? 5 : 30); b++) {
      aoa.push(PAYROLL_COLUMNS.map(function () { return ''; }));
    }

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = PAYROLL_COLUMNS.map(function (c) { return { wch: c.width }; });
    ws['!postprocess'] = { xSplit: 1, ySplit: 1, headerRow: 1, numberFormats: { N: '#,##0' } };

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Karaneh Employee Master');
    var mws = XLSX.utils.aoa_to_sheet(metaSheetRows(config, opts.period, 'employee'));
    mws['!cols'] = [{ wch: 24 }, { wch: 60 }];
    XLSX.utils.book_append_sheet(wb, mws, META_SHEET);
    wb.Workbook = { Views: [{ RTL: false }], Sheets: [{}, { Hidden: 1 }] };
    return wb;
  }

  return {
    META_SHEET: META_SHEET,
    QUESTIONNAIRE_SHEET: QUESTIONNAIRE_SHEET,
    PAYROLL_COLUMNS: PAYROLL_COLUMNS,
    describe: describe,
    signature: signature,
    readMeta: readMeta,
    verifyAgainstTemplate: verifyAgainstTemplate,
    verifyColumns: verifyColumns,
    buildBarsSheet: buildBarsSheet,
    buildQuestionnaireTemplate: buildQuestionnaireTemplate,
    buildEmployeeTemplate: buildEmployeeTemplate
  };
}));
