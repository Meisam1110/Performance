/* ============================================================================
 * excel-import.js — file ingestion and smart column mapping
 * ----------------------------------------------------------------------------
 * Reads .xlsx / .xls / .xlsb / .csv through SheetJS, finds the header row
 * wherever it happens to sit, and resolves real-world header wording onto the
 * canonical field names the calculation engine expects.
 *
 * Position is never trusted. A file whose columns have been reordered, or
 * whose headers were retyped slightly differently, still imports correctly.
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ExcelImport = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ------------------------------------------------------------------------
   * Canonical fields and the header wording seen in the wild. Admin-editable:
   * the whole table is persisted in state.columnMappings and surfaced in the
   * Settings screen, so a new spelling is a data change, not a code change.
   * ---------------------------------------------------------------------- */
  var DEFAULT_MAPPINGS = {
    employeeId: {
      label: 'شماره پرسنلی', group: 'both',
      synonyms: ['شماره پرسنلی', 'کد پرسنلی', 'کدپرسنلی', 'شماره کارمندی', 'پرسنلی',
                 'emp no', 'empno', 'employee no', 'employee number', 'employee id',
                 'personnel no', 'personnel number', 'staff no', 'staff id', 'id']
    },
    employeeStatus: {
      label: 'وضعیت', group: 'employee',
      synonyms: ['وضعیت', 'وضعیت پرسنل', 'emp status', 'employee status', 'status']
    },
    firstName: {
      label: 'نام', group: 'employee',
      synonyms: ['نام', 'first name', 'firstname', 'given name']
    },
    lastName: {
      label: 'نام خانوادگی', group: 'employee',
      synonyms: ['نام خانوادگی', 'فامیلی', 'last name', 'lastname', 'surname', 'family name']
    },
    fullName: {
      label: 'نام و نام خانوادگی', group: 'both',
      synonyms: ['نام و نام خانوادگی', 'نام کامل', 'full name', 'fullname', 'name',
                 'employee name']
    },
    nationalId: {
      label: 'کد ملی', group: 'employee',
      synonyms: ['کد ملی', 'کدملی', 'شماره ملی', 'national id', 'national code', 'nid']
    },
    gender: {
      label: 'جنسیت', group: 'employee',
      synonyms: ['جنسیت', 'gender', 'sex']
    },
    dateOfEmployment: {
      label: 'تاریخ استخدام', group: 'employee',
      synonyms: ['date of employment', 'تاریخ استخدام', 'تاریخ شروع', 'start date',
                 'hire date', 'max of start date', 'date of employment-shamsi',
                 'max of start date-shamsi']
    },
    dateOfLeaving: {
      label: 'تاریخ خروج', group: 'employee',
      synonyms: ['تاریخ خروج', 'تاریخ ترک خدمت', 'date of leaving', 'leaving date',
                 'termination date', 'date of leaving-shamsi']
    },
    positionTitle: {
      label: 'عنوان شغلی', group: 'both',
      synonyms: ['عنوان شغلی', 'سمت', 'شغل', 'pos title', 'position title', 'position',
                 'job title', 'title']
    },
    assignmentType: {
      label: 'نوع همکاری', group: 'employee',
      synonyms: ['نوع همکاری', 'نوع قرارداد کاری', 'assignment type', 'assignment']
    },
    employmentType: {
      label: 'نوع استخدام', group: 'employee',
      synonyms: ['نوع استخدام', 'employment type', 'employment']
    },
    jobLevel: {
      label: 'سطح شغلی (JL)', group: 'both',
      synonyms: ['jl', 'سطح شغلی', 'سطح', 'گرید', 'job level', 'joblevel', 'level',
                 'grade level', 'job grade']
    },
    division: {
      label: 'واحد سازمانی', group: 'both',
      synonyms: ['واحد سازمانی', 'واحد', 'بخش', 'دپارتمان', 'division', 'division alias',
                 'department', 'dept', 'unit', 'business unit', 'مرکز هزینه', 'cost center']
    },
    department: {
      label: 'دپارتمان', group: 'employee',
      synonyms: ['دپارتمان', 'زیرمجموعه', 'sub unit', 'sub-unit', 'section']
    },
    workingDays: {
      label: 'روزهای کارکرد', group: 'employee',
      synonyms: ['روزهای کارکرد', 'روز کارکرد', 'کارکرد', 'working days', "working day's",
                 "working day'sq1", 'working days q1', 'total', 'working day']
    },
    probationStatus: {
      label: 'وضعیت دوره آزمایشی', group: 'employee',
      synonyms: ['وضعیت دوره آزمایشی', 'دوره آزمایشی', 'probation', 'probation checking',
                 'probation status']
    },
    checking: {
      label: 'بررسی', group: 'employee',
      synonyms: ['بررسی', 'checking', 'check']
    },
    comment: {
      label: 'توضیحات', group: 'both',
      synonyms: ['توضیحات', 'توضیح', 'کامنت', 'comment', 'comments', 'note', 'notes',
                 'description']
    },
    directManager: {
      label: 'مدیر مستقیم', group: 'employee',
      synonyms: ['مدیر مستقیم', 'سرپرست مستقیم', 'direct manager', 'manager',
                 'manager name', 'line manager', 'reports to']
    },
    managerLevel1: { label: 'مدیر سطح ۱', group: 'employee', synonyms: ['مدیر سطح 1', 'manager level 1', 'manager 1', 'manager name 1', 'manager hierarchy 1'] },
    managerLevel2: { label: 'مدیر سطح ۲', group: 'employee', synonyms: ['مدیر سطح 2', 'manager level 2', 'manager 2', 'manager name 2', 'manager hierarchy 2'] },
    managerLevel3: { label: 'مدیر سطح ۳', group: 'employee', synonyms: ['مدیر سطح 3', 'manager level 3', 'manager 3', 'manager name 3', 'manager hierarchy 3'] },

    /* ---- questionnaire ---- */
    q1: { label: 'سؤال ۱', group: 'questionnaire', synonyms: ['q1', 'س1', 'سوال 1', 'سؤال 1', 'question 1'] },
    q2: { label: 'سؤال ۲', group: 'questionnaire', synonyms: ['q2', 'س2', 'سوال 2', 'سؤال 2', 'question 2'] },
    q3: { label: 'سؤال ۳', group: 'questionnaire', synonyms: ['q3', 'س3', 'سوال 3', 'سؤال 3', 'question 3'] },
    q4: { label: 'سؤال ۴', group: 'questionnaire', synonyms: ['q4', 'س4', 'سوال 4', 'سؤال 4', 'question 4'] },
    q5: { label: 'سؤال ۵ (اطلاعاتی)', group: 'questionnaire', synonyms: ['q5', 'س5', 'سوال 5', 'سؤال 5', 'question 5'] },
    specialProject: {
      label: 'اثرگذاری ویژه', group: 'questionnaire',
      synonyms: ['اثرگذاری ویژه', 'اثر گذاری ویژه', 'پروژه ویژه', 'special project',
                 'special impact', 'اثرگذاری خاص']
    },
    specialImpactAmount: {
      label: 'امتیاز اثرگذاری ویژه', group: 'questionnaire',
      synonyms: ['کارانه اثرگذاری ویژه', 'کارانه پروژه ویژه', 'امتیاز اثرگذاری ویژه',
                 'special impact amount', 'special project amount']
    },
    hodAdjustment: {
      label: 'تغییرات معاون بخش', group: 'both',
      synonyms: ['تغییرات معاون بخش', 'تغییرات معاون', 'hod adjustment', 'hod override',
                 'hod adjust', 'مبلغ پیشنهادی معاون']
    },
    hodComment: {
      label: 'توضیح معاون بخش', group: 'both',
      synonyms: ['hod comment', 'توضیح معاون بخش', 'کامنت معاون', 'توضیح معاون']
    }
  };

  /* Sheet names that mark a file as a team questionnaire. */
  var QUESTIONNAIRE_SHEET_HINTS = [
    'پرسشنامه کارانه تیمی', 'پرسشنامه کارانه', 'پرسشنامه', 'questionnaire', 'team'
  ];

  /* ------------------------------------------------------------------------
   * Header normalisation. Persian text arrives with Arabic ي/ك, zero-width
   * joiners, Arabic-Indic digits and stray punctuation; all of it has to
   * collapse to one comparable form before matching.
   * ---------------------------------------------------------------------- */
  var DIGIT_MAP = {
    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4', '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9',
    '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4', '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'
  };

  function normalizeText(v) {
    if (v === null || v === undefined) return '';
    var s = String(v);
    s = s.replace(/[۰-۹٠-٩]/g, function (d) { return DIGIT_MAP[d]; });
    s = s.replace(/‌|‏|‎|﻿/g, ' ');   // ZWNJ + bidi marks
    s = s.replace(/ي/g, 'ی').replace(/ك/g, 'ک').replace(/ﻩ|ة/g, 'ه');
    s = s.replace(/[ً-ْ]/g, '');                // harakat
    s = s.replace(/[()\[\]{}:؛;,.\-_/\\*#"'?؟!]/g, ' ');
    s = s.replace(/\s+/g, ' ').trim().toLowerCase();
    return s;
  }

  /* Levenshtein, capped — used only as a last resort for near-miss headers. */
  function editDistance(a, b) {
    if (a === b) return 0;
    if (!a.length || !b.length) return Math.max(a.length, b.length);
    var prev = [], cur = [], i, j;
    for (j = 0; j <= b.length; j++) prev[j] = j;
    for (i = 1; i <= a.length; i++) {
      cur[0] = i;
      for (j = 1; j <= b.length; j++) {
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1,
                          prev[j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1));
      }
      for (j = 0; j <= b.length; j++) prev[j] = cur[j];
    }
    return prev[b.length];
  }

  /**
   * Score how well one header cell matches one canonical field.
   * 100 exact · 90 exact after normalisation · 70-85 containment · 50-69 fuzzy.
   */
  function scoreHeader(header, field, mappings) {
    var norm = normalizeText(header);
    if (!norm) return 0;
    var syns = mappings[field].synonyms, best = 0, i, s, d, ratio;
    for (i = 0; i < syns.length; i++) {
      s = normalizeText(syns[i]);
      if (!s) continue;
      /* Exact match. Earlier synonyms are the canonical spellings, so a tiny
         ordinal penalty makes them win a tie against a later variant — e.g.
         "Date of Employment" beats "Max of Start Date" when a file has both. */
      if (norm === s) return 100 - i * 0.1;
      if (norm.indexOf(s) === 0 || s.indexOf(norm) === 0) best = Math.max(best, 85);
      else if (norm.indexOf(s) !== -1) best = Math.max(best, 75);
      else if (s.indexOf(norm) !== -1 && norm.length >= 3) best = Math.max(best, 70);
      else if (Math.abs(norm.length - s.length) <= 3 && s.length >= 4) {
        d = editDistance(norm, s);
        ratio = 1 - d / Math.max(norm.length, s.length);
        if (ratio >= 0.8) best = Math.max(best, Math.round(50 + ratio * 19));
      }
    }
    return best;
  }

  /**
   * Resolve a whole header row onto canonical fields.
   * Greedy by score, one column per field, so a strong match cannot be stolen
   * by a weaker one later in the row.
   *
   * @returns {{mapping:Object, unmapped:Array, decisions:Array}}
   *          mapping is field → column index.
   */
  function mapColumns(headers, opts) {
    opts = opts || {};
    var mappings = opts.mappings || DEFAULT_MAPPINGS;
    var group = opts.group;               // 'employee' | 'questionnaire' | undefined
    var minScore = opts.minScore || 60;

    var fields = Object.keys(mappings).filter(function (f) {
      if (!group) return true;
      var g = mappings[f].group;
      return g === 'both' || g === group;
    });

    var candidates = [];
    fields.forEach(function (field) {
      headers.forEach(function (h, idx) {
        var sc = scoreHeader(h, field, mappings);
        if (sc >= minScore) candidates.push({ field: field, index: idx, score: sc, header: h });
      });
    });
    candidates.sort(function (a, b) { return b.score - a.score; });

    var mapping = {}, usedCol = {}, decisions = [];
    candidates.forEach(function (c) {
      if (mapping[c.field] !== undefined || usedCol[c.index]) return;
      mapping[c.field] = c.index;
      usedCol[c.index] = true;
      decisions.push(c);
    });

    var unmapped = headers.map(function (h, i) { return { index: i, header: h }; })
      .filter(function (x) { return !usedCol[x.index] && normalizeText(x.header) !== ''; });

    return { mapping: mapping, unmapped: unmapped, decisions: decisions };
  }

  /**
   * Positional fallback for the reference layout of «پرسشنامه کارانه تیمی».
   * Only fills fields the header matcher could not resolve, and only when the
   * row shape matches, so it can never override a confident header match.
   */
  var QUESTIONNAIRE_POSITIONS = {
    employeeId: 0, fullName: 1, division: 2, positionTitle: 3, jobLevel: 4,
    q1: 5, q2: 6, q3: 7, q4: 8, q5: 9, specialProject: 12, specialImpactAmount: 13
  };

  function applyPositionalFallback(mapping, headers) {
    var filled = [];
    if (headers.length < 10) return filled;
    Object.keys(QUESTIONNAIRE_POSITIONS).forEach(function (f) {
      var idx = QUESTIONNAIRE_POSITIONS[f];
      if (mapping[f] === undefined && idx < headers.length) {
        var taken = Object.keys(mapping).some(function (k) { return mapping[k] === idx; });
        if (!taken) { mapping[f] = idx; filled.push(f); }
      }
    });
    return filled;
  }

  /* ------------------------------------------------------------------------
   * Header row detection — the reference files carry a title band, a merged
   * group row, and sometimes two header rows before the data starts.
   * ---------------------------------------------------------------------- */
  function detectHeaderRow(rows, opts) {
    var limit = Math.min(rows.length, 25), best = null, i;
    for (i = 0; i < limit; i++) {
      var headers = rows[i] || [];
      var nonEmpty = headers.filter(function (h) { return normalizeText(h) !== ''; }).length;
      if (nonEmpty < 2) continue;
      var res = mapColumns(headers, opts);
      var score = Object.keys(res.mapping).length * 10 + nonEmpty;
      /* A header row must at least identify the person. */
      if (res.mapping.employeeId === undefined &&
          res.mapping.fullName === undefined) score -= 25;
      if (!best || score > best.score) {
        best = { index: i, score: score, headers: headers, result: res };
      }
    }
    return best;
  }

  /* ------------------------------------------------------------------------
   * Value coercion
   * ---------------------------------------------------------------------- */
  function cellString(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).replace(/‌/g, '‌').trim();
  }

  function cellNumber(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return v;
    var s = String(v).replace(/[۰-۹٠-٩]/g, function (d) { return DIGIT_MAP[d]; })
                     .replace(/[,\s٬]/g, '');
    var n = parseFloat(s);
    return isFinite(n) ? n : null;
  }

  /** Employee numbers must survive as text: leading zeros and codes like BEKI001. */
  function cellId(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(v);
    return String(v).replace(/[۰-۹٠-٩]/g, function (d) { return DIGIT_MAP[d]; }).trim();
  }

  var NUMERIC_FIELDS = { workingDays: 1, specialImpactAmount: 1, hodAdjustment: 1 };

  function buildRecord(row, mapping, meta) {
    var rec = {}, empty = true;
    Object.keys(mapping).forEach(function (field) {
      var raw = row[mapping[field]];
      var val;
      if (field === 'employeeId') val = cellId(raw);
      else if (NUMERIC_FIELDS[field]) val = cellNumber(raw);
      else val = cellString(raw);
      if (val !== '' && val !== null && val !== undefined) empty = false;
      rec[field] = val;
    });
    if (empty) return null;
    if (!rec.fullName && (rec.firstName || rec.lastName)) {
      rec.fullName = ((rec.firstName || '') + ' ' + (rec.lastName || '')).trim();
    }
    rec.sourceFile = meta.fileName;
    rec.sourceSheet = meta.sheetName;
    rec.sourceRow = meta.rowNumber;
    rec.importedAt = meta.importedAt;
    rec.importBatchId = meta.batchId;
    return rec;
  }

  /* ------------------------------------------------------------------------
   * Sheet selection
   * ---------------------------------------------------------------------- */
  function pickSheet(workbook, kind) {
    var names = workbook.SheetNames, i, n;
    if (kind === 'questionnaire') {
      for (i = 0; i < names.length; i++) {
        n = normalizeText(names[i]);
        for (var j = 0; j < QUESTIONNAIRE_SHEET_HINTS.length; j++) {
          if (n.indexOf(normalizeText(QUESTIONNAIRE_SHEET_HINTS[j])) !== -1) return names[i];
        }
      }
    }
    /* Otherwise take the sheet with the most rows — the data sheet, not a
       lookup or a cover page. */
    var best = names[0], bestRows = -1;
    for (i = 0; i < names.length; i++) {
      if (names[i] === '_Template') continue;   // signature sheet, never data
      var ws = workbook.Sheets[names[i]];
      if (!ws || !ws['!ref']) continue;
      var range = XLSXRef(ws['!ref']);
      if (range > bestRows) { bestRows = range; best = names[i]; }
    }
    return best;
  }

  function XLSXRef(ref) {
    var m = /:(?:[A-Z]+)(\d+)$/.exec(ref);
    return m ? parseInt(m[1], 10) : 0;
  }

  /* ==========================================================================
   * PUBLIC — parse one workbook
   * ========================================================================*/

  /**
   * @param {ArrayBuffer} buffer  file contents
   * @param {Object} opts  { fileName, kind:'employee'|'questionnaire',
   *                         mappings, sheetName, XLSX }
   * @returns {{records, mapping, headers, sheetName, warnings, unmapped, ...}}
   */
  function parseWorkbook(buffer, opts) {
    opts = opts || {};
    var XLSX = opts.XLSX || (typeof window !== 'undefined' ? window.XLSX : null);
    if (!XLSX) throw new Error('SheetJS در دسترس نیست.');

    var wb = XLSX.read(buffer, { type: 'array', cellDates: true, cellNF: false, cellText: false });
    var sheetName = opts.sheetName || pickSheet(wb, opts.kind);
    var ws = wb.Sheets[sheetName];
    if (!ws) throw new Error('شیت «' + sheetName + '» در فایل یافت نشد.');

    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    var warnings = [];

    var mapOpts = {
      mappings: opts.mappings || DEFAULT_MAPPINGS,
      group: opts.kind === 'employee' ? 'employee'
           : opts.kind === 'questionnaire' ? 'questionnaire' : undefined
    };
    var header = detectHeaderRow(rows, mapOpts);
    if (!header) throw new Error('ردیف عنوان ستون‌ها در این فایل شناسایی نشد.');

    var mapping = header.result.mapping;
    var fallbackFilled = [];
    if (opts.kind === 'questionnaire') {
      fallbackFilled = applyPositionalFallback(mapping, header.headers);
      if (fallbackFilled.length) {
        warnings.push('این ستون‌ها از روی موقعیت استاندارد فایل مرجع تشخیص داده شدند: ' +
                      fallbackFilled.join('، '));
      }
    }

    if (mapping.employeeId === undefined) {
      throw new Error('ستون «شماره پرسنلی» شناسایی نشد. لطفاً نگاشت ستون‌ها را در تنظیمات اصلاح کنید.');
    }

    /* Skip a repeated header band: the reference file carries a second, short
       header row (Q1..Q5) directly under the descriptive one. */
    var start = header.index + 1;
    while (start < rows.length && looksLikeHeader(rows[start], mapping, mapOpts)) start++;

    var records = [], skipped = 0, importedAt = new Date().toISOString();
    var batchId = opts.batchId || (opts.fileName + '@' + importedAt);
    for (var i = start; i < rows.length; i++) {
      var rec = buildRecord(rows[i] || [], mapping, {
        fileName: opts.fileName || sheetName,
        sheetName: sheetName,
        rowNumber: i + 1,
        importedAt: importedAt,
        batchId: batchId
      });
      if (!rec) continue;
      if (!rec.employeeId) { skipped++; continue; }
      records.push(rec);
    }
    if (skipped) warnings.push(skipped + ' ردیف به دلیل نداشتن شماره پرسنلی نادیده گرفته شد.');

    return {
      workbook: wb,
      batchId: batchId,
      fileName: opts.fileName,
      sheetName: sheetName,
      sheetNames: wb.SheetNames,
      headerRow: header.index + 1,
      headers: header.headers,
      mapping: mapping,
      decisions: header.result.decisions,
      unmapped: header.result.unmapped,
      positionalFallback: fallbackFilled,
      records: records,
      warnings: warnings,
      importedAt: importedAt
    };
  }

  /** A row is another header band if its mapped cells repeat header wording. */
  function looksLikeHeader(row, mapping, mapOpts) {
    if (!row) return false;
    var idIdx = mapping.employeeId;
    var idVal = row[idIdx];
    if (idVal === null || idVal === undefined || idVal === '') {
      /* No id at all — only treat as header if the row still has text. */
      return row.some(function (c) { return normalizeText(c) !== ''; }) &&
             mapColumnsCount(row, mapOpts) >= 3;
    }
    if (typeof idVal === 'number') return false;
    return scoreHeader(idVal, 'employeeId', mapOpts.mappings) >= 70;
  }

  function mapColumnsCount(row, mapOpts) {
    return Object.keys(mapColumns(row, mapOpts).mapping).length;
  }

  return {
    DEFAULT_MAPPINGS: DEFAULT_MAPPINGS,
    QUESTIONNAIRE_SHEET_HINTS: QUESTIONNAIRE_SHEET_HINTS,
    normalizeText: normalizeText,
    scoreHeader: scoreHeader,
    mapColumns: mapColumns,
    detectHeaderRow: detectHeaderRow,
    parseWorkbook: parseWorkbook,
    cellNumber: cellNumber,
    cellId: cellId
  };
}));
