/* ============================================================================
 * calculation-engine.js — Karaneh Management System
 * ----------------------------------------------------------------------------
 * A faithful re-implementation of the Merit.xlsb calculation chain.
 *
 * Every function below maps 1:1 onto a column of one of the two calculation
 * sheets in the reference workbook. The Excel column letter is quoted on each
 * step so a reviewer can put the workbook and this file side by side.
 *
 *   Sheet 1 — «پرسشنامه کارانه تیمی»  (questionnaire → karaneh coefficient)
 *   Sheet 2 — «روش پرداخت کارانه»     (coefficient → rial payout)
 *
 * The engine is pure: it takes (records, config) and returns a new result set.
 * It never touches the DOM, storage, or globals, so it can be unit-tested
 * under Node and re-run on every keystroke in the browser.
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.KaranehEngine = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ------------------------------------------------------------------------
   * Default configuration — mirrors the constants found in Merit.xlsb.
   * Nothing here is hard-coded into the formulas; the engine only ever reads
   * from the config object it is handed.
   * ---------------------------------------------------------------------- */
  var DEFAULT_CONFIG = {
    /* روش پرداخت کارانه!C1 — total pot to distribute, in rial */
    budget: 100000000000,

    /* Data!H4:I8 — questionnaire answer wording → numeric score */
    answerScale: {
      'خیلی کم': 1,
      'کم': 2,
      'متوسط': 3,
      'زیاد': 4,
      'خیلی زیاد': 5
    },

    /* Data!C4:D9 — job level → grade score (the VLOOKUP target) */
    gradeMap: { '0': 100, '1': 150, '2': 200, '3': 250, '3H': 300, '4': 350 },

    /* «عدد کارانه» scaling: score × maxPerformanceScore / questionCount */
    maxPerformanceScore: 120,
    questionCount: 4,

    /* The questionnaire itself. Text, weight and whether an answer is scored
       are all editable in the designer, so the instrument can change without
       touching this file. Merit scores Q1..Q4 with equal weight; Q5 is
       collected but deliberately excluded from the score. */
    questions: [
      { id: 'q1', text: 'تا چه حد فرد وظایف محوله را با بهره‌گیری بهینه از زمان، منابع و انرژی انجام می‌دهد؟', weight: 1, scored: true },
      { id: 'q2', text: 'آیا فرد در موقعیت‌های واقعی کاری، رفتارهای همسو با ارزش‌های سازمانی نشان می‌دهد؟', weight: 1, scored: true },
      { id: 'q3', text: 'تا چه حد فرد با ذی‌نفعان داخل و خارج تیم همکاری مؤثر دارد؟', weight: 1, scored: true },
      { id: 'q4', text: 'تا چه حد فرد مالکیت خروجی‌های خود را می‌پذیرد و آن‌ها را به سطح قابل قبول یا فراتر از انتظار می‌رساند؟', weight: 1, scored: true },
      { id: 'q5', text: 'نسبت به بازخوردها واکنش سازنده نشان داده و در مسیر رشد یا آمادگی برای مسئولیت‌های بالاتر حرکت کرده است', weight: 0, scored: false }
    ],

    /* Text of the special-impact question shown on the template. */
    specialImpactQuestion: 'فرد تا چه حد از طریق کار خود، ارزش‌آفرینی قابل مشاهده و مبتنی بر شواهد برای تیم، بخش یا سازمان ایجاد کرده است؟',

    /* روش پرداخت کارانه!D4 — minimum evaluation score to receive karaneh.
       thresholdMode 'gt' reproduces Merit exactly: a person scoring exactly
       2.00 is excluded, so the test is score > threshold, not >=. */
    minPerformanceThreshold: 2,
    thresholdMode: 'gt',            // 'gt' | 'gte'
    belowThresholdRule: 'zero',     // 'zero' | 'keep'

    /* روش پرداخت کارانه!D2 — grade influence. 0 in the reference workbook,
       which is why job level currently has no effect on the payout. */
    gradeImpactFactor: 0,

    /* «کارانه اثرگذاری ویژه» — flat bonus coefficient for special impact. */
    specialImpactAmount: 300,

    /* The special-impact question may only be answered once the four scored
       questions have already earned at least this many karaneh points. This is
       what the reference workbook does: the one employee flagged for special
       impact whose base score was 90 received no bonus, while the three
       scoring 105, 120 and 105 all did. */
    specialImpactMinScore: 100,

    /* پرسشنامه کارانه تیمی!B5 = A5 × 100 — the coefficient pool is normalised
       so the average employee carries exactly this many coefficient points. */
    baselineCoefficientPerPerson: 100,

    /* Toggles for the three balancing mechanisms. All on = Merit behaviour. */
    normalizeCoefficients: true,     // sheet 1, column R
    redistributeIneligible: true,    // sheet 2, J5
    redistributeHodDiff: true,       // sheet 2, O5

    /* 'global' reproduces Merit (one pool for the whole workbook).
       'division' pools each organisational unit separately. */
    normalizationScope: 'global',

    /* Presentation only — never applied before the budget reconciliation. */
    displayDecimals: 0,

    /* Guard rails */
    allowNegativePayout: false
  };

  /* ------------------------------------------------------------------------
   * Small helpers
   * ---------------------------------------------------------------------- */
  function num(v) {
    if (v === null || v === undefined || v === '') return 0;
    var n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[,\s]/g, ''));
    return isFinite(n) ? n : 0;
  }

  function isBlank(v) {
    return v === null || v === undefined || v === '' ||
           (typeof v === 'string' && v.trim() === '');
  }

  function mergeConfig(cfg) {
    var out = {}, k;
    for (k in DEFAULT_CONFIG) if (DEFAULT_CONFIG.hasOwnProperty(k)) out[k] = DEFAULT_CONFIG[k];
    if (cfg) for (k in cfg) if (cfg.hasOwnProperty(k) && cfg[k] !== undefined) out[k] = cfg[k];
    return out;
  }

  /* Job level keys arrive as 3, '3', '3.0', ' 3h ' — normalise before lookup. */
  function normalizeJobLevel(v) {
    if (isBlank(v)) return '';
    var s = String(v).trim().toUpperCase().replace(/\s+/g, '');
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    return s;
  }

  /* ==========================================================================
   * STAGE 1 — «پرسشنامه کارانه تیمی»
   * ========================================================================*/

  /** The questions that actually contribute to the score, in order. */
  function scoredQuestions(cfg) {
    return (cfg.questions || []).filter(function (q) {
      return q.scored !== false && num(q.weight) > 0;
    });
  }

  /**
   * Column K «ارزیابی» — weighted mean of the scored questionnaire answers.
   * Excel:  =AVERAGE(F8:I8) after each answer is mapped through Data!H:I
   *
   * With every weight at 1 this is the plain average the workbook computes.
   * Unequal weights let one question count for more without changing the
   * 1..5 range of the result.
   *
   * Returns null when any scored answer is missing or uses wording the scale
   * does not recognise, so incomplete questionnaires surface as a validation
   * error instead of a silent zero.
   */
  function calculatePerformanceScore(record, config) {
    var cfg = mergeConfig(config);
    var qs = scoredQuestions(cfg);
    if (!qs.length) return null;
    var total = 0, weight = 0, i, raw, mapped, w;
    for (i = 0; i < qs.length; i++) {
      raw = record[qs[i].id];
      if (isBlank(raw)) return null;
      mapped = typeof raw === 'number' ? raw : cfg.answerScale[String(raw).trim()];
      if (mapped === undefined) return null;
      w = num(qs[i].weight);
      total += mapped * w;
      weight += w;
    }
    return weight ? total / weight : null;
  }

  /**
   * Column L «عدد کارانه» — rescale the 1..5 mean onto the karaneh band.
   * Excel:  =K8*120/4
   */
  function calculateKaranehScore(performanceScore, config) {
    var cfg = mergeConfig(config);
    if (performanceScore === null) return null;
    return performanceScore * cfg.maxPerformanceScore / cfg.questionCount;
  }

  /**
   * Column N «کارانه اثرگذاری ویژه» — coefficient bonus for special impact.
   *
   * The bonus is gated: it applies only once the scored questions have already
   * earned `specialImpactMinScore` karaneh points. A manager can flag anyone,
   * but the flag only pays out above the bar. A per-record amount overrides
   * the configured flat value.
   *
   * @param karanehScore column L for this employee, or null if unscored.
   */
  function calculateSpecialImpact(record, config, karanehScore) {
    var cfg = mergeConfig(config);
    if (!isSpecialImpactUnlocked(karanehScore, cfg)) return 0;
    return specialImpactEntered(record, cfg);
  }

  /**
   * The bonus the manager's answer implies, before the gate is applied.
   * This is what column N of the workbook holds: employee 4 carries 300 there
   * even though column O — which applies the gate — added nothing.
   */
  function specialImpactEntered(record, config) {
    var cfg = mergeConfig(config);
    if (!isSpecialImpact(record)) return 0;
    var override = record.specialImpactAmount;
    if (!isBlank(override) && num(override) !== 0) return num(override);
    return num(cfg.specialImpactAmount);
  }

  /**
   * Whether the special-impact question may be answered at all. Drives both
   * the payout gate above and the enabled state of the input in the UI, so
   * the rule is stated in exactly one place.
   */
  function isSpecialImpactUnlocked(karanehScore, config) {
    var cfg = mergeConfig(config);
    var floor = num(cfg.specialImpactMinScore);
    if (!floor) return true;
    return karanehScore !== null && karanehScore !== undefined && karanehScore >= floor;
  }

  function isSpecialImpact(record) {
    var v = record.specialProject !== undefined && record.specialProject !== null
          ? record.specialProject : record.specialImpact;
    if (isBlank(v)) return false;
    if (v === true) return true;
    if (v === false) return false;
    var s = String(v).trim().toLowerCase();
    return s === 'بله' || s === 'yes' || s === 'y' || s === 'true' || s === '1';
  }

  /**
   * Column O «ضریب کارانه» — raw coefficient before normalisation.
   * Excel:  =L8+N8
   */
  function calculateRawCoefficient(karanehScore, specialImpact) {
    if (karanehScore === null) return null;
    return karanehScore + specialImpact;
  }

  /* ==========================================================================
   * STAGE 2 — «روش پرداخت کارانه»
   * ========================================================================*/

  /**
   * Column F «سطح شغلی/گرید» — grade score for a job level.
   * Excel:  =VLOOKUP(E8,Data!$C:$D,2,0)
   * Returns null for an unmapped level so it lands in the validation report
   * rather than quietly scoring zero.
   */
  function getGradeScore(jobLevel, config) {
    var cfg = mergeConfig(config);
    var key = normalizeJobLevel(jobLevel);
    if (key === '') return null;
    if (cfg.gradeMap.hasOwnProperty(key)) return num(cfg.gradeMap[key]);
    return null;
  }

  /**
   * Column G «عدد گرید» — how much the grade actually moves the payout.
   * Excel:  =F8*$D$2
   */
  function calculateGradeImpact(gradeScore, config) {
    var cfg = mergeConfig(config);
    return num(gradeScore) * num(cfg.gradeImpactFactor);
  }

  /**
   * Column I «عدد ارزیابی عملکرد» — the eligibility gate.
   * Excel:  =IF(H8>$D$4,H8,0)
   */
  function applyMinimumThreshold(performanceScore, config) {
    var cfg = mergeConfig(config);
    if (performanceScore === null) return 0;
    var thr = num(cfg.minPerformanceThreshold);
    var passes = cfg.thresholdMode === 'gte'
      ? performanceScore >= thr
      : performanceScore > thr;
    if (passes) return performanceScore;
    return cfg.belowThresholdRule === 'keep' ? performanceScore : 0;
  }

  /**
   * Column L «امتیاز کل» — grade contribution plus performance contribution.
   * Excel:  =G8+K8
   */
  function calculateTotalScore(gradeImpact, performanceContribution) {
    return num(gradeImpact) + num(performanceContribution);
  }

  /**
   * Column M «دریافتی قبل از تغییرات معاون بخش» — proportional share of budget.
   * Excel:  =L8/SUM($L$8:$L$107)*$C$1
   */
  function calculateBudgetAllocation(totalScore, sumOfTotalScores, budget) {
    if (!sumOfTotalScores) return 0;
    return num(totalScore) / sumOfTotalScores * num(budget);
  }

  /* ==========================================================================
   * THE PIPELINE
   * ========================================================================*/

  /**
   * Run the whole Merit chain over a set of questionnaire records.
   *
   * @param {Array<Object>} records  one entry per employee holding the raw,
   *        user-editable fields: q1..q5, jobLevel, specialProject,
   *        specialImpactAmount, hodAdjustment, hodComment.
   * @param {Object} config          see DEFAULT_CONFIG.
   * @returns {{rows:Array, totals:Object, issues:Array, groups:Object}}
   */
  function calculate(records, config) {
    var cfg = mergeConfig(config);
    var rows = [], issues = [], i, r, row;

    /* -- Pass 1: per-record values that depend on nothing else ------------ */
    for (i = 0; i < records.length; i++) {
      r = records[i];
      row = {
        employeeId:    r.employeeId,
        fullName:      r.fullName || '',
        division:      r.division || '',
        positionTitle: r.positionTitle || '',
        jobLevel:      normalizeJobLevel(r.jobLevel),
        specialProject: isSpecialImpact(r),
        hodComment:    r.hodComment || '',
        sourceFile:    r.sourceFile || '',
        excluded:      !!r.excluded,
        _input:        r
      };
      /* Answers are copied by the configured question ids, so adding a
         question in the designer needs no change here. */
      (cfg.questions || []).forEach(function (q) { row[q.id] = r[q.id]; });

      row.performanceScore   = calculatePerformanceScore(r, cfg);          // K
      row.performanceKaraneh = calculateKaranehScore(row.performanceScore, cfg); // L
      row.specialImpactUnlocked = isSpecialImpactUnlocked(row.performanceKaraneh, cfg);
      row.specialImpactEntered = specialImpactEntered(r, cfg);             // N
      row.specialImpactValue = calculateSpecialImpact(r, cfg, row.performanceKaraneh);
      row.specialImpactBlocked = row.specialProject && !row.specialImpactUnlocked;
      row.rawCoefficient     = calculateRawCoefficient(row.performanceKaraneh,
                                                       row.specialImpactValue); // O
      row.gradeScore         = getGradeScore(r.jobLevel, cfg);             // F
      row.gradeImpact        = calculateGradeImpact(row.gradeScore, cfg);  // G

      row.hasQuestionnaire   = row.performanceScore !== null;
      row.hodAdjustment      = isBlank(r.hodAdjustment) ? null : num(r.hodAdjustment);
      rows.push(row);
    }

    /* Records that cannot enter the maths at all: no usable questionnaire,
       or explicitly excluded as a duplicate. They keep a zero payout and are
       reported, but they must not distort any denominator. */
    var active = [];
    for (i = 0; i < rows.length; i++) {
      row = rows[i];
      if (row.excluded || !row.hasQuestionnaire) {
        row.inScope = false;
        row.finalCoefficient = null;
        row.evalScore = 0;
        row.eligibleEvalScore = 0;
        row.eligible = false;
        row.performanceContribution = 0;
        row.totalScore = 0;
        row.initialAllocation = 0;
        row.finalKaraneh = 0;
        row.status = row.excluded ? 'Exception' : 'Incomplete';
      } else {
        row.inScope = true;
        active.push(row);
      }
    }

    /* -- Pass 2: normalise the coefficient pool (sheet 1, column R) --------
       Merit keeps the average coefficient pinned at
       `baselineCoefficientPerPerson`, so special-impact bonuses are funded by
       everyone else rather than inflating the pool.

         O5 (excess) = SUM(O) - count × baseline
         N5 (per person) = O5 / count
         R  = O - N5
    ------------------------------------------------------------------------*/
    var groups = groupBy(active, cfg.normalizationScope === 'division'
      ? function (x) { return x.division || '—'; }
      : function () { return '__all__'; });

    var groupStats = {};
    Object.keys(groups).forEach(function (key) {
      var g = groups[key], sumRaw = 0, j;
      for (j = 0; j < g.length; j++) sumRaw += g[j].rawCoefficient;
      var baselineTotal = g.length * num(cfg.baselineCoefficientPerPerson);
      var excess = sumRaw - baselineTotal;
      var perPerson = cfg.normalizeCoefficients && g.length ? excess / g.length : 0;
      groupStats[key] = {
        count: g.length, sumRaw: sumRaw, baselineTotal: baselineTotal,
        excess: excess, perPerson: perPerson
      };
      for (j = 0; j < g.length; j++) {
        g[j].normalizationAdjustment = -perPerson;
        g[j].finalCoefficient = g[j].rawCoefficient - perPerson;   // column R
        g[j].normalizationGroup = key;
      }
    });

    /* -- Pass 3: eligibility gate + redistribution of forfeited points -----
         I  = IF(H > threshold, H, 0)
         J5 = SUM(R of the gated-out) / COUNT(eligible)
         K  = eligible ? R + J5 : 0
       The pool total is preserved: points the below-threshold employees lose
       are shared equally among everyone who cleared the bar.
    ------------------------------------------------------------------------*/
    var eligible = [], forfeited = 0;
    for (i = 0; i < active.length; i++) {
      row = active[i];
      row.evalScore = row.performanceScore;                                  // H
      row.eligibleEvalScore = applyMinimumThreshold(row.performanceScore, cfg); // I
      row.eligible = row.eligibleEvalScore !== 0;
      if (row.eligible) eligible.push(row);
      else forfeited += row.finalCoefficient;
    }
    var ineligibleRedistribution =
      (cfg.redistributeIneligible && eligible.length) ? forfeited / eligible.length : 0;

    var sumTotalScore = 0;
    for (i = 0; i < active.length; i++) {
      row = active[i];
      row.baseCoefficient = row.eligible ? row.finalCoefficient : 0;         // J
      row.redistributionBonus = row.eligible ? ineligibleRedistribution : 0;
      row.performanceContribution = row.eligible
        ? row.finalCoefficient + ineligibleRedistribution : 0;               // K
      row.totalScore = row.eligible
        ? calculateTotalScore(row.gradeImpact, row.performanceContribution) : 0; // L
      sumTotalScore += row.totalScore;
    }

    /* -- Pass 4: budget allocation (column M) ----------------------------- */
    var budget = num(cfg.budget);
    for (i = 0; i < active.length; i++) {
      row = active[i];
      row.initialAllocation = calculateBudgetAllocation(row.totalScore, sumTotalScore, budget);
    }

    /* -- Pass 5: HOD overrides + rebalancing (columns N, O, O5, P, Q) ------
       An override replaces the person's payout outright. The difference
       between what they would have received and what the HOD granted is
       spread evenly across every other eligible employee, so the budget total
       is preserved to the rial.
    ------------------------------------------------------------------------*/
    var overriddenSum = 0, unadjusted = [];
    for (i = 0; i < active.length; i++) {
      row = active[i];
      row.isOverridden = row.eligible && row.hodAdjustment !== null;
      if (row.isOverridden) {
        row.diff = row.initialAllocation - row.hodAdjustment;               // O
        overriddenSum += row.diff;
      } else {
        row.diff = null;
        if (row.eligible) unadjusted.push(row);
      }
    }
    var hodRedistribution =
      (cfg.redistributeHodDiff && unadjusted.length) ? overriddenSum / unadjusted.length : 0;

    for (i = 0; i < active.length; i++) {
      row = active[i];
      if (!row.eligible) {
        row.hodRedistribution = 0;
        row.finalKaraneh = 0;
      } else if (row.isOverridden) {
        row.hodRedistribution = 0;
        row.finalKaraneh = row.hodAdjustment;                                // P/Q
      } else {
        row.hodRedistribution = hodRedistribution;
        row.finalKaraneh = row.initialAllocation + hodRedistribution;        // P/Q
      }
    }

    /* -- Pass 6: totals, statuses, and the issue log ----------------------- */
    var totals = {
      recordCount:            rows.length,
      inScopeCount:           active.length,
      eligibleCount:          eligible.length,
      ineligibleCount:        active.length - eligible.length,
      excludedCount:          rows.length - active.length,
      overriddenCount:        0,
      sumRawCoefficient:      0,
      sumFinalCoefficient:    0,
      sumTotalScore:          sumTotalScore,
      sumInitialAllocation:   0,
      sumFinalKaraneh:        0,
      budget:                 budget,
      ineligibleRedistribution: ineligibleRedistribution,
      hodRedistribution:      hodRedistribution,
      forfeitedCoefficient:   forfeited,
      negativePayoutCount:    0,
      groupStats:             groupStats
    };

    for (i = 0; i < rows.length; i++) {
      row = rows[i];
      if (row.inScope) {
        totals.sumRawCoefficient   += row.rawCoefficient;
        totals.sumFinalCoefficient += row.finalCoefficient;
        totals.sumInitialAllocation += row.initialAllocation;
        totals.sumFinalKaraneh     += row.finalKaraneh;
        if (row.isOverridden) totals.overriddenCount++;
        if (row.finalKaraneh < 0) { totals.negativePayoutCount++; row.negative = true; }
      }
      row.status = deriveStatus(row);
    }

    totals.allocatedBudget = totals.sumFinalKaraneh;
    totals.remainingBudget = budget - totals.sumFinalKaraneh;
    /* Floating point leaves a few rial of dust on a 100-billion pot; anything
       under a rial is reconciliation noise, not an overrun. */
    totals.budgetOverrun = totals.remainingBudget < -1;
    totals.budgetStatus = totals.budgetOverrun ? 'OVERRUN'
      : (totals.negativePayoutCount > 0 ? 'INVALID' : 'BALANCED');

    return { rows: rows, totals: totals, config: cfg, issues: issues };
  }

  /**
   * Per-employee workflow state, derived rather than stored so it can never
   * drift out of step with the numbers.
   */
  function deriveStatus(row) {
    if (row.excluded) return 'Exception';
    if (!row.hasQuestionnaire) return 'Pending';
    if (row.gradeScore === null) return 'Incomplete';
    if (row.negative) return 'Exception';
    if (!row.eligible) return 'Below Threshold';
    if (row.isOverridden) return row.hodComment ? 'HOD Adjusted' : 'Exception';
    if (row.finalized) return 'Finalized';
    return 'Calculated';
  }

  function groupBy(list, keyFn) {
    var out = {}, i, k;
    for (i = 0; i < list.length; i++) {
      k = keyFn(list[i]);
      (out[k] || (out[k] = [])).push(list[i]);
    }
    return out;
  }

  /**
   * Budget guard used by the HOD screen before a change is committed.
   * Returns the headroom an override may consume without pushing anyone below
   * zero, which is the real failure mode: the total always reconciles to the
   * budget, but an oversized override drains everybody else.
   */
  function validateBudget(result) {
    var t = result.totals, problems = [];
    if (t.budgetOverrun) {
      problems.push({
        severity: 'error', code: 'BUDGET_OVERRUN',
        message: 'مجموع کارانه نهایی از بودجه تعیین‌شده بیشتر است.',
        detail: { allocated: t.allocatedBudget, budget: t.budget }
      });
    }
    if (t.negativePayoutCount > 0) {
      problems.push({
        severity: 'error', code: 'NEGATIVE_PAYOUT',
        message: t.negativePayoutCount + ' نفر با اعمال تغییرات معاون بخش دریافتی منفی پیدا کرده‌اند.',
        detail: { count: t.negativePayoutCount }
      });
    }
    return { ok: problems.length === 0, problems: problems };
  }

  /**
   * How much a single HOD override may be raised to before it drives the
   * smallest remaining payout below zero. Used to show the allowed ceiling
   * instead of just refusing the edit.
   */
  function maxAllowedAdjustment(result, employeeId) {
    var rows = result.rows, i, row, target = null, others = [];
    for (i = 0; i < rows.length; i++) {
      row = rows[i];
      if (!row.inScope || !row.eligible) continue;
      if (row.employeeId === employeeId) target = row;
      else if (!row.isOverridden) others.push(row);
    }
    if (!target) return null;
    if (!others.length) return target.initialAllocation;

    /* Raising this person by X lowers each of the N unadjusted peers by X/N.
       The binding constraint is the peer with the smallest headroom. */
    var minHeadroom = Infinity;
    for (i = 0; i < others.length; i++) {
      minHeadroom = Math.min(minHeadroom, others[i].finalKaraneh);
    }
    var current = target.isOverridden ? target.hodAdjustment : target.initialAllocation;
    return current + minHeadroom * others.length;
  }

  return {
    DEFAULT_CONFIG: DEFAULT_CONFIG,
    mergeConfig: mergeConfig,
    normalizeJobLevel: normalizeJobLevel,
    isSpecialImpact: isSpecialImpact,
    scoredQuestions: scoredQuestions,
    isSpecialImpactUnlocked: isSpecialImpactUnlocked,
    calculatePerformanceScore: calculatePerformanceScore,
    calculateKaranehScore: calculateKaranehScore,
    calculateSpecialImpact: calculateSpecialImpact,
    specialImpactEntered: specialImpactEntered,
    calculateRawCoefficient: calculateRawCoefficient,
    getGradeScore: getGradeScore,
    calculateGradeImpact: calculateGradeImpact,
    applyMinimumThreshold: applyMinimumThreshold,
    calculateTotalScore: calculateTotalScore,
    calculateBudgetAllocation: calculateBudgetAllocation,
    calculate: calculate,
    validateBudget: validateBudget,
    maxAllowedAdjustment: maxAllowedAdjustment
  };
}));
