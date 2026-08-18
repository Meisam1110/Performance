/* ============================================================================
 * data-store.js — persistence layer for the Karaneh Management System
 * ----------------------------------------------------------------------------
 * Everything lives in the browser. IndexedDB is used when available (it copes
 * with thousands of employee rows); localStorage is the fallback for the
 * file:// double-click case on browsers that block IDB there.
 *
 * The API is deliberately storage-agnostic and promise-based so that swapping
 * in a Node/SQL backend later means re-implementing `read` and `write` only —
 * nothing above this file knows where the bytes end up.
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DataStore = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DB_NAME = 'karaneh-system';
  var DB_VERSION = 1;
  var STORE = 'state';
  var KEY = 'current';
  var LS_KEY = 'karaneh-system:state';

  /* ------------------------------------------------------------------------
   * The complete application state. One document, versioned, so an export is
   * a full backup and an import is a full restore.
   * ---------------------------------------------------------------------- */
  function emptyState() {
    return {
      schemaVersion: 1,
      period: 'Q1 1405',
      config: null,                 // filled from KaranehEngine.DEFAULT_CONFIG
      employees: [],                // employee master data
      questionnaires: [],           // consolidated questionnaire records
      importBatches: [],            // one entry per uploaded file
      columnMappings: null,         // admin-editable header synonyms
      auditLog: [],
      finalizedAt: null,
      updatedAt: null
    };
  }

  /* --------------------------------------------------------------- IndexedDB */
  var idbAvailable = (function () {
    try { return typeof indexedDB !== 'undefined' && indexedDB !== null; }
    catch (e) { return false; }
  }());

  function openDb() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
      req.onblocked = function () { reject(new Error('IndexedDB blocked')); };
    });
  }

  function idbRead() {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, 'readonly');
        var req = tx.objectStore(STORE).get(KEY);
        req.onsuccess = function () { resolve(req.result || null); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function idbWrite(state) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, 'readwrite');
        tx.objectStore(STORE).put(state, KEY);
        tx.oncomplete = function () { resolve(state); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  /* ------------------------------------------------------------ localStorage */
  function lsRead() {
    try {
      var raw = localStorage.getItem(LS_KEY);
      return Promise.resolve(raw ? JSON.parse(raw) : null);
    } catch (e) { return Promise.resolve(null); }
  }

  function lsWrite(state) {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(state));
      return Promise.resolve(state);
    } catch (e) {
      return Promise.reject(new Error(
        'ذخیره‌سازی مرورگر پر شده است. لطفاً خروجی بگیرید و داده‌ها را پاک کنید.'));
    }
  }

  /* ------------------------------------------------------------------ facade */
  var backend = 'memory';
  var memory = null;

  function read() {
    if (idbAvailable) {
      return idbRead()
        .then(function (s) { backend = 'indexeddb'; return s; })
        .catch(function () { backend = 'localstorage'; return lsRead(); });
    }
    backend = 'localstorage';
    return lsRead();
  }

  function write(state) {
    state.updatedAt = new Date().toISOString();
    memory = state;
    if (backend === 'indexeddb') {
      return idbWrite(state).catch(function () { backend = 'localstorage'; return lsWrite(state); });
    }
    if (backend === 'localstorage') return lsWrite(state);
    return Promise.resolve(state);
  }

  function clear() {
    memory = null;
    try { localStorage.removeItem(LS_KEY); } catch (e) { /* ignore */ }
    if (!idbAvailable) return Promise.resolve();
    return openDb().then(function (db) {
      return new Promise(function (resolve) {
        var tx = db.transaction(STORE, 'readwrite');
        tx.objectStore(STORE).delete(KEY);
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { resolve(); };
      });
    }).catch(function () { /* nothing to clear */ });
  }

  /* ==========================================================================
   * Audit trail — every change that can move money leaves a row here.
   * ========================================================================*/
  function audit(state, entry) {
    state.auditLog.push({
      id: state.auditLog.length + 1,
      timestamp: new Date().toISOString(),
      user: entry.user || state.currentUser || 'local-user',
      entity: entry.entity || 'employee',
      employeeId: entry.employeeId || '',
      employeeName: entry.employeeName || '',
      field: entry.field || '',
      oldValue: normaliseAuditValue(entry.oldValue),
      newValue: normaliseAuditValue(entry.newValue),
      reason: entry.reason || ''
    });
    /* Keep the log bounded so a long session cannot exhaust storage; the
       oldest entries are dropped only after an export-worthy amount. */
    if (state.auditLog.length > 20000) state.auditLog.splice(0, state.auditLog.length - 20000);
    return state;
  }

  function normaliseAuditValue(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'object') return JSON.stringify(v);
    return String(v);
  }

  /* ==========================================================================
   * Employee master — Employee No is the unique key.
   * ========================================================================*/

  /**
   * Merge an imported batch into the master list without ever silently
   * overwriting. Returns the incoming rows split into new / conflicting so the
   * caller can put the conflicts in front of the user.
   */
  function stageEmployees(existing, incoming) {
    var byId = {}, i;
    for (i = 0; i < existing.length; i++) byId[existing[i].employeeId] = existing[i];

    var fresh = [], conflicts = [], dupInBatch = [], seen = {};
    for (i = 0; i < incoming.length; i++) {
      var rec = incoming[i];
      if (!rec.employeeId) continue;
      if (seen[rec.employeeId]) {
        dupInBatch.push({ employeeId: rec.employeeId, incoming: rec, first: seen[rec.employeeId] });
        continue;
      }
      seen[rec.employeeId] = rec;
      var current = byId[rec.employeeId];
      if (!current) { fresh.push(rec); continue; }
      var diff = diffRecords(current, rec);
      if (!diff.length) continue;                 // identical re-import, nothing to do
      conflicts.push({
        employeeId: rec.employeeId,
        existing: current,
        incoming: rec,
        changedFields: diff,
        newer: whichIsNewer(current, rec)
      });
    }
    return { fresh: fresh, conflicts: conflicts, duplicatesWithinBatch: dupInBatch };
  }

  var COMPARABLE = ['fullName', 'firstName', 'lastName', 'employeeStatus', 'positionTitle',
    'assignmentType', 'employmentType', 'jobLevel', 'division', 'department',
    'dateOfEmployment', 'dateOfLeaving', 'workingDays', 'probationStatus',
    'directManager', 'managerLevel1', 'managerLevel2', 'managerLevel3',
    'nationalId', 'gender'];

  function diffRecords(a, b) {
    var out = [], i, k, av, bv;
    for (i = 0; i < COMPARABLE.length; i++) {
      k = COMPARABLE[i];
      if (b[k] === undefined) continue;           // the import didn't carry this column
      av = a[k] === undefined || a[k] === null ? '' : String(a[k]).trim();
      bv = b[k] === null ? '' : String(b[k]).trim();
      if (av !== bv) out.push({ field: k, from: a[k], to: b[k] });
    }
    return out;
  }

  /**
   * Which of two records for the same person is the more recent? The import
   * timestamp is the only trustworthy signal — a later file wins — but a
   * leaving date or a longer working-day count breaks a tie when both came in
   * during the same batch.
   */
  function whichIsNewer(existing, incoming) {
    var a = Date.parse(existing.importedAt || '') || 0;
    var b = Date.parse(incoming.importedAt || '') || 0;
    if (b > a) return 'incoming';
    if (a > b) return 'existing';
    var wa = Number(existing.workingDays) || 0, wb = Number(incoming.workingDays) || 0;
    if (wb !== wa) return wb > wa ? 'incoming' : 'existing';
    return 'unknown';
  }

  /* ==========================================================================
   * Questionnaire consolidation — the same person may legitimately appear in
   * one file only. Anything else is a duplicate that must be resolved before
   * it can affect a payout.
   * ========================================================================*/
  function detectQuestionnaireDuplicates(records) {
    var byId = {}, i, r;
    for (i = 0; i < records.length; i++) {
      r = records[i];
      (byId[r.employeeId] || (byId[r.employeeId] = [])).push(r);
    }
    var dups = [];
    Object.keys(byId).forEach(function (id) {
      if (byId[id].length > 1) {
        dups.push({
          employeeId: id,
          count: byId[id].length,
          sources: byId[id].map(function (x) { return x.sourceFile || '(نامشخص)'; }),
          records: byId[id]
        });
      }
    });
    return dups;
  }

  /**
   * Apply a duplicate resolution. Every record for the person is marked
   * excluded except the one that wins, so nothing is deleted and the decision
   * stays visible and reversible.
   *
   * @param strategy 'keepFirst' | 'keepLatest' | 'keepSelected' | 'merge'
   */
  function resolveDuplicate(records, employeeId, strategy, selectedIndex) {
    var group = records.filter(function (r) { return r.employeeId === employeeId; });
    if (group.length < 2) return records;

    var keeper;
    if (strategy === 'keepFirst') keeper = group[0];
    else if (strategy === 'keepLatest') {
      keeper = group.reduce(function (best, r) {
        return (Date.parse(r.importedAt || '') || 0) >= (Date.parse(best.importedAt || '') || 0) ? r : best;
      }, group[0]);
    } else if (strategy === 'keepSelected') keeper = group[selectedIndex] || group[0];
    else if (strategy === 'merge') {
      keeper = group[0];
      group.slice(1).forEach(function (r) {
        ['q1', 'q2', 'q3', 'q4', 'q5', 'specialProject', 'specialImpactAmount',
         'jobLevel', 'division', 'positionTitle', 'fullName'].forEach(function (k) {
          if ((keeper[k] === undefined || keeper[k] === null || keeper[k] === '') &&
              r[k] !== undefined && r[k] !== null && r[k] !== '') keeper[k] = r[k];
        });
      });
      keeper.mergedFrom = group.slice(1).map(function (r) { return r.sourceFile; });
    } else keeper = group[0];

    group.forEach(function (r) {
      r.excluded = r !== keeper;
      r.duplicateResolution = strategy;
    });
    return records;
  }

  return {
    emptyState: emptyState,
    read: read,
    write: write,
    clear: clear,
    audit: audit,
    stageEmployees: stageEmployees,
    diffRecords: diffRecords,
    whichIsNewer: whichIsNewer,
    detectQuestionnaireDuplicates: detectQuestionnaireDuplicates,
    resolveDuplicate: resolveDuplicate,
    backend: function () { return backend; }
  };
}));
