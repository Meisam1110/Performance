/* ============================================================================
 * ui.js — rendering primitives for the Karaneh Management System
 * ----------------------------------------------------------------------------
 * Formatting, DOM helpers, the sortable/filterable data grid, modals and
 * toasts. No business logic lives here: everything numeric arrives already
 * computed by calculation-engine.js.
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.UI = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ======================================================================
   * Formatting — mirrors the number formats used in the reference workbook.
   * ====================================================================*/

  /** Thousand-separated integer rial, e.g. 4,055,789,474. */
  function money(v, decimals) {
    if (v === null || v === undefined || v === '' || !isFinite(v)) return '—';
    var d = decimals === undefined ? 0 : decimals;
    var s = Math.abs(v).toLocaleString('en-US', {
      minimumFractionDigits: d, maximumFractionDigits: d
    });
    /* Budget reconciliation leaves sub-rial dust on a 100-billion pot. Showing
       it as "−0" reads like an overrun, so a rounded-away value loses its sign. */
    var isZero = !/[1-9]/.test(s);
    return (v < 0 && !isZero ? '−' : '') + s;
  }

  /** Compact rial for KPI tiles: 97.4B / 250.0M. */
  function moneyShort(v) {
    if (v === null || v === undefined || !isFinite(v)) return '—';
    var a = Math.abs(v), body;
    if (a >= 1e12)      body = (a / 1e12).toFixed(2) + 'T';
    else if (a >= 1e9)  body = (a / 1e9).toFixed(2) + 'B';
    else if (a >= 1e6)  body = (a / 1e6).toFixed(1) + 'M';
    else if (a >= 1e3)  body = (a / 1e3).toFixed(1) + 'K';
    else                body = a.toFixed(0);
    /* Sub-rial reconciliation dust must not surface as "−0". */
    return (v < 0 && /[1-9]/.test(body) ? '−' : '') + body;
  }

  function score(v, decimals) {
    if (v === null || v === undefined || v === '' || !isFinite(v)) return '—';
    return Number(v).toFixed(decimals === undefined ? 2 : decimals);
  }

  function percent(v, decimals) {
    if (v === null || v === undefined || !isFinite(v)) return '—';
    return (v * 100).toFixed(decimals === undefined ? 1 : decimals) + '٪';
  }

  function int(v) {
    if (v === null || v === undefined || !isFinite(v)) return '—';
    return Math.round(v).toLocaleString('en-US');
  }

  function dateTime(iso) {
    if (!iso) return '—';
    var d = new Date(iso);
    if (isNaN(d)) return String(iso);
    function p(n) { return n < 10 ? '0' + n : String(n); }
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) +
           ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
  }

  function esc(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /* ======================================================================
   * DOM helpers
   * ====================================================================*/
  function el(tag, attrs, children) {
    var node = document.createElement(tag), k;
    if (attrs) for (k in attrs) {
      if (!attrs.hasOwnProperty(k) || attrs[k] === null || attrs[k] === undefined) continue;
      if (k === 'class') node.className = attrs[k];
      else if (k === 'html') node.innerHTML = attrs[k];
      else if (k === 'text') node.textContent = attrs[k];
      else if (k.indexOf('on') === 0 && typeof attrs[k] === 'function') {
        node.addEventListener(k.slice(2).toLowerCase(), attrs[k]);
      } else if (k === 'dataset') {
        Object.keys(attrs[k]).forEach(function (d) { node.dataset[d] = attrs[k][d]; });
      } else node.setAttribute(k, attrs[k]);
    }
    (children || []).forEach(function (c) {
      if (c === null || c === undefined || c === false) return;
      node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return node;
  }

  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); return node; }
  function $(sel, ctx) { return (ctx || document).querySelector(sel); }
  function $$(sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); }

  /* ======================================================================
   * Toast
   * ====================================================================*/
  function toast(message, kind, ms) {
    var host = $('.toast-host');
    if (!host) { host = el('div', { class: 'toast-host' }); document.body.appendChild(host); }
    var node = el('div', { class: 'toast ' + (kind || ''), text: message });
    host.appendChild(node);
    setTimeout(function () {
      node.style.transition = 'opacity .3s'; node.style.opacity = '0';
      setTimeout(function () { if (node.parentNode) node.parentNode.removeChild(node); }, 320);
    }, ms || 3800);
  }

  /* ======================================================================
   * Modal
   * ====================================================================*/
  function modal(opts) {
    var backdrop = el('div', { class: 'modal-backdrop' });
    var box = el('div', { class: 'modal ' + (opts.size || '') });

    var header = el('header', {}, [
      el('span', { text: opts.title || '' }),
      el('button', {
        class: 'btn ghost close', text: '✕ بستن',
        onclick: function () { close(); }
      })
    ]);
    var content = el('div', { class: 'content' });
    if (typeof opts.content === 'string') content.innerHTML = opts.content;
    else if (opts.content) content.appendChild(opts.content);

    box.appendChild(header);
    box.appendChild(content);

    if (opts.buttons && opts.buttons.length) {
      var footer = el('footer', {});
      opts.buttons.forEach(function (b) {
        if (b === 'spacer') { footer.appendChild(el('div', { class: 'spacer' })); return; }
        footer.appendChild(el('button', {
          class: 'btn ' + (b.kind || ''), text: b.label,
          disabled: b.disabled ? 'disabled' : null,
          onclick: function () { if (b.onClick && b.onClick(close) === false) return; if (b.keepOpen !== true) close(); }
        }));
      });
      box.appendChild(footer);
    }

    backdrop.appendChild(box);
    backdrop.addEventListener('mousedown', function (e) {
      if (e.target === backdrop && opts.dismissable !== false) close();
    });
    document.body.appendChild(backdrop);

    function onKey(e) { if (e.key === 'Escape' && opts.dismissable !== false) close(); }
    document.addEventListener('keydown', onKey);

    function close() {
      document.removeEventListener('keydown', onKey);
      if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
      if (opts.onClose) opts.onClose();
    }
    return { close: close, content: content, box: box };
  }

  function confirm(message, opts) {
    opts = opts || {};
    return new Promise(function (resolve) {
      modal({
        title: opts.title || 'تأیید',
        size: 'narrow',
        content: el('div', { html: '<p style="margin:0">' + esc(message).replace(/\n/g, '<br>') + '</p>' }),
        buttons: [
          { label: opts.confirmLabel || 'تأیید', kind: opts.danger ? 'danger' : 'primary', onClick: function () { resolve(true); } },
          { label: 'انصراف', onClick: function () { resolve(false); } }
        ],
        onClose: function () { resolve(false); }
      });
    });
  }

  /* ======================================================================
   * DataGrid — sortable, filterable, column-toggleable table.
   *
   * columns: [{ key, label, type, width, group, hidden, editable,
   *             value(row), render(row), className(row), sortValue(row) }]
   * type drives alignment and formatting: 'money' | 'score' | 'int' | 'text'
   * ====================================================================*/
  function DataGrid(opts) {
    var state = {
      columns: opts.columns,
      rows: opts.rows || [],
      sortKey: opts.sortKey || null,
      sortDir: opts.sortDir || 'asc',
      filter: '',
      facets: {},                      // key -> selected value
      hidden: {},
      page: 0,
      pageSize: opts.pageSize || 250
    };
    opts.columns.forEach(function (c) { if (c.hidden) state.hidden[c.key] = true; });

    var root = el('div', { class: 'card' });
    var head = el('h2', {}, [
      el('span', { text: opts.title || '' }),
      el('span', { class: 'hint', text: '' }),
      el('span', { class: 'right' })
    ]);
    var toolbar = el('div', { class: 'table-toolbar' });
    var wrap = el('div', { class: 'table-wrap' });
    var table = el('table', { class: 'grid' });
    wrap.appendChild(table);
    root.appendChild(head);
    root.appendChild(toolbar);
    root.appendChild(wrap);

    /* -- toolbar ------------------------------------------------------- */
    var search = el('input', {
      type: 'search', placeholder: 'جستجو در نام، شماره پرسنلی، واحد…',
      style: 'min-width:250px', oninput: function () { state.filter = this.value; state.page = 0; render(); }
    });
    toolbar.appendChild(search);

    (opts.facets || []).forEach(function (f) {
      var sel = el('select', {
        onchange: function () { state.facets[f.key] = this.value; state.page = 0; render(); }
      }, [el('option', { value: '', text: f.label })]);
      sel.dataset.facet = f.key;
      toolbar.appendChild(sel);
    });

    toolbar.appendChild(el('div', { class: 'grow' }));
    (opts.actions || []).forEach(function (a) {
      toolbar.appendChild(el('button', { class: 'btn sm ' + (a.kind || ''), text: a.label, onclick: a.onClick }));
    });
    toolbar.appendChild(el('button', {
      class: 'btn sm', text: '⚙ ستون‌ها', onclick: function () { columnPicker(); }
    }));

    function columnPicker() {
      var body = el('div', {});
      state.columns.forEach(function (c) {
        if (c.alwaysVisible) return;
        var cb = el('input', { type: 'checkbox' });
        cb.checked = !state.hidden[c.key];
        cb.addEventListener('change', function () {
          state.hidden[c.key] = !cb.checked; render();
        });
        body.appendChild(el('label', { class: 'checkline' }, [cb, el('span', { text: c.label })]));
      });
      modal({ title: 'نمایش ستون‌ها', size: 'narrow', content: body, buttons: [{ label: 'بستن', kind: 'primary' }] });
    }

    /* -- data ---------------------------------------------------------- */
    function cellValue(col, row) {
      return col.value ? col.value(row) : row[col.key];
    }

    function visibleColumns() {
      return state.columns.filter(function (c) { return !state.hidden[c.key]; });
    }

    function filtered() {
      var q = state.filter.trim().toLowerCase();
      var out = state.rows.filter(function (r) {
        var ok = true;
        Object.keys(state.facets).forEach(function (k) {
          var want = state.facets[k];
          if (want === '' || want === undefined) return;
          var f = (opts.facets || []).filter(function (x) { return x.key === k; })[0];
          var got = f && f.value ? f.value(r) : r[k];
          if (String(got === undefined || got === null ? '' : got) !== want) ok = false;
        });
        if (!ok) return false;
        if (!q) return true;
        var hay = (opts.searchFields || ['employeeId', 'fullName', 'division', 'positionTitle'])
          .map(function (k) { return r[k] === undefined || r[k] === null ? '' : r[k]; })
          .join(' ').toLowerCase();
        return hay.indexOf(q) !== -1;
      });

      if (state.sortKey) {
        var col = state.columns.filter(function (c) { return c.key === state.sortKey; })[0];
        if (col) {
          var dir = state.sortDir === 'asc' ? 1 : -1;
          out = out.slice().sort(function (a, b) {
            var av = col.sortValue ? col.sortValue(a) : cellValue(col, a);
            var bv = col.sortValue ? col.sortValue(b) : cellValue(col, b);
            if (av === null || av === undefined) av = col.type === 'text' ? '' : -Infinity;
            if (bv === null || bv === undefined) bv = col.type === 'text' ? '' : -Infinity;
            if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
            return naturalCompare(String(av), String(bv)) * dir;
          });
        }
      }
      return out;
    }

    function formatCell(col, row) {
      if (col.render) return col.render(row);
      var v = cellValue(col, row);
      if (col.type === 'money') return money(v, col.decimals);
      if (col.type === 'score') return score(v, col.decimals);
      if (col.type === 'int')   return int(v);
      if (col.type === 'percent') return percent(v, col.decimals);
      return v === null || v === undefined || v === '' ? '—' : String(v);
    }

    /* -- render -------------------------------------------------------- */
    function render() {
      var rows = filtered();
      var cols = visibleColumns();

      head.querySelector('.hint').textContent =
        rows.length === state.rows.length
          ? state.rows.length.toLocaleString('en-US') + ' ردیف'
          : rows.length.toLocaleString('en-US') + ' از ' + state.rows.length.toLocaleString('en-US') + ' ردیف';

      /* facet options are rebuilt from the data so they never go stale */
      (opts.facets || []).forEach(function (f) {
        var sel = toolbar.querySelector('select[data-facet="' + f.key + '"]');
        if (!sel) return;
        var seen = {}, values = [];
        state.rows.forEach(function (r) {
          var v = f.value ? f.value(r) : r[f.key];
          if (v === null || v === undefined || v === '') return;
          v = String(v);
          if (!seen[v]) { seen[v] = 1; values.push(v); }
        });
        values.sort(function (a, b) { return a.localeCompare(b, 'fa'); });
        var current = state.facets[f.key] || '';
        clear(sel);
        sel.appendChild(el('option', { value: '', text: f.label }));
        values.forEach(function (v) { sel.appendChild(el('option', { value: v, text: v })); });
        sel.value = current;
      });

      clear(table);

      /* grouped header band, mirroring the merged title row in the workbook */
      var groups = [], last = null;
      cols.forEach(function (c) {
        var g = c.group || '';
        if (last && last.name === g) last.span++;
        else { last = { name: g, span: 1 }; groups.push(last); }
      });
      var thead = el('thead');
      if (groups.filter(function (g) { return g.name; }).length) {
        var gr = el('tr');
        groups.forEach(function (g) {
          gr.appendChild(el('th', {
            class: g.name ? 'group-head' : '', colspan: g.span, text: g.name
          }));
        });
        thead.appendChild(gr);
      }
      var hr = el('tr');
      cols.forEach(function (c) {
        var th = el('th', {
          class: (state.sortKey === c.key ? 'sorted ' : '') + (c.type && c.type !== 'text' ? 'right' : ''),
          title: c.title || c.label,
          onclick: function () {
            if (state.sortKey === c.key) state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
            else { state.sortKey = c.key; state.sortDir = c.type && c.type !== 'text' ? 'desc' : 'asc'; }
            render();
          }
        }, [
          document.createTextNode(c.label),
          el('span', { class: 'sort', text: state.sortKey === c.key ? (state.sortDir === 'asc' ? '▲' : '▼') : '↕' })
        ]);
        if (c.width) th.style.minWidth = c.width;
        hr.appendChild(th);
      });
      thead.appendChild(hr);
      table.appendChild(thead);

      var tbody = el('tbody');
      var slice = rows.slice(state.page * state.pageSize, (state.page + 1) * state.pageSize);
      if (!slice.length) {
        tbody.appendChild(el('tr', {}, [
          el('td', { colspan: cols.length, class: 'empty', text: 'ردیفی برای نمایش وجود ندارد.' })
        ]));
      }
      slice.forEach(function (r) {
        var cls = opts.rowClass ? (opts.rowClass(r) || '') : '';
        var tr = el('tr', { class: cls });
        if (opts.onRowClick) {
          tr.style.cursor = 'pointer';
          tr.addEventListener('click', function (e) {
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' ||
                e.target.tagName === 'BUTTON' || e.target.tagName === 'TEXTAREA') return;
            opts.onRowClick(r);
          });
        }
        cols.forEach(function (c) {
          var td = el('td', {
            class: [(c.type && c.type !== 'text') ? 'num' : '',
                    c.editable ? '' : (c.calculated ? 'calc' : ''),
                    c.className ? (c.className(r) || '') : ''].join(' ').trim()
          });
          var out = formatCell(c, r);
          if (out && out.nodeType) td.appendChild(out);
          else td.textContent = out;
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);

      if (opts.footer) {
        var tf = el('tfoot'), fr = el('tr');
        var totals = opts.footer(rows);
        cols.forEach(function (c) {
          var v = totals[c.key];
          fr.appendChild(el('td', {
            class: (c.type && c.type !== 'text') ? 'num' : '',
            text: v === undefined ? '' : v
          }));
        });
        tf.appendChild(fr);
        table.appendChild(tf);
      }

      /* pagination */
      var pageBar = root.querySelector('.pagebar');
      if (pageBar) pageBar.parentNode.removeChild(pageBar);
      if (rows.length > state.pageSize) {
        var pages = Math.ceil(rows.length / state.pageSize);
        var bar = el('div', { class: 'table-toolbar pagebar' }, [
          el('button', {
            class: 'btn sm', text: '‹ قبلی', disabled: state.page === 0 ? 'disabled' : null,
            onclick: function () { state.page--; render(); wrap.scrollTop = 0; }
          }),
          el('span', { class: 'small', text: 'صفحه ' + (state.page + 1) + ' از ' + pages }),
          el('button', {
            class: 'btn sm', text: 'بعدی ›', disabled: state.page >= pages - 1 ? 'disabled' : null,
            onclick: function () { state.page++; render(); wrap.scrollTop = 0; }
          })
        ]);
        root.appendChild(bar);
      }
    }

    render();

    return {
      node: root,
      setRows: function (rows) {
        state.rows = rows;
        var maxPage = Math.max(0, Math.ceil(rows.length / state.pageSize) - 1);
        if (state.page > maxPage) state.page = maxPage;
        render();
      },
      getVisibleRows: filtered,
      getColumns: visibleColumns,
      render: render,
      state: state
    };
  }

  /**
   * Compare strings with embedded numbers the way a person would.
   * Employee numbers are text (they may carry leading zeros or a prefix such
   * as BEKI001), so a plain string sort puts 10 before 2. This walks the
   * digit runs numerically and everything else with a Persian collation.
   */
  var NUM_CHUNK = /(\d+)|(\D+)/g;
  function naturalCompare(a, b) {
    if (a === b) return 0;
    var ax = a.match(NUM_CHUNK) || [], bx = b.match(NUM_CHUNK) || [];
    for (var i = 0; i < Math.min(ax.length, bx.length); i++) {
      var an = /^\d/.test(ax[i]), bn = /^\d/.test(bx[i]);
      if (an && bn) {
        var d = parseInt(ax[i], 10) - parseInt(bx[i], 10);
        if (d) return d < 0 ? -1 : 1;
      } else {
        var c = ax[i].localeCompare(bx[i], 'fa');
        if (c) return c;
      }
    }
    return ax.length - bx.length;
  }

  /* ======================================================================
   * KPI tile
   * ====================================================================*/
  function kpi(label, value, opts) {
    opts = opts || {};
    return el('div', { class: 'kpi ' + (opts.kind || ''), title: opts.title || '' }, [
      el('div', { class: 'label', text: label }),
      el('div', { class: 'value' }, [bidi(value)]),
      opts.sub ? el('div', { class: 'sub', text: opts.sub }) : null
    ]);
  }

  /**
   * Wrap a value in <bdi> so it is bidi-isolated from the surrounding Persian.
   * Without this a leading minus on a rial figure detaches and renders after
   * the digits, turning −237,651,387 into 237,651,387−.
   */
  function bidi(value) {
    return el('bdi', { text: value === null || value === undefined ? '' : String(value) });
  }

  function alert(kind, title, body, action) {
    var icons = { err: '⛔', warn: '⚠️', ok: '✅', info: 'ℹ️' };
    return el('div', { class: 'alert ' + kind }, [
      el('span', { class: 'ico', text: icons[kind] || 'ℹ️' }),
      el('div', {}, [
        title ? el('b', { text: title }) : null,
        el('span', { html: typeof body === 'string' ? body : '' })
      ].concat(typeof body === 'object' && body ? [body] : [])),
      action ? el('div', { class: 'act' }, [action]) : null
    ]);
  }

  function card(title, bodyNode, opts) {
    opts = opts || {};
    var h = el('h2', {}, [
      el('span', { text: title }),
      opts.hint ? el('span', { class: 'hint', text: opts.hint }) : null,
      opts.right ? el('span', { class: 'right' }, opts.right) : null
    ]);
    var b = el('div', { class: 'body ' + (opts.tight ? 'tight' : '') });
    if (bodyNode) b.appendChild(bodyNode);
    return el('div', { class: 'card' }, [h, b]);
  }

  return {
    money: money, moneyShort: moneyShort, score: score, percent: percent,
    int: int, dateTime: dateTime, esc: esc,
    el: el, clear: clear, $: $, $$: $$,
    toast: toast, modal: modal, confirm: confirm,
    DataGrid: DataGrid, kpi: kpi, alert: alert, card: card, bidi: bidi,
    naturalCompare: naturalCompare
  };
}));
