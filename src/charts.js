/* ============================================================================
 * charts.js — inline SVG charts for the dashboard
 * ----------------------------------------------------------------------------
 * No charting library: the page must stay a single self-contained file with no
 * network access, so every mark is drawn as SVG here.
 *
 * Design rules followed throughout:
 *   · form first — magnitude gets a bar, part-to-whole gets a stacked bar,
 *     a single ratio against a limit gets a meter; nothing is a pie
 *   · one hue for single-series magnitude (bar length already encodes it);
 *     the reserved status palette only where the categories really are states
 *   · thin marks, rounded data-ends, a 2px surface gap between segments,
 *     recessive axes, selective direct labels — never a number on every mark
 *   · every chart has a hover layer, and the dashboard table below is the
 *     table view that makes the same numbers readable without color
 *
 * Charts are laid out right-to-left to match the rest of the interface.
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Charts = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var NS = 'http://www.w3.org/2000/svg';

  function svgEl(tag, attrs) {
    var n = document.createElementNS(NS, tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      if (attrs[k] !== null && attrs[k] !== undefined) n.setAttribute(k, attrs[k]);
    });
    return n;
  }

  function text(x, y, str, opts) {
    opts = opts || {};
    var t = svgEl('text', {
      x: x, y: y,
      'text-anchor': opts.anchor || 'start',
      'dominant-baseline': opts.baseline || 'middle',
      'font-size': opts.size || 11,
      'font-weight': opts.weight || 400,
      fill: opts.fill || 'var(--chart-ink-2)',
      class: opts.class || null
    });
    t.textContent = str;
    return t;
  }

  /* ------------------------------------------------------------------------
   * Shared tooltip — one node per chart, moved rather than recreated.
   * ---------------------------------------------------------------------- */
  function attachTooltip(container) {
    var tip = document.createElement('div');
    tip.className = 'chart-tip';
    tip.style.display = 'none';
    container.appendChild(tip);

    return {
      show: function (html, evt) {
        tip.innerHTML = html;
        tip.style.display = 'block';
        var box = container.getBoundingClientRect();
        var x = evt.clientX - box.left;
        var y = evt.clientY - box.top;
        /* Keep the tip inside the card rather than letting it clip. */
        var w = tip.offsetWidth, h = tip.offsetHeight;
        tip.style.left = Math.max(4, Math.min(box.width - w - 4, x - w / 2)) + 'px';
        tip.style.top = Math.max(4, y - h - 12) + 'px';
      },
      hide: function () { tip.style.display = 'none'; }
    };
  }

  function frame(container, height) {
    container.classList.add('chart');
    var width = container.clientWidth || 640;
    var svg = svgEl('svg', {
      viewBox: '0 0 ' + width + ' ' + height,
      width: '100%', height: height, role: 'img'
    });
    container.appendChild(svg);
    return { svg: svg, w: width, h: height };
  }

  function niceMax(v) {
    if (v <= 0) return 1;
    var mag = Math.pow(10, Math.floor(Math.log10(v)));
    var n = v / mag;
    var step = n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10;
    return step * mag;
  }

  /* ========================================================================
   * Horizontal bar — magnitude across named categories.
   * Bars grow right-to-left; the label sits on the right, the value on the left.
   * ======================================================================*/
  function horizontalBar(container, rows, opts) {
    opts = opts || {};
    container.innerHTML = '';
    if (!rows.length) { container.appendChild(emptyNote()); return; }

    var labelW = opts.labelWidth || 150;
    var valueW = opts.valueWidth || 96;
    var barH = opts.barHeight || 16;
    var gap = opts.gap || 12;
    var top = 6;
    var height = top + rows.length * (barH + gap);

    var f = frame(container, height);
    var tip = attachTooltip(container);
    var plotRight = f.w - labelW;
    var plotW = Math.max(40, plotRight - valueW);
    var max = niceMax(Math.max.apply(null, rows.map(function (r) { return Math.abs(r.value); })));

    rows.forEach(function (r, i) {
      var y = top + i * (barH + gap);
      var len = max ? Math.max(2, Math.abs(r.value) / max * plotW) : 2;

      /* Track: recessive, shows the full scale behind each bar. */
      f.svg.appendChild(svgEl('rect', {
        x: plotRight - plotW, y: y, width: plotW, height: barH,
        rx: 3, fill: 'var(--chart-track)'
      }));

      var bar = svgEl('rect', {
        x: plotRight - len, y: y, width: len, height: barH,
        rx: 4, fill: r.color || 'var(--chart-series)', class: 'chart-mark'
      });
      f.svg.appendChild(bar);

      /* SVG text neither wraps nor ellipsises, so a long unit name would run
         under the bars. Trim to what the gutter holds and let the tooltip
         carry the full name. */
      f.svg.appendChild(text(f.w - 6, y + barH / 2, fit(r.label, labelW - 10),
        { anchor: 'end', size: 11.5, fill: 'var(--chart-ink)' }));
      f.svg.appendChild(text(plotRight - plotW - 8, y + barH / 2,
        opts.format ? opts.format(r.value) : String(r.value),
        { anchor: 'end', size: 11, weight: 700, fill: 'var(--chart-ink)' }));

      bar.addEventListener('mousemove', function (e) {
        tip.show('<b>' + esc(r.label) + '</b>' +
          (r.detail ? '<br>' + r.detail : '<br>' +
            (opts.format ? opts.format(r.value) : r.value)), e);
      });
      bar.addEventListener('mouseleave', tip.hide);
    });
    return f.svg;
  }

  /* ========================================================================
   * Column chart — distribution across ordered bins, laid out right-to-left.
   * ======================================================================*/
  function columns(container, rows, opts) {
    opts = opts || {};
    container.innerHTML = '';
    if (!rows.length) { container.appendChild(emptyNote()); return; }

    var height = opts.height || 190;
    var padTop = 18, padBottom = 34, padSide = 8;
    var f = frame(container, height);
    var tip = attachTooltip(container);

    var plotH = height - padTop - padBottom;
    var slot = (f.w - padSide * 2) / rows.length;
    var barW = Math.min(opts.maxBarWidth || 46, slot * 0.62);
    var max = niceMax(Math.max.apply(null, rows.map(function (r) { return r.value; })));
    var fmt = opts.format || fmtNum;

    /* Two recessive gridlines are enough to read magnitude. */
    [0.5, 1].forEach(function (frac) {
      var y = padTop + plotH - plotH * frac;
      f.svg.appendChild(svgEl('line', {
        x1: padSide, x2: f.w - padSide, y1: y, y2: y,
        stroke: 'var(--chart-grid)', 'stroke-width': 1
      }));
      f.svg.appendChild(text(f.w - padSide, y - 5, fmt(max * frac),
        { anchor: 'end', size: 10, fill: 'var(--chart-ink-3)' }));
    });

    rows.forEach(function (r, i) {
      /* index 0 on the right — the page reads right-to-left */
      var cx = f.w - padSide - slot * (i + 0.5);
      var h = max ? Math.max(r.value > 0 ? 2 : 0, r.value / max * plotH) : 0;
      var y = padTop + plotH - h;

      if (h > 0) {
        var bar = svgEl('rect', {
          x: cx - barW / 2, y: y, width: barW, height: h,
          rx: 4, fill: r.color || 'var(--chart-series)', class: 'chart-mark'
        });
        f.svg.appendChild(bar);
        bar.addEventListener('mousemove', function (e) {
          tip.show('<b>' + esc(r.label) + '</b><br>' +
            (r.detail || fmt(r.value)), e);
        });
        bar.addEventListener('mouseleave', tip.hide);
      }

      /* Direct-label only the bars with something to say. */
      if (r.value > 0) {
        f.svg.appendChild(text(cx, y - 8, fmt(r.value),
          { anchor: 'middle', size: 10.5, weight: 700, fill: 'var(--chart-ink)' }));
      }
      f.svg.appendChild(text(cx, height - padBottom + 13, r.label,
        { anchor: 'middle', size: 10.5, fill: 'var(--chart-ink-2)' }));
      if (r.sub) {
        f.svg.appendChild(text(cx, height - padBottom + 26, r.sub,
          { anchor: 'middle', size: 9.5, fill: 'var(--chart-ink-3)' }));
      }
    });
    /* Baseline last so marks sit on it. */
    f.svg.appendChild(svgEl('line', {
      x1: padSide, x2: f.w - padSide, y1: padTop + plotH, y2: padTop + plotH,
      stroke: 'var(--chart-axis)', 'stroke-width': 1
    }));
    return f.svg;
  }

  /* ========================================================================
   * Stacked bar — part-to-whole across a handful of states.
   * Segments are separated by a 2px surface gap and always direct-labelled,
   * so identity never rests on color alone.
   * ======================================================================*/
  function stackedBar(container, segments, opts) {
    opts = opts || {};
    container.innerHTML = '';
    var total = segments.reduce(function (a, s) { return a + s.value; }, 0);
    if (!total) { container.appendChild(emptyNote()); return; }

    var barH = opts.barHeight || 30;
    var f = frame(container, barH + 6);
    var tip = attachTooltip(container);
    var GAP = 2;
    var usable = f.w - GAP * Math.max(0, segments.length - 1);
    var x = f.w;

    segments.forEach(function (s, i) {
      if (!s.value) return;
      var w = s.value / total * usable;
      x -= w;
      var rect = svgEl('rect', {
        x: x, y: 3, width: Math.max(1, w), height: barH,
        rx: i === 0 || i === segments.length - 1 ? 4 : 0,
        fill: s.color, class: 'chart-mark'
      });
      f.svg.appendChild(rect);
      /* Label inside the segment only when it genuinely fits. */
      if (w > 44) {
        f.svg.appendChild(text(x + w / 2, 3 + barH / 2, fmtNum(s.value),
          { anchor: 'middle', size: 11.5, weight: 700, fill: '#fff' }));
      }
      rect.addEventListener('mousemove', function (e) {
        tip.show('<b>' + esc(s.label) + '</b><br>' + fmtNum(s.value) + ' نفر · ' +
          (s.value / total * 100).toFixed(1) + '٪', e);
      });
      rect.addEventListener('mouseleave', tip.hide);
      x -= GAP;
    });
    return f.svg;
  }

  /** Legend for the stacked bar — always present, name beside every swatch. */
  function legend(segments, total) {
    var wrap = document.createElement('div');
    wrap.className = 'chart-legend';
    segments.forEach(function (s) {
      if (!s.value && s.hideWhenZero) return;
      var item = document.createElement('span');
      item.className = 'chart-legend-item';
      var sw = document.createElement('i');
      sw.style.background = s.color;
      item.appendChild(sw);
      var label = document.createElement('span');
      label.textContent = s.label + ' — ' + fmtNum(s.value) +
        (total ? ' (' + (s.value / total * 100).toFixed(0) + '٪)' : '');
      item.appendChild(label);
      wrap.appendChild(item);
    });
    return wrap;
  }

  /* ========================================================================
   * Meter — one ratio against a limit.
   * ======================================================================*/
  function meter(container, used, limit, opts) {
    opts = opts || {};
    container.innerHTML = '';
    var f = frame(container, 14);
    var frac = limit ? Math.min(1, used / limit) : 0;
    /* Reconciliation leaves sub-rial dust on a hundred-billion pot, so an exact
       allocation can land a hair above the budget. Anything inside a rial is
       balanced, not an overrun — painting it red contradicted the status chip
       sitting right beside it. */
    var slack = limit ? Math.max(1, Math.abs(limit) * 1e-9) : 0;
    var over = limit && used - limit > slack ? Math.min(1, (used - limit) / limit) : 0;

    f.svg.appendChild(svgEl('rect', {
      x: 0, y: 1, width: f.w, height: 12, rx: 6, fill: 'var(--chart-track)'
    }));
    if (frac > 0) {
      f.svg.appendChild(svgEl('rect', {
        x: f.w - f.w * frac, y: 1, width: f.w * frac, height: 12, rx: 6,
        fill: over ? 'var(--status-critical)' : 'var(--chart-series)'
      }));
    }
    return f.svg;
  }

  /* --------------------------------------------------------------- helpers */
  function emptyNote() {
    var d = document.createElement('div');
    d.className = 'chart-empty';
    d.textContent = 'داده‌ای برای نمایش وجود ندارد.';
    return d;
  }

  /** Approximate character budget for an 11.5px Persian label. */
  function fit(label, pixels) {
    var max = Math.max(4, Math.floor(pixels / 6.6));
    var s = String(label);
    return s.length <= max ? s : s.slice(0, max - 1) + '…';
  }

  function fmtNum(v) {
    return Number(v).toLocaleString('en-US', { maximumFractionDigits: 0 });
  }

  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/[&<>"]/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
      });
  }

  /* The reserved status palette. Validated as a set against both surfaces;
     warning sits below 3:1 on light by design, which is why every segment
     carries a visible label. */
  var STATUS = {
    good:     '#0ca30c',
    neutral:  '#2a78d6',
    warning:  '#fab219',
    critical: '#d03b3b'
  };

  return {
    horizontalBar: horizontalBar,
    columns: columns,
    stackedBar: stackedBar,
    legend: legend,
    meter: meter,
    STATUS: STATUS
  };
}));
