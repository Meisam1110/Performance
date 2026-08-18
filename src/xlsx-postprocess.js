/* ============================================================================
 * xlsx-postprocess.js — freeze panes and cell formatting for SheetJS output
 * ----------------------------------------------------------------------------
 * The community build of SheetJS writes column widths, autofilters and RTL
 * sheet views, but neither frozen panes nor cell styles: there is no `!freeze`
 * handling anywhere in the bundle, and `cell.z` never reaches styles.xml.
 * Frozen headers and thousand-separated rial amounts both matter on a payment
 * sheet that runs to thousands of rows, so this module edits the generated
 * .xlsx directly.
 *
 * An .xlsx is a ZIP. When SheetJS is asked to write without compression every
 * entry is STORED, which means each file's bytes sit verbatim in the archive
 * and can be replaced in place. This module rewrites styles.xml and the
 * worksheet parts, then rebuilds the ZIP with corrected sizes, CRCs and
 * offsets.
 *
 * It is deliberately conservative: any structure it does not recognise makes
 * it return the original bytes untouched, so a failure here can never corrupt
 * an export.
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XlsxPostprocess = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ------------------------------------------------------------------ CRC32 */
  var CRC_TABLE = (function () {
    var table = new Int32Array(256), c, n, k;
    for (n = 0; n < 256; n++) {
      c = n;
      for (k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c;
    }
    return table;
  }());

  function crc32(bytes) {
    var c = -1;
    for (var i = 0; i < bytes.length; i++) {
      c = (c >>> 8) ^ CRC_TABLE[(c ^ bytes[i]) & 0xFF];
    }
    return (c ^ -1) >>> 0;
  }

  /* -------------------------------------------------------------- utilities */
  var utf8Decode, utf8Encode;
  if (typeof TextDecoder !== 'undefined') {
    var dec = new TextDecoder('utf-8'), enc = new TextEncoder();
    utf8Decode = function (b) { return dec.decode(b); };
    utf8Encode = function (s) { return enc.encode(s); };
  } else {
    utf8Decode = function (b) { return Buffer.from(b).toString('utf8'); };
    utf8Encode = function (s) { return new Uint8Array(Buffer.from(s, 'utf8')); };
  }

  function u16(b, o) { return b[o] | (b[o + 1] << 8); }
  function u32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }
  function w16(b, o, v) { b[o] = v & 0xFF; b[o + 1] = (v >>> 8) & 0xFF; }
  function w32(b, o, v) {
    b[o] = v & 0xFF; b[o + 1] = (v >>> 8) & 0xFF;
    b[o + 2] = (v >>> 16) & 0xFF; b[o + 3] = (v >>> 24) & 0xFF;
  }

  var SIG_LOCAL   = 0x04034B50;
  var SIG_CENTRAL = 0x02014B50;
  var SIG_EOCD    = 0x06054B50;

  /**
   * Read the archive into a list of entries. Returns null for anything this
   * module is not prepared to rewrite safely — a compressed entry, a Zip64
   * archive, an entry using a trailing data descriptor.
   */
  function readZip(bytes) {
    /* The end-of-central-directory record sits in the last 64KB. */
    var eocd = -1;
    for (var i = bytes.length - 22; i >= 0 && i >= bytes.length - 65558; i--) {
      if (u32(bytes, i) === SIG_EOCD) { eocd = i; break; }
    }
    if (eocd < 0) return null;

    var count = u16(bytes, eocd + 10);
    var cdOffset = u32(bytes, eocd + 16);
    if (cdOffset === 0xFFFFFFFF) return null;          // Zip64

    var entries = [], p = cdOffset;
    for (var n = 0; n < count; n++) {
      if (u32(bytes, p) !== SIG_CENTRAL) return null;
      var method   = u16(bytes, p + 10);
      var flags    = u16(bytes, p + 8);
      var nameLen  = u16(bytes, p + 28);
      var extraLen = u16(bytes, p + 30);
      var cmtLen   = u16(bytes, p + 32);
      var local    = u32(bytes, p + 42);
      if (method !== 0) return null;                   // only STORED
      if (flags & 0x08) return null;                   // data descriptor
      if (u32(bytes, local) !== SIG_LOCAL) return null;

      var lNameLen  = u16(bytes, local + 26);
      var lExtraLen = u16(bytes, local + 28);
      var dataStart = local + 30 + lNameLen + lExtraLen;
      var size      = u32(bytes, p + 24);

      entries.push({
        name: utf8Decode(bytes.subarray(p + 46, p + 46 + nameLen)),
        central: bytes.subarray(p, p + 46 + nameLen + extraLen + cmtLen),
        localHeader: bytes.subarray(local, local + 30 + lNameLen + lExtraLen),
        data: bytes.subarray(dataStart, dataStart + size)
      });
      p += 46 + nameLen + extraLen + cmtLen;
    }
    return { entries: entries, eocd: bytes.subarray(eocd, bytes.length) };
  }

  /** Rebuild the archive from (possibly edited) entries. */
  function writeZip(zip) {
    var total = 0, i, e;
    for (i = 0; i < zip.entries.length; i++) {
      e = zip.entries[i];
      total += e.localHeader.length + e.data.length + e.central.length;
    }
    total += zip.eocd.length;

    var out = new Uint8Array(total), pos = 0, cdStart, offsets = [];
    for (i = 0; i < zip.entries.length; i++) {
      e = zip.entries[i];
      offsets.push(pos);
      var lh = e.localHeader.slice();
      w32(lh, 14, e.crc); w32(lh, 18, e.data.length); w32(lh, 22, e.data.length);
      out.set(lh, pos); pos += lh.length;
      out.set(e.data, pos); pos += e.data.length;
    }
    cdStart = pos;
    for (i = 0; i < zip.entries.length; i++) {
      e = zip.entries[i];
      var cd = e.central.slice();
      w32(cd, 16, e.crc); w32(cd, 20, e.data.length); w32(cd, 24, e.data.length);
      w32(cd, 42, offsets[i]);
      out.set(cd, pos); pos += cd.length;
    }
    var eocd = zip.eocd.slice();
    w32(eocd, 12, pos - cdStart);
    w32(eocd, 16, cdStart);
    out.set(eocd, pos); pos += eocd.length;
    return out.subarray(0, pos);
  }

  /* ------------------------------------------------------- worksheet lookup */

  /**
   * Map sheet display names onto their worksheet parts, via workbook.xml and
   * its relationships, rather than assuming sheet order matches sheetN.xml.
   */
  function sheetPartsByName(entries) {
    var byName = {}, wbXml = null, relXml = null;
    entries.forEach(function (e) {
      if (e.name === 'xl/workbook.xml') wbXml = utf8Decode(e.data);
      else if (e.name === 'xl/_rels/workbook.xml.rels') relXml = utf8Decode(e.data);
    });
    if (!wbXml || !relXml) return byName;

    var rels = {}, m;
    var relRe = /<Relationship\b[^>]*>/g;
    while ((m = relRe.exec(relXml))) {
      var id = /Id="([^"]+)"/.exec(m[0]);
      var target = /Target="([^"]+)"/.exec(m[0]);
      if (id && target) {
        var t = target[1].replace(/^\/?xl\//, '').replace(/^\.\//, '');
        rels[id[1]] = 'xl/' + t;
      }
    }

    var sheetRe = /<sheet\b[^>]*>/g;
    while ((m = sheetRe.exec(wbXml))) {
      var name = /name="([^"]*)"/.exec(m[0]);
      var rid = /r:id="([^"]+)"/.exec(m[0]);
      if (name && rid && rels[rid[1]]) byName[unescapeXml(name[1])] = rels[rid[1]];
    }
    return byName;
  }

  function unescapeXml(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }

  function colLetter(i) {
    var s = ''; i += 1;
    while (i > 0) { var r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }

  /**
   * Insert a frozen pane into one worksheet's XML.
   * @param xml  the worksheet part
   * @param x    number of columns to freeze (0 = none)
   * @param y    number of rows to freeze (0 = none)
   */
  function addPane(xml, x, y) {
    if (!x && !y) return xml;
    var topLeft = colLetter(x) + (y + 1);
    var active = x && y ? 'bottomRight' : (x ? 'topRight' : 'bottomLeft');
    var pane = '<pane' +
      (x ? ' xSplit="' + x + '"' : '') +
      (y ? ' ySplit="' + y + '"' : '') +
      ' topLeftCell="' + topLeft + '" activePane="' + active + '" state="frozen"/>' +
      '<selection pane="' + active + '" activeCell="' + topLeft + '" sqref="' + topLeft + '"/>';

    /* Self-closing sheetView: <sheetView .../> → <sheetView ...>pane</sheetView> */
    var selfClosing = /<sheetView\b([^>]*?)\/>/;
    if (selfClosing.test(xml)) {
      return xml.replace(selfClosing, function (_, attrs) {
        return '<sheetView' + attrs + '>' + pane + '</sheetView>';
      });
    }
    /* Open tag already present: insert the pane as its first child. */
    var openTag = /<sheetView\b([^>]*)>/;
    if (openTag.test(xml)) {
      return xml.replace(openTag, function (whole) { return whole + pane; });
    }
    /* No sheetViews at all: add one before <sheetData>. */
    if (xml.indexOf('<sheetData') !== -1) {
      return xml.replace('<sheetData',
        '<sheetViews><sheetView workbookViewId="0">' + pane + '</sheetView></sheetViews><sheetData');
    }
    return xml;
  }

  /* ------------------------------------------------------------- styling */

  /**
   * Append custom number formats, a bold header font and a header fill to the
   * stylesheet SheetJS emitted.
   *
   * @returns {{xml:string, headerStyle:number, formatStyle:Object}}
   *          formatStyle maps a format code onto its cellXfs index.
   */
  function extendStyles(xml, formatCodes) {
    var FIRST_CUSTOM_FMT = 164;   // ids below this are reserved by the format

    /* How many cellXfs already exist — new entries are appended after them. */
    var xfsMatch = /<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/.exec(xml);
    if (!xfsMatch) return null;
    var xfCount = parseInt(xfsMatch[1], 10);
    var xfBody = xfsMatch[2];

    /* Highest numFmtId in use, so custom ids cannot collide. */
    var nextFmtId = FIRST_CUSTOM_FMT, m;
    var fmtIdRe = /numFmtId="(\d+)"/g;
    while ((m = fmtIdRe.exec(xml))) {
      var id = parseInt(m[1], 10);
      if (id >= nextFmtId) nextFmtId = id + 1;
    }

    var newFmts = '', formatStyle = {}, addedXfs = '';
    formatCodes.forEach(function (code) {
      var fmtId = nextFmtId++;
      newFmts += '<numFmt numFmtId="' + fmtId + '" formatCode="' + escapeXmlAttr(code) + '"/>';
      formatStyle[code] = xfCount + 1 + Object.keys(formatStyle).length;
      addedXfs += '<xf numFmtId="' + fmtId + '" fontId="0" fillId="0" borderId="0" xfId="0"' +
                  ' applyNumberFormat="1"/>';
    });

    /* The header style occupies the first appended slot, the number formats
       follow it — matching the order the entries are concatenated below. */
    var headerStyle = xfCount;

    var fontsMatch = /<fonts count="(\d+)">/.exec(xml);
    var fillsMatch = /<fills count="(\d+)">/.exec(xml);
    if (!fontsMatch || !fillsMatch) return null;
    var headerFontId = parseInt(fontsMatch[1], 10);
    var headerFillId = parseInt(fillsMatch[1], 10);

    var headerXf = '<xf numFmtId="0" fontId="' + headerFontId + '" fillId="' + headerFillId +
      '" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>';

    var out = xml;

    /* numFmts may be absent entirely on a workbook with no dates. */
    if (/<numFmts count="(\d+)">/.test(out)) {
      out = out.replace(/<numFmts count="(\d+)">/, function (_, c) {
        return '<numFmts count="' + (parseInt(c, 10) + formatCodes.length) + '">';
      }).replace('</numFmts>', newFmts + '</numFmts>');
    } else if (formatCodes.length) {
      out = out.replace(/(<styleSheet[^>]*>)/,
        '$1<numFmts count="' + formatCodes.length + '">' + newFmts + '</numFmts>');
    }

    out = out.replace(/<fonts count="(\d+)">/, function (_, c) {
      return '<fonts count="' + (parseInt(c, 10) + 1) + '">';
    }).replace('</fonts>',
      '<font><b/><sz val="12"/><color theme="1"/><name val="Calibri"/>' +
      '<family val="2"/><scheme val="minor"/></font></fonts>');

    out = out.replace(/<fills count="(\d+)">/, function (_, c) {
      return '<fills count="' + (parseInt(c, 10) + 1) + '">';
    }).replace('</fills>',
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEEF2F7"/>' +
      '<bgColor indexed="64"/></patternFill></fill></fills>');

    out = out.replace(/<cellXfs count="\d+">[\s\S]*?<\/cellXfs>/,
      '<cellXfs count="' + (xfCount + 1 + formatCodes.length) + '">' +
      xfBody + headerXf + addedXfs + '</cellXfs>');

    return { xml: out, headerStyle: headerStyle, formatStyle: formatStyle };
  }

  function escapeXmlAttr(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /**
   * Stamp style indices onto worksheet cells.
   *
   * @param xml          the worksheet part
   * @param headerRow    1-based row to render as a header (0 = none)
   * @param headerStyle  cellXfs index for the header style
   * @param columnStyle  column letter → cellXfs index, applied below the header
   */
  function styleCells(xml, headerRow, headerStyle, columnStyle) {
    var hasColumnStyles = Object.keys(columnStyle).length > 0;
    if (!headerRow && !hasColumnStyles) return xml;

    return xml.replace(/<c r="([A-Z]+)(\d+)"([^>]*)(\/?)>/g,
      function (whole, col, row, attrs, selfClose) {
        if (attrs.indexOf(' s="') !== -1) return whole;   // already styled
        var rowNum = parseInt(row, 10), style;
        if (headerRow && rowNum === headerRow) style = headerStyle;
        else if (rowNum > headerRow && columnStyle[col] !== undefined) style = columnStyle[col];
        if (style === undefined) return whole;
        return '<c r="' + col + row + '" s="' + style + '"' + attrs + selfClose + '>';
      });
  }

  /**
   * Apply frozen panes, header styling and number formats to an .xlsx that
   * SheetJS produced with compression disabled.
   *
   * @param {Uint8Array|ArrayBuffer} buffer
   * @param {Object} specBySheet  sheet display name → {
   *          xSplit, ySplit,          frozen columns / rows
   *          headerRow,               1-based header row to embolden
   *          numberFormats            { 'M': '#,##0', 'F': '0.00', ... }
   *        }
   * @returns {Uint8Array} the patched workbook, or the input unchanged if the
   *          archive is not in the simple STORED form this module can rewrite.
   *          A failure here can never corrupt the export.
   */
  function applyFormatting(buffer, specBySheet) {
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    var zip;
    try { zip = readZip(bytes); } catch (e) { return bytes; }
    if (!zip) return bytes;

    /* Collect every distinct format code used across all sheets. */
    var codes = [], seen = {};
    Object.keys(specBySheet).forEach(function (name) {
      var fmts = specBySheet[name].numberFormats || {};
      Object.keys(fmts).forEach(function (col) {
        if (!seen[fmts[col]]) { seen[fmts[col]] = 1; codes.push(fmts[col]); }
      });
    });

    var styleEntry = null;
    zip.entries.forEach(function (e) { if (e.name === 'xl/styles.xml') styleEntry = e; });

    var styles = null;
    if (styleEntry) {
      try { styles = extendStyles(utf8Decode(styleEntry.data), codes); }
      catch (err) { styles = null; }
    }

    var parts = sheetPartsByName(zip.entries);
    var touched = false;

    zip.entries.forEach(function (e) { e.crc = crc32(e.data); });

    if (styles) {
      styleEntry.data = utf8Encode(styles.xml);
      styleEntry.crc = crc32(styleEntry.data);
      touched = true;
    }

    Object.keys(specBySheet).forEach(function (sheetName) {
      var part = parts[sheetName];
      if (!part) return;
      var spec = specBySheet[sheetName];

      zip.entries.forEach(function (e) {
        if (e.name !== part) return;
        var xml = utf8Decode(e.data), before = xml;

        xml = addPane(xml, spec.xSplit || 0, spec.ySplit || 0);

        if (styles) {
          var columnStyle = {}, fmts = spec.numberFormats || {};
          Object.keys(fmts).forEach(function (col) {
            var idx = styles.formatStyle[fmts[col]];
            if (idx !== undefined) columnStyle[col] = idx;
          });
          xml = styleCells(xml, spec.headerRow || 0, styles.headerStyle, columnStyle);
        }

        if (xml === before) return;
        e.data = utf8Encode(xml);
        e.crc = crc32(e.data);
        touched = true;
      });
    });

    if (!touched) return bytes;
    try { return writeZip(zip); } catch (e) { return bytes; }
  }

  return { applyFormatting: applyFormatting, crc32: crc32 };
}));
