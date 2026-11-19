/* =============================================================================
   TXT to Excel Converter
   -----------------------------------------------------------------------------
   Everything runs locally in the browser. Files are read with the File API,
   parsed in memory, and written with SheetJS. Nothing is uploaded or persisted
   (only the light/dark preference is kept in localStorage).

   Layout of this file
     1. Constants
     2. Pure core  – decoding, delimiter detection, parsing, type inference,
                     worksheet / CSV building (no DOM, unit-testable in Node)
     3. App        – DOM wiring, progress, toasts, preview, export
   ========================================================================== */
(function (root) {
  'use strict';

  /* ===========================================================================
     1. CONSTANTS
     ======================================================================== */
  const LIMITS = {
    maxRows: 1048576,        // Excel worksheet row limit (including header)
    maxCols: 16384,          // Excel worksheet column limit
    maxCellChars: 32767,     // Excel cell character limit
    sheetNameMax: 31,
    chunk: 5000              // rows processed between UI yields
  };

  const ALLOWED_EXT = ['txt', 'csv', 'tsv', 'tab', 'dat', 'log', 'text', 'prn', 'psv', 'asc'];

  const DELIMITERS = {
    tab:       { char: '\t', label: 'Tab' },
    comma:     { char: ',',  label: 'Comma (,)' },
    pipe:      { char: '|',  label: 'Pipe (|)' },
    semicolon: { char: ';',  label: 'Semicolon (;)' }
  };

  const ENGINE_SOURCES = [
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',  // SheetJS, pinned version
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',       // fallback CDN
    'vendor/xlsx.full.min.js'                                               // optional local copy for fully offline use
  ];

  class UserError extends Error {}
  class CancelError extends Error {}

  const nextTick = () => new Promise((r) => setTimeout(r, 0));

  /* ===========================================================================
     2. PURE CORE
     ======================================================================== */

  /* ---------- Decoding & file sanity ---------------------------------------- */

  function sniffBinary(buf) {
    const b = buf;
    if (b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04) {
      return 'This looks like a ZIP/Office file (e.g. .xlsx or .docx), not a text file.';
    }
    if (b.length >= 4 && b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46) {
      return 'This looks like a PDF, not a text file.';
    }
    if (b.length >= 4 && b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) {
      return 'This looks like a legacy Office file (.xls/.doc), not a text file.';
    }
    const n = Math.min(b.length, 8192);
    for (let i = 0; i < n; i++) {
      if (b[i] === 0) return 'This file contains binary data and cannot be read as text.';
    }
    return null;
  }

  function decodeBuffer(buf) {
    if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
      return { text: new TextDecoder('utf-8').decode(buf.subarray(3)), encoding: 'UTF-8 (with BOM)' };
    }
    if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) {
      return { text: new TextDecoder('utf-16le').decode(buf.subarray(2)), encoding: 'UTF-16 LE' };
    }
    if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff) {
      return { text: new TextDecoder('utf-16be').decode(buf.subarray(2)), encoding: 'UTF-16 BE' };
    }
    const problem = sniffBinary(buf);
    if (problem) throw new UserError(problem);
    try {
      return { text: new TextDecoder('utf-8', { fatal: true }).decode(buf), encoding: 'UTF-8' };
    } catch (e) {
      return { text: new TextDecoder('windows-1252').decode(buf), encoding: 'Windows-1252 (legacy)' };
    }
  }

  /* ---------- Line cleaning ------------------------------------------------- */

  const NON_DECOR = /[^\s=\-_+*#~|:.]/;

  /** Splits text into lines and drops blank lines and decorative rules (====, ----, +---+). */
  function cleanLines(text) {
    const raw = text.split(/\r\n|\n|\r/);
    const lines = [];
    let blank = 0;
    let decorative = 0;
    for (let i = 0; i < raw.length; i++) {
      const line = raw[i];
      if (NON_DECOR.test(line)) { lines.push(line); continue; }
      const t = line.trim();
      if (t === '') blank++;
      else if (t.length >= 3) decorative++;
      else lines.push(line);
    }
    return { lines, blank, decorative, total: raw.length };
  }

  /* ---------- Delimiter detection & splitting ------------------------------- */

  function countDelim(line, d) {
    let n = 0;
    if (line.indexOf('"') === -1) {
      let i = line.indexOf(d);
      while (i !== -1) { n++; i = line.indexOf(d, i + d.length); }
      return n;
    }
    let inQ = false;
    for (let k = 0; k < line.length; k++) {
      const ch = line[k];
      if (ch === '"') inQ = !inQ;
      else if (!inQ && ch === d) n++;
    }
    return n;
  }

  /** Returns the most plausible delimiter character, or null if none fits. */
  function detectDelimiter(lines) {
    const sample = lines.slice(0, 200);
    if (!sample.length) return null;
    const candidates = ['\t', '|', ';', ','];
    let best = null;
    let bestScore = 0;
    for (const d of candidates) {
      const freq = new Map();
      for (const l of sample) {
        const c = countDelim(l, d);
        freq.set(c, (freq.get(c) || 0) + 1);
      }
      let mode = 0;
      let modeN = 0;
      for (const [c, n] of freq) {
        if (c > 0 && (n > modeN || (n === modeN && c > mode))) { mode = c; modeN = n; }
      }
      if (!modeN) continue;
      const consistency = modeN / sample.length;
      if (consistency < 0.6) continue;
      const score = consistency * Math.min(mode, 50);
      if (score > bestScore + 1e-9) { best = d; bestScore = score; }
    }
    return best;
  }

  /** Splits one line on `delim`, honouring "double quoted" fields (RFC-4180 style). */
  function splitLine(line, delim) {
    if (line.indexOf('"') === -1) return line.split(delim);
    const out = [];
    const n = line.length;
    const dl = delim.length;
    let i = 0;
    for (;;) {
      if (line[i] === '"') {
        i++;
        let val = '';
        while (i < n) {
          if (line[i] === '"') {
            if (line[i + 1] === '"') { val += '"'; i += 2; continue; }
            i++;
            break;
          }
          val += line[i++];
        }
        let rest = '';
        while (i < n && !line.startsWith(delim, i)) rest += line[i++];
        out.push(val + rest);
      } else {
        const j = line.indexOf(delim, i);
        if (j === -1) { out.push(line.slice(i)); break; }
        out.push(line.slice(i, j));
        i = j;
      }
      if (i >= n) break;
      i += dl;
      if (i >= n) { out.push(''); break; }
    }
    return out;
  }

  /* ---------- Parsing ------------------------------------------------------- */

  /**
   * Parses cleaned lines into an array of string arrays.
   * Pipe-bordered tables ("| a | b |") have their outer pipes removed.
   */
  async function parseRows(lines, delim, opts, hooks) {
    const trim = !opts || opts.trim !== false;
    const total = lines.length;
    const rows = [];
    let maxCols = 0;
    let emptySkipped = 0;
    let dashBlanked = 0;
    const blankDash = !!(opts && opts.blankDash);
    let leftB = false;
    let rightB = false;

    if (delim === '|') {
      const sample = lines.slice(0, 200).map((l) => l.trim());
      const lc = sample.filter((l) => l.charCodeAt(0) === 124).length;
      const rc = sample.filter((l) => l.charCodeAt(l.length - 1) === 124).length;
      leftB = lc / sample.length >= 0.9;
      rightB = rc / sample.length >= 0.9;
    }

    for (let s = 0; s < total; s += LIMITS.chunk) {
      const e = Math.min(s + LIMITS.chunk, total);
      for (let i = s; i < e; i++) {
        let line = lines[i];
        let cells;
        if (delim === null || delim === '') {
          cells = [trim ? line.trim() : line];
        } else {
          if (delim === '|') {
            line = line.trim();
            if (leftB && line.charCodeAt(0) === 124) line = line.slice(1);
            if (rightB && line.charCodeAt(line.length - 1) === 124) line = line.slice(0, -1);
          }
          cells = splitLine(line, delim);
          if (trim) for (let k = 0; k < cells.length; k++) cells[k] = cells[k].trim();
        }
        let empty = true;
        for (let k = 0; k < cells.length; k++) {
          if (cells[k] !== '' && cells[k].trim() !== '') { empty = false; break; }
        }
        if (empty) { emptySkipped++; continue; }
        if (blankDash) {
          // A cell that is only "-" is a placeholder, not data. Hyphens inside values (5230-0310003744) are untouched.
          for (let k = 0; k < cells.length; k++) if (cells[k] === '-') { cells[k] = ''; dashBlanked++; }
        }
        if (cells.length > maxCols) maxCols = cells.length;
        rows.push(cells);
      }
      if (hooks) await hooks.tick(e / total);
    }
    return { rows, maxCols, emptySkipped, dashBlanked };
  }

  /* ---------- Numbers & dates ---------------------------------------------- */

  // Accepts plain numbers, 1,234,567.89 and South-Asian lakh/crore grouping (1,00,000.00)
  const NUM_RE = /^[-+]?(\d{1,3}(?:,\d{3})+|\d{1,2}(?:,\d{2})+,\d{3}|\d+)(?:\.(\d+))?$/;

  function matchNumber(s) {
    let t = s;
    let neg = false;
    if (t.charCodeAt(0) === 40) {                       // "(1,234.50)" accounting negatives
      if (t.charCodeAt(t.length - 1) !== 41) return null;
      t = t.slice(1, -1);
      neg = true;
    }
    const m = NUM_RE.exec(t);
    if (!m) return null;
    return { m, neg, t };
  }

  /** "-511.50" or "(511.50)" -> "511.50" (display text for amount columns). */
  function absText(raw) {
    let t = raw;
    if (t.charCodeAt(0) === 40 && t.charCodeAt(t.length - 1) === 41) t = t.slice(1, -1);
    if (t.charCodeAt(0) === 45 || t.charCodeAt(0) === 43) t = t.slice(1);
    return t.trim();
  }

  function parseNumber(s) {
    const r = matchNumber(s);
    if (!r) return NaN;
    let v = Number(r.t.replace(/,/g, ''));
    if (r.neg) v = -v;
    return v;
  }

  const MONTHS = {};
  ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
    .forEach((name, i) => { MONTHS[name] = i + 1; MONTHS[name.slice(0, 3)] = i + 1; });
  MONTHS.sept = 9;

  const TIME = '(?:[ T]+(\\d{1,2}):(\\d{2})(?::(\\d{2}))?(?:\\s*([AaPp])[Mm])?)?';
  const RE_NUMDATE = new RegExp('^(\\d{1,2})([/.-])(\\d{1,2})\\2(\\d{4})' + TIME + '$');
  const RE_MONDATE = new RegExp('^(\\d{1,2})([- /])([A-Za-z]{3,9})\\2(\\d{4})' + TIME + '$');
  const RE_ISODATE = new RegExp('^(\\d{4})([-/.])(\\d{1,2})\\2(\\d{1,2})' + TIME + '$');

  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
  function validYMD(y, m, d) {
    return y >= 1900 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m);
  }

  function parseTimePart(h, mi, s, ap) {
    if (h === undefined) return { ok: true, time: null };
    let H = +h;
    const M = +mi;
    const S = s === undefined ? 0 : +s;
    if (M > 59 || S > 59) return { ok: false };
    if (ap) {
      if (H < 1 || H > 12) return { ok: false };
      H = H % 12;
      if (ap.toLowerCase() === 'p') H += 12;
    } else if (H > 23) {
      return { ok: false };
    }
    return { ok: true, time: { h: H, mi: M, s: S, hasSec: s !== undefined, ampm: !!ap } };
  }

  /** Recognises dd/mm/yyyy, dd-MON-yyyy and yyyy-mm-dd, each with an optional time. */
  function parseDateCell(s) {
    const c = s.charCodeAt(0);
    if (c < 48 || c > 57 || s.length < 8 || s.length > 32) return null;
    let m;
    if ((m = RE_NUMDATE.exec(s))) {
      const t = parseTimePart(m[5], m[6], m[7], m[8]);
      if (!t.ok) return null;
      return { kind: 'num', sep: m[2], a: +m[1], b: +m[3], y: +m[4], time: t.time };
    }
    if ((m = RE_MONDATE.exec(s))) {
      const mon = MONTHS[m[3].toLowerCase()];
      if (!mon) return null;
      const t = parseTimePart(m[5], m[6], m[7], m[8]);
      if (!t.ok) return null;
      return { kind: 'mon', sep: m[2], d: +m[1], m: mon, y: +m[4], time: t.time };
    }
    if ((m = RE_ISODATE.exec(s))) {
      const t = parseTimePart(m[5], m[6], m[7], m[8]);
      if (!t.ok) return null;
      return { kind: 'iso', sep: m[2], y: +m[1], m: +m[3], d: +m[4], time: t.time };
    }
    return null;
  }

  const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
  function toSerial(y, m, d, h, mi, s) {
    const days = Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH) / 86400000);
    return days + (h * 3600 + mi * 60 + s) / 86400;
  }

  function dateToSerial(raw, order) {
    const p = parseDateCell(raw);
    if (!p) return NaN;
    let y = p.y, m, d;
    if (p.kind === 'num') {
      if (order === 'mdy') { m = p.a; d = p.b; } else { d = p.a; m = p.b; }
    } else { m = p.m; d = p.d; }
    if (!validYMD(y, m, d)) return NaN;
    const t = p.time || { h: 0, mi: 0, s: 0 };
    return toSerial(y, m, d, t.h, t.mi, t.s);
  }

  /* ---------- Header detection & table preparation ------------------------- */

  function isNumericLike(s) { return matchNumber(s) !== null; }

  function looksLikeHeader(row, rowCount) {
    if (!row || rowCount < 2) return false;
    const filled = row.filter((c) => c !== '');
    if (filled.length < Math.max(1, row.length * 0.5)) return false;
    for (const c of filled) {
      if (isNumericLike(c) || parseDateCell(c)) return false;
    }
    return true;
  }

  function sameRow(a, b) {
    const n = Math.max(a.length, b.length);
    for (let i = 0; i < n; i++) {
      if ((a[i] === undefined ? '' : a[i]) !== (b[i] === undefined ? '' : b[i])) return false;
    }
    return true;
  }

  function colLetter(i) {
    let s = '';
    let n = i + 1;
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function buildHeaderNames(headerRow, colCount) {
    const seen = new Map();
    const names = [];
    for (let i = 0; i < colCount; i++) {
      let name = ((headerRow && headerRow[i]) || '').replace(/\s+/g, ' ').trim();
      if (!name) name = 'Column ' + (i + 1);
      const key = name.toLowerCase();
      const count = (seen.get(key) || 0) + 1;
      seen.set(key, count);
      if (count > 1) name = name + '_' + count;
      names.push(name);
    }
    return names;
  }

  function prepareTable(parsed, useHeader) {
    const colCount = parsed.maxCols;
    if (colCount > LIMITS.maxCols) {
      throw new UserError('This file has ' + colCount.toLocaleString('en-US') + ' columns, more than Excel supports (' + LIMITS.maxCols.toLocaleString('en-US') + '). Check the delimiter setting.');
    }
    const rows = parsed.rows;
    let headerRow = null;
    let data = rows;
    let removedRepeats = 0;
    if (useHeader && rows.length) {
      headerRow = rows[0];
      data = [];
      for (let i = 1; i < rows.length; i++) {
        if (sameRow(rows[i], headerRow)) removedRepeats++;   // repeated page headers in report-style files
        else data.push(rows[i]);
      }
    }
    let ragged = 0;
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (r.length !== colCount) {
        ragged++;
        while (r.length < colCount) r.push('');
      }
    }
    const hasHeader = !!headerRow;
    const headers = hasHeader
      ? buildHeaderNames(headerRow, colCount)
      : Array.from({ length: colCount }, (_, i) => colLetter(i));
    return { headers, hasHeader, dataRows: data, colCount, ragged, removedRepeats };
  }

  /* ---------- Column type inference ---------------------------------------- */

  const ID_TOKENS = new Set([
    'card', 'ac', 'acc', 'acct', 'account', 'rrn', 'id', 'ids', 'code', 'ref', 'reference',
    'phone', 'mobile', 'msisdn', 'nid', 'tin', 'zip', 'pin', 'cif', 'no', 'num', 'number',
    'iban', 'swift', 'pan', 'sku', 'barcode'
  ]);

  function hasIdToken(header) {
    return header.toLowerCase().split(/[^a-z0-9]+/).some((t) => ID_TOKENS.has(t));
  }

  function newStat() {
    return {
      n: 0, maxLen: 0,
      numOk: true, negCount: 0, intOnly: true, maxDec: 0, thousands: false, leadZero: false, longInt: false, maxIntLen: 0,
      dateOk: true, kind: null, sep: null, dmyOk: true, mdyOk: true, hasTime: false, hasSec: false, ampm: false
    };
  }

  function updateNumStat(st, s) {
    const r = matchNumber(s);
    if (!r) { st.numOk = false; return; }
    if (r.neg || s.charCodeAt(0) === 45) st.negCount++;
    const intPart = r.m[1];
    const dec = r.m[2];
    const digits = intPart.replace(/,/g, '');
    if (dec) {
      st.intOnly = false;
      if (dec.length > st.maxDec) st.maxDec = dec.length;
    }
    if (intPart.indexOf(',') !== -1) st.thousands = true;
    if (digits.length > 1 && digits.charCodeAt(0) === 48) st.leadZero = true;   // 007, 0810...
    if (digits.length > 15 || (!dec && digits.length >= 12)) st.longInt = true; // card numbers, RRNs, phones
    if (!dec && intPart.indexOf(',') === -1 && digits.length > st.maxIntLen) st.maxIntLen = digits.length;
  }

  function updateDateStat(st, s) {
    const p = parseDateCell(s);
    if (!p) { st.dateOk = false; return; }
    if (st.kind === null) { st.kind = p.kind; st.sep = p.sep; }
    else if (st.kind !== p.kind || st.sep !== p.sep) { st.dateOk = false; return; }
    if (p.kind === 'num') {
      if (!validYMD(p.y, p.b, p.a)) st.dmyOk = false;
      if (!validYMD(p.y, p.a, p.b)) st.mdyOk = false;
      if (!st.dmyOk && !st.mdyOk) { st.dateOk = false; return; }
    } else if (!validYMD(p.y, p.m, p.d)) {
      st.dateOk = false;
      return;
    }
    if (p.time) {
      st.hasTime = true;
      if (p.time.hasSec) st.hasSec = true;
      if (p.time.ampm) st.ampm = true;
    }
  }

  function dateFormat(st, order) {
    const sep = st.sep;
    let base;
    if (st.kind === 'iso') base = 'yyyy' + sep + 'mm' + sep + 'dd';
    else if (st.kind === 'mon') base = 'dd' + sep + 'mmm' + sep + 'yyyy';
    else base = order === 'mdy' ? 'mm' + sep + 'dd' + sep + 'yyyy' : 'dd' + sep + 'mm' + sep + 'yyyy';
    if (st.hasTime) base += ' ' + (st.hasSec ? 'hh:mm:ss' : 'hh:mm') + (st.ampm ? ' AM/PM' : '');
    return base;
  }

  function decideColumn(st, header, smart) {
    const info = { type: 'text', fmt: null, order: 'dmy', maxLen: st.maxLen, abs: false, negCount: 0 };
    if (!smart || st.n === 0) return info;
    if (st.dateOk && st.kind) {
      info.type = 'date';
      info.order = st.kind === 'num' && !st.dmyOk ? 'mdy' : 'dmy';
      info.fmt = dateFormat(st, info.order);
      return info;
    }
    if (st.numOk && !st.leadZero && !st.longInt) {
      const idLike = hasIdToken(header) && st.intOnly && !st.thousands && st.maxIntLen >= 6;
      if (!idLike) {
        info.type = 'number';
        // "Amount" columns are exported as positive numbers (reversals carry a minus sign in the TXT)
        if (/amount/i.test(header)) { info.abs = true; info.negCount = st.negCount; }
        info.fmt = st.maxDec > 0 ? '#,##0.' + '0'.repeat(Math.min(st.maxDec, 6)) : (st.thousands ? '#,##0' : '0');
      }
    }
    return info;
  }

  async function inferColumns(table, smart, hooks) {
    const { dataRows, colCount, headers } = table;
    const stats = Array.from({ length: colCount }, newStat);
    const total = dataRows.length;
    for (let s = 0; s < total; s += LIMITS.chunk) {
      const e = Math.min(s + LIMITS.chunk, total);
      for (let i = s; i < e; i++) {
        const row = dataRows[i];
        for (let c = 0; c < colCount; c++) {
          const v = row[c];
          if (v === '') continue;
          const st = stats[c];
          st.n++;
          if (v.length > st.maxLen) st.maxLen = v.length;
          if (smart) {
            if (st.numOk) updateNumStat(st, v);
            if (st.dateOk) updateDateStat(st, v);
          }
        }
      }
      if (hooks) await hooks.tick(e / total);
    }
    return stats.map((st, c) => decideColumn(st, headers[c], smart));
  }

  /* ---------- Worksheet & CSV building ------------------------------------- */

  function sanitizeSheetName(name, fallback) {
    let s = String(name || '').replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '');
    if (s.length > LIMITS.sheetNameMax) s = s.slice(0, LIMITS.sheetNameMax).trim().replace(/'+$/, '');
    if (!s) s = fallback || 'Sheet1';
    if (s.toLowerCase() === 'history') s = 'History_';
    return s;
  }

  function defaultSheetName(fileName) {
    let base = String(fileName || '').replace(/\.[^.]+$/, '');
    if (base.length > LIMITS.sheetNameMax) {
      let prev;
      do {
        prev = base;
        base = base.replace(/^(updated|new|final|latest|copy of)[\s_-]+/i, '');
      } while (base !== prev);
    }
    return sanitizeSheetName(base, 'Sheet1');
  }

  function safeFileName(name) {
    const base = String(name || '').replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').trim();
    return base || 'converted';
  }

  function uniqueSheetNames(base, count) {
    if (count <= 1) return [base];
    const names = [base];
    for (let i = 2; i <= count; i++) {
      const suffix = ' (' + i + ')';
      names.push(base.slice(0, LIMITS.sheetNameMax - suffix.length) + suffix);
    }
    return names;
  }

  function makeCell(raw, col) {
    if (col.type === 'number') {
      const v = parseNumber(raw);
      if (!Number.isNaN(v)) return { t: 'n', v: col.abs ? Math.abs(v) : v, z: col.fmt };
      if (col.abs) return { t: 's', v: raw };
    } else if (col.type === 'date') {
      const v = dateToSerial(raw, col.order);
      if (!Number.isNaN(v)) return { t: 'n', v, z: col.fmt };
    }
    // Text (also the safe fallback): keeps leading zeros, long IDs and Unicode / Bangla intact
    return { t: 's', v: raw.length > LIMITS.maxCellChars ? raw.slice(0, LIMITS.maxCellChars) : raw };
  }

  function columnWidths(table, columns) {
    return columns.map((col, c) => {
      let w = Math.max(col.maxLen || 0, table.hasHeader ? table.headers[c].length : 0);
      if (col.type === 'date') w = Math.max(w, (col.fmt || '').length);
      if (col.type === 'number') w = Math.max(w + 2, 12);
      return { wch: Math.min(Math.max(w + 2, 8), 60) };
    });
  }

  /** Builds plain SheetJS worksheet objects (one per 1,048,576-row block). */
  async function buildSheets(table, columns, hooks) {
    const { headers, hasHeader, dataRows, colCount } = table;
    const letters = Array.from({ length: colCount }, (_, i) => colLetter(i));
    const perSheet = LIMITS.maxRows - (hasHeader ? 1 : 0);
    const sheetCount = Math.max(1, Math.ceil(dataRows.length / perSheet));
    const widths = columnWidths(table, columns);
    const sheets = [];
    let processed = 0;

    for (let si = 0; si < sheetCount; si++) {
      const start = si * perSheet;
      const end = Math.min(start + perSheet, dataRows.length);
      const ws = {};
      let r = 1;
      if (hasHeader) {
        for (let c = 0; c < colCount; c++) ws[letters[c] + '1'] = { t: 's', v: headers[c] };
        r = 2;
      }
      for (let i = start; i < end; i++, r++) {
        const row = dataRows[i];
        for (let c = 0; c < colCount; c++) {
          const raw = row[c];
          if (raw === undefined || raw === '') continue;
          ws[letters[c] + r] = makeCell(raw, columns[c]);
        }
        processed++;
        if (hooks && processed % LIMITS.chunk === 0) await hooks.tick(processed / Math.max(1, dataRows.length));
      }
      const lastRow = Math.max(1, r - 1);
      ws['!ref'] = 'A1:' + letters[colCount - 1] + lastRow;
      ws['!cols'] = widths;
      if (hasHeader) ws['!autofilter'] = { ref: ws['!ref'] };
      sheets.push({ ws, rows: end - start });
    }
    if (hooks) await hooks.tick(1);
    return sheets;
  }

  function csvField(v) {
    if (v === '') return '';
    return /[",\r\n]/.test(v) || v !== v.trim() ? '"' + v.replace(/"/g, '""') + '"' : v;
  }

  /** UTF-8 CSV with BOM so Excel opens Bangla / Unicode text correctly. */
  async function buildCsvBlob(table, columns, hooks) {
    const { headers, hasHeader, dataRows, colCount } = table;
    const parts = ['\uFEFF'];
    if (hasHeader) parts.push(headers.map(csvField).join(',') + '\r\n');
    const total = dataRows.length;
    for (let s = 0; s < total; s += LIMITS.chunk) {
      const e = Math.min(s + LIMITS.chunk, total);
      let buf = '';
      for (let i = s; i < e; i++) {
        const row = dataRows[i];
        const out = new Array(colCount);
        for (let c = 0; c < colCount; c++) {
          let v = row[c] === undefined ? '' : row[c];
          if (v !== '' && columns[c].abs && columns[c].type === 'number') v = absText(v);
          out[c] = csvField(v);
        }
        buf += out.join(',') + '\r\n';
      }
      parts.push(buf);
      if (hooks) await hooks.tick(e / Math.max(1, total));
    }
    return new Blob(parts, { type: 'text/csv;charset=utf-8' });
  }

  const core = {
    LIMITS, ALLOWED_EXT, DELIMITERS, UserError, CancelError,
    decodeBuffer, cleanLines, detectDelimiter, splitLine, parseRows, prepareTable,
    looksLikeHeader, inferColumns, absText, buildSheets, buildCsvBlob, parseNumber, parseDateCell, dateToSerial,
    sanitizeSheetName, defaultSheetName, safeFileName, uniqueSheetNames, colLetter
  };

  if (typeof document === 'undefined') {       // Node / test harness
    root.TxtToExcelCore = core;
    return;
  }

  /* ===========================================================================
     3. APP
     ======================================================================== */
  function init() {
    const $ = (id) => document.getElementById(id);
    const el = {
      themeToggle: $('themeToggle'), engineStatus: $('engineStatus'),
      dropzone: $('dropzone'), fileInput: $('fileInput'), alertBox: $('alertBox'),
      progressWrap: $('progressWrap'), progressBar: $('progressBar'), progressLabel: $('progressLabel'), progressPct: $('progressPct'),
      fileSection: $('fileSection'), fiName: $('fiName'), fiSize: $('fiSize'), fiEncoding: $('fiEncoding'),
      fiLines: $('fiLines'), fiSkipped: $('fiSkipped'), fiNotes: $('fiNotes'),
      statRows: $('statRows'), statCols: $('statCols'), statDelim: $('statDelim'),
      delimiter: $('delimiterSelect'), customWrap: $('customDelimWrap'), customDelim: $('customDelim'),
      optHeader: $('optHeader'), optTrim: $('optTrim'), optSmart: $('optSmart'), optDash: $('optDash'),
      sheetName: $('sheetName'), sheetHint: $('sheetHint'), previewRows: $('previewRows'),
      btnXlsx: $('btnXlsx'), btnCsv: $('btnCsv'), btnReset: $('btnReset'),
      previewMeta: $('previewMeta'), previewEmpty: $('previewEmpty'), previewWrap: $('previewWrap'),
      toastRoot: $('toastRoot')
    };

    const state = {
      jobId: 0,
      busy: false,
      file: null,
      lines: null,
      lineStats: null,
      encoding: '',
      detected: null,
      parsed: null,
      table: null,
      columns: null,
      headerTouched: false,
      sheetNameTouched: false
    };

    /* ---------- Formatting helpers ---------- */
    const fmtInt = (n) => n.toLocaleString('en-US');
    function fmtBytes(n) {
      if (n < 1024) return n + ' B';
      if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
      if (n < 1073741824) return (n / 1048576).toFixed(1) + ' MB';
      return (n / 1073741824).toFixed(2) + ' GB';
    }
    const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
    const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ESC[c]);

    /* ---------- Theme ---------- */
    function applyTheme(dark) {
      document.documentElement.classList.toggle('dark', dark);
      el.themeToggle.setAttribute('aria-pressed', String(dark));
      el.themeToggle.setAttribute('aria-label', dark ? 'Switch to light mode' : 'Switch to dark mode');
    }
    applyTheme(document.documentElement.classList.contains('dark'));
    el.themeToggle.addEventListener('click', () => {
      const dark = !document.documentElement.classList.contains('dark');
      applyTheme(dark);
      try { localStorage.setItem('txt2xlsx-theme', dark ? 'dark' : 'light'); } catch (e) { /* storage unavailable */ }
    });

    /* ---------- Toasts & alerts ---------- */
    const TOAST_ICONS = {
      success: '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M4 10.5l4 4 8-9" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
      error: '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 5v6m0 3.5v.01" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="10" cy="10" r="8" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>',
      info: '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 9v5m0-8.5v.01" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="10" cy="10" r="8" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>'
    };
    function toast(message, kind) {
      kind = kind || 'info';
      const t = document.createElement('div');
      t.className = 'toast toast-' + kind;
      t.setAttribute('role', kind === 'error' ? 'alert' : 'status');
      const icon = document.createElement('span');
      icon.className = 'toast-icon';
      icon.innerHTML = TOAST_ICONS[kind];
      const msg = document.createElement('p');
      msg.textContent = message;
      const close = document.createElement('button');
      close.type = 'button';
      close.className = 'toast-close';
      close.setAttribute('aria-label', 'Dismiss notification');
      close.textContent = '×';
      t.append(icon, msg, close);
      el.toastRoot.appendChild(t);
      const remove = () => { if (t.parentNode) t.remove(); };
      close.addEventListener('click', remove);
      setTimeout(remove, kind === 'error' ? 9000 : 5000);
      while (el.toastRoot.children.length > 4) el.toastRoot.firstChild.remove();
    }
    function showAlert(message) { el.alertBox.textContent = message; el.alertBox.hidden = false; }
    function clearAlert() { el.alertBox.hidden = true; el.alertBox.textContent = ''; }

    /* ---------- Progress & busy state ---------- */
    function setProgress(label, frac) {
      const pct = Math.max(0, Math.min(100, Math.round(frac * 100)));
      el.progressLabel.textContent = label;
      el.progressPct.textContent = pct + '%';
      el.progressBar.style.width = pct + '%';
      el.progressWrap.setAttribute('aria-valuenow', String(pct));
    }
    function updateControls() {
      const noTable = !state.table;
      el.btnXlsx.disabled = state.busy || noTable;
      el.btnCsv.disabled = state.busy || noTable;
      [el.delimiter, el.customDelim, el.optHeader, el.optTrim, el.optSmart, el.optDash, el.sheetName]
        .forEach((c) => { c.disabled = state.busy || !state.lines; });
    }
    function setBusy(b) {
      state.busy = b;
      el.progressWrap.hidden = !b;
      el.dropzone.setAttribute('aria-busy', String(b));
      if (!b) setProgress('', 0);
      updateControls();
    }
    function phaseHooks(id, label, from, to) {
      return {
        tick: async (f) => {
          if (id !== state.jobId) throw new CancelError();
          setProgress(label, from + (to - from) * f);
          await nextTick();
        }
      };
    }
    const assertCurrent = (id) => { if (id !== state.jobId) throw new CancelError(); };

    /* ---------- Spreadsheet engine loader (SheetJS) ---------- */
    let enginePromise = null;
    function loadScript(src) {
      return new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = src;
        s.async = true;
        s.onload = () => resolve();
        s.onerror = () => { s.remove(); reject(new Error('Could not load ' + src)); };
        document.head.appendChild(s);
      });
    }
    function ensureEngine() {
      if (root.XLSX) return Promise.resolve(root.XLSX);
      if (!enginePromise) {
        enginePromise = (async () => {
          for (const src of ENGINE_SOURCES) {
            try { await loadScript(src); if (root.XLSX) return root.XLSX; } catch (e) { /* try next source */ }
          }
          throw new UserError('The Excel engine (SheetJS) could not be loaded. Connect to the internet once, or place xlsx.full.min.js inside a "vendor" folder next to index.html. CSV export still works.');
        })().catch((e) => { enginePromise = null; throw e; });
      }
      return enginePromise;
    }
    function setEngineStatus(kind, text) {
      el.engineStatus.dataset.state = kind;
      el.engineStatus.textContent = text;
    }
    ensureEngine()
      .then(() => setEngineStatus('ready', 'Excel engine ready'))
      .catch(() => setEngineStatus('error', 'Excel engine offline'));

    /* ---------- Reading files ---------- */
    async function readFileBytes(file, hooks) {
      if (file.stream) {
        const reader = file.stream().getReader();
        const chunks = [];
        let loaded = 0;
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          chunks.push(value);
          loaded += value.length;
          await hooks.tick(file.size ? loaded / file.size : 1);
        }
        const buf = new Uint8Array(loaded);
        let o = 0;
        for (const c of chunks) { buf.set(c, o); o += c.length; }
        return buf;
      }
      return new Uint8Array(await file.arrayBuffer());
    }

    /* ---------- Delimiter helpers ---------- */
    function delimiterLabel(d) {
      if (d === null || d === '') return 'None (single column)';
      if (d === ' - ') return 'Hyphen ( - )';
      for (const k in DELIMITERS) if (DELIMITERS[k].char === d) return DELIMITERS[k].label;
      return 'Custom "' + d + '"';
    }
    function currentDelimiter() {
      const v = el.delimiter.value;
      if (v === 'auto') return state.detected;
      if (v === 'none') return null;
      if (v === 'custom') {
        const raw = el.customDelim.value;
        if (raw === '\\t') return '\t';
        if (raw.trim() === '-') return ' - ';   // a hyphen only separates when spaced; "5230-0310" and dates stay intact
        return raw;
      }
      return DELIMITERS[v].char;
    }
    function updateAutoLabel() {
      const opt = el.delimiter.querySelector('option[value="auto"]');
      opt.textContent = state.lines ? 'Auto-detect (' + delimiterLabel(state.detected).replace(/ \(.*\)$/, '') + ')' : 'Auto-detect';
    }

    /* ---------- Pipeline: parse → header → infer → render ---------- */
    async function runPipeline(id, reparse) {
      if (reparse || !state.parsed) {
        const delim = currentDelimiter();
        if (delim === '' ) throw new UserError('Enter a custom delimiter character.');
        setProgress('Parsing rows…', 0.3);
        state.parsed = await parseRows(state.lines, delim, { trim: el.optTrim.checked, blankDash: el.optDash.checked }, phaseHooks(id, 'Parsing rows…', 0.3, 0.65));
        assertCurrent(id);
        if (!state.parsed.rows.length) throw new UserError('No data rows could be parsed from this file.');
        if (!state.headerTouched) el.optHeader.checked = delim !== null && looksLikeHeader(state.parsed.rows[0], state.parsed.rows.length);
      }
      const table = prepareTable(state.parsed, el.optHeader.checked);
      if (!table.dataRows.length) throw new UserError('The file only contains a header row – there is no data to convert.');
      const columns = await inferColumns(table, el.optSmart.checked, phaseHooks(id, 'Detecting column types…', 0.65, 0.98));
      assertCurrent(id);
      state.table = table;
      state.columns = columns;
      renderAll();
    }

    function handleError(e, id) {
      if (e instanceof CancelError || id !== state.jobId) return;
      if (!(e instanceof UserError)) console.error(e);
      const msg = e instanceof UserError ? e.message : 'Something went wrong while processing the file: ' + (e && e.message ? e.message : e);
      showAlert(msg);
      toast(msg, 'error');
      renderAll();
    }

    async function refresh(reparse) {
      if (!state.lines) return;
      const id = ++state.jobId;
      clearAlert();
      setBusy(true);
      try {
        await runPipeline(id, reparse);
      } catch (e) {
        if (!(e instanceof CancelError)) { state.table = null; state.columns = null; }
        handleError(e, id);
      } finally {
        if (id === state.jobId) setBusy(false);
      }
    }

    /* ---------- Loading a file ---------- */
    function clearData() {
      state.file = null; state.lines = null; state.lineStats = null; state.parsed = null;
      state.table = null; state.columns = null; state.detected = null; state.encoding = '';
      state.headerTouched = false; state.sheetNameTouched = false;
    }

    async function loadFile(file) {
      const id = ++state.jobId;
      clearData();
      clearAlert();
      renderAll();
      setBusy(true);
      try {
        const ext = (file.name.split('.').pop() || '').toLowerCase();
        const hasExt = file.name.indexOf('.') > 0;
        if (hasExt && ALLOWED_EXT.indexOf(ext) === -1) {
          throw new UserError('“.' + ext + '” files are not supported. Please choose a text file (.txt, .csv, .tsv, .dat, .log).');
        }
        if (file.size === 0) throw new UserError('This file is empty (0 bytes).');

        setProgress('Reading file…', 0);
        const buf = await readFileBytes(file, phaseHooks(id, 'Reading file…', 0, 0.15));
        assertCurrent(id);

        setProgress('Decoding text…', 0.16);
        await nextTick();
        const dec = decodeBuffer(buf);
        assertCurrent(id);
        const cleaned = cleanLines(dec.text);
        if (!cleaned.lines.length) throw new UserError('This file contains no readable data (only blank or decorative lines).');

        state.file = file;
        state.encoding = dec.encoding;
        state.lines = cleaned.lines;
        state.lineStats = { total: cleaned.total, blank: cleaned.blank, decorative: cleaned.decorative };
        state.detected = detectDelimiter(cleaned.lines);

        el.delimiter.value = 'auto';
        el.customWrap.hidden = true;
        el.sheetName.value = defaultSheetName(file.name);
        updateAutoLabel();

        setProgress('Parsing rows…', 0.25);
        await runPipeline(id, true);
        assertCurrent(id);
        toast('Ready: ' + fmtInt(state.table.dataRows.length) + ' rows × ' + fmtInt(state.table.colCount) + ' columns from ' + file.name, 'success');
      } catch (e) {
        if (!(e instanceof CancelError) && id === state.jobId) { clearData(); }
        handleError(e, id);
      } finally {
        if (id === state.jobId) setBusy(false);
      }
    }

    function handleFiles(fileList) {
      const files = Array.from(fileList || []).filter((f) => f);
      if (!files.length) return;
      if (files.length > 1) toast('One file is converted at a time – using “' + files[0].name + '”.', 'info');
      loadFile(files[0]);
    }

    /* ---------- Rendering ---------- */
    const TYPE_LABEL = { text: 'Text', number: 'Number', date: 'Date' };

    function renderPreview() {
      const t = state.table;
      if (!t) {
        el.previewWrap.hidden = true;
        el.previewWrap.innerHTML = '';
        el.previewEmpty.hidden = false;
        el.previewMeta.textContent = '';
        return;
      }
      const limit = Number(el.previewRows.value) || 100;
      const shown = Math.min(limit, t.dataRows.length);
      const cols = state.columns;
      let html = '<table class="data-table"><thead><tr><th class="rownum" scope="col">#</th>';
      for (let c = 0; c < t.colCount; c++) {
        html += '<th scope="col" class="col-' + cols[c].type + '" title="' + esc(t.headers[c]) + '">' +
          '<span class="th-name">' + esc(t.headers[c]) + '</span>' +
          '<span class="type-chip type-' + cols[c].type + '">' + TYPE_LABEL[cols[c].type] + '</span></th>';
      }
      html += '</tr></thead><tbody>';
      for (let i = 0; i < shown; i++) {
        const row = t.dataRows[i];
        html += '<tr><td class="rownum">' + (i + 1) + '</td>';
        for (let c = 0; c < t.colCount; c++) {
          let v = row[c] === undefined ? '' : row[c];
          if (v !== '' && cols[c].abs) v = absText(v);
          html += '<td class="col-' + cols[c].type + '" title="' + esc(v) + '">' + esc(v) + '</td>';
        }
        html += '</tr>';
      }
      html += '</tbody></table>';
      el.previewWrap.innerHTML = html;
      el.previewWrap.hidden = false;
      el.previewEmpty.hidden = true;
      el.previewMeta.textContent = 'Showing ' + fmtInt(shown) + ' of ' + fmtInt(t.dataRows.length) + ' rows';
    }

    function renderAll() {
      const hasFile = !!state.file;
      el.fileSection.hidden = !hasFile;
      el.dropzone.classList.toggle('compact', hasFile);
      if (hasFile) {
        el.fiName.textContent = state.file.name;
        el.fiSize.textContent = fmtBytes(state.file.size);
        el.fiEncoding.textContent = state.encoding;
        const ls = state.lineStats;
        el.fiLines.textContent = fmtInt(ls.total);
        el.fiSkipped.textContent = fmtInt(ls.blank + ls.decorative) + (ls.decorative ? ' (' + fmtInt(ls.decorative) + ' separator)' : '');
        const notes = [];
        const t = state.table;
        if (t) {
          state.columns.forEach((c, i) => {
            if (c.abs && c.negCount) notes.push(fmtInt(c.negCount) + ' negative value' + (c.negCount > 1 ? 's' : '') + ' in “' + t.headers[i] + '” exported as positive');
          });
          if (state.parsed && state.parsed.dashBlanked) notes.push(fmtInt(state.parsed.dashBlanked) + ' placeholder “-” cell' + (state.parsed.dashBlanked > 1 ? 's' : '') + ' left empty');
          if (t.removedRepeats) notes.push(fmtInt(t.removedRepeats) + ' repeated header row' + (t.removedRepeats > 1 ? 's' : '') + ' removed');
          if (t.ragged) notes.push(fmtInt(t.ragged) + ' row' + (t.ragged > 1 ? 's' : '') + ' with a different column count were padded');
          if (state.parsed && state.parsed.emptySkipped) notes.push(fmtInt(state.parsed.emptySkipped) + ' empty row' + (state.parsed.emptySkipped > 1 ? 's' : '') + ' skipped');
          if (t.dataRows.length > LIMITS.maxRows - 1) notes.push('Rows exceed Excel’s limit per sheet – extra sheets will be added');
        }
        el.fiNotes.innerHTML = '';
        notes.forEach((n) => { const li = document.createElement('li'); li.textContent = n; el.fiNotes.appendChild(li); });
        el.fiNotes.hidden = notes.length === 0;
        el.statRows.textContent = state.table ? fmtInt(state.table.dataRows.length) : '–';
        el.statCols.textContent = state.table ? fmtInt(state.table.colCount) : '–';
        const d = currentDelimiter();
        el.statDelim.textContent = el.delimiter.value === 'auto' ? delimiterLabel(d).replace(/ \(.*\)$/, '') : delimiterLabel(d).replace(/ \(.*\)$/, '');
      }
      renderPreview();
      updateControls();
    }

    /* ---------- Exporting ---------- */
    function downloadBlob(blob, filename) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 15000);
    }
    const getSheetName = () => sanitizeSheetName(el.sheetName.value, defaultSheetName(state.file ? state.file.name : ''));

    async function exportExcel() {
      if (!state.table || state.busy) return;
      const id = ++state.jobId;
      clearAlert();
      setBusy(true);
      try {
        setProgress('Loading Excel engine…', 0.02);
        const XLSX = await ensureEngine();
        assertCurrent(id);
        const sheets = await buildSheets(state.table, state.columns, phaseHooks(id, 'Building worksheet…', 0.05, 0.8));
        assertCurrent(id);
        const wb = XLSX.utils.book_new();
        const names = uniqueSheetNames(getSheetName(), sheets.length);
        sheets.forEach((s, i) => XLSX.utils.book_append_sheet(wb, s.ws, names[i]));
        setProgress('Writing .xlsx file…', 0.85);
        await new Promise((r) => setTimeout(r, 30));
        assertCurrent(id);
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
        const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const filename = safeFileName(state.file.name) + '.xlsx';
        setProgress('Starting download…', 1);
        downloadBlob(blob, filename);
        toast('Downloaded ' + filename + ' (' + fmtBytes(blob.size) + ')', 'success');
      } catch (e) {
        handleError(e, id);
      } finally {
        if (id === state.jobId) setBusy(false);
      }
    }

    async function exportCsv() {
      if (!state.table || state.busy) return;
      const id = ++state.jobId;
      clearAlert();
      setBusy(true);
      try {
        const blob = await buildCsvBlob(state.table, state.columns, phaseHooks(id, 'Building CSV…', 0, 0.95));
        assertCurrent(id);
        const filename = safeFileName(state.file.name) + '.csv';
        downloadBlob(blob, filename);
        toast('Downloaded ' + filename + ' (' + fmtBytes(blob.size) + ')', 'success');
      } catch (e) {
        handleError(e, id);
      } finally {
        if (id === state.jobId) setBusy(false);
      }
    }

    /* ---------- Reset ---------- */
    function resetAll(silent) {
      state.jobId++;                       // cancels anything still running
      clearData();
      el.fileInput.value = '';
      el.delimiter.value = 'auto';
      el.customDelim.value = '';
      el.customWrap.hidden = true;
      el.optHeader.checked = true;
      el.optTrim.checked = true;
      el.optSmart.checked = true;
      el.optDash.checked = true;
      el.sheetName.value = '';
      el.sheetHint.textContent = '';
      updateAutoLabel();
      clearAlert();
      setBusy(false);
      renderAll();
      if (!silent) toast('Cleared. Nothing was stored.', 'info');
    }

    /* ---------- Events ---------- */
    el.fileInput.addEventListener('change', () => { handleFiles(el.fileInput.files); el.fileInput.value = ''; });
    el.dropzone.addEventListener('click', (e) => { if (e.target !== el.fileInput) el.fileInput.click(); });
    el.dropzone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); el.fileInput.click(); }
    });

    // Drag & drop anywhere on the page; the drop zone just highlights.
    let dragDepth = 0;
    const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).indexOf('Files') !== -1;
    window.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; el.dropzone.classList.add('dragging'); });
    window.addEventListener('dragover', (e) => { if (!hasFiles(e)) return; e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
    window.addEventListener('dragleave', (e) => { if (!hasFiles(e)) return; dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) el.dropzone.classList.remove('dragging'); });
    window.addEventListener('drop', (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      dragDepth = 0;
      el.dropzone.classList.remove('dragging');
      handleFiles(e.dataTransfer.files);
    });

    el.delimiter.addEventListener('change', () => {
      const custom = el.delimiter.value === 'custom';
      el.customWrap.hidden = !custom;
      state.headerTouched = false;
      if (custom) { el.customDelim.focus(); if (!el.customDelim.value) return; }
      refresh(true);
    });
    let customTimer = null;
    el.customDelim.addEventListener('input', () => {
      clearTimeout(customTimer);
      if (!el.customDelim.value) return;
      customTimer = setTimeout(() => { state.headerTouched = false; refresh(true); }, 350);
    });
    el.optHeader.addEventListener('change', () => { state.headerTouched = true; refresh(false); });
    el.optTrim.addEventListener('change', () => refresh(true));
    el.optDash.addEventListener('change', () => refresh(true));
    el.optSmart.addEventListener('change', () => refresh(false));
    el.previewRows.addEventListener('change', renderPreview);

    el.sheetName.addEventListener('input', () => {
      state.sheetNameTouched = true;
      const raw = el.sheetName.value;
      const cleaned = raw.replace(/[\[\]:*?\/\\]/g, '');
      if (cleaned !== raw) {
        el.sheetName.value = cleaned;
        el.sheetHint.textContent = 'Excel does not allow  [ ] : * ? / \\  in sheet names – removed.';
      } else {
        el.sheetHint.textContent = '';
      }
    });
    el.sheetName.addEventListener('blur', () => {
      if (state.file && !el.sheetName.value.trim()) el.sheetName.value = defaultSheetName(state.file.name);
    });

    el.btnXlsx.addEventListener('click', exportExcel);
    el.btnCsv.addEventListener('click', exportCsv);
    el.btnReset.addEventListener('click', () => resetAll(false));

    resetAll(true);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();

})(typeof globalThis !== 'undefined' ? globalThis : this);