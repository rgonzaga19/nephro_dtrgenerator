/* DTR core: parse a biometric punch log, compute daily time records, write the payroll workbook. */
(function (root) {
  'use strict';

  var DAY = 86400;
  var MIN_SHIFT_SEC = 30 * 60; // scans closer than this count as one scan (missing time-out)
  var DUTY_DEFAULT = 8.5;
  var ACC = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
  var MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

  function pad(n) { return String(n).padStart(2, '0'); }
  function dayKey(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function keyToUTC(key) { var p = key.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]); }
  function utcToKey(ms) { var d = new Date(ms); return dayKey(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
  function addDays(key, n) { return utcToKey(keyToUTC(key) + n * 86400000); }
  function weekday(key) { return new Date(keyToUTC(key)).getUTCDay(); }
  function fmtUS(key) { var p = key.split('-'); return p[1] + '/' + p[2] + '/' + p[0]; }

  /* ---------- reading the log ---------- */

  function cellText(v) {
    if (v == null) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return String(v);
    if (v instanceof Date) return '';
    if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
    if (v.result != null) return cellText(v.result);
    if (v.text != null) return String(v.text);
    return '';
  }

  function parseStamp(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'object' && !(v instanceof Date)) {
      if (v.result != null) return parseStamp(v.result);
      var t = cellText(v);
      return t ? parseStamp(t) : null;
    }
    if (typeof v === 'number') {
      if (v < 20000) return null;
      v = new Date(Math.round((v - 25569) * DAY) * 1000);
    }
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      return {
        key: dayKey(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate()),
        sec: v.getUTCHours() * 3600 + v.getUTCMinutes() * 60 + v.getUTCSeconds()
      };
    }
    var s = String(v).trim();
    var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?/.exec(s), y, mo, d, h, mi, se, ap;
    if (m) {
      mo = +m[1]; d = +m[2];
      if (mo > 12 && d <= 12) { var tmp = mo; mo = d; d = tmp; }
      y = +m[3]; h = +m[4]; mi = +m[5]; se = +(m[6] || 0); ap = m[7];
    } else {
      m = /^(\d{4})-(\d{2})-(\d{2})[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?/.exec(s);
      if (!m) return null;
      y = +m[1]; mo = +m[2]; d = +m[3]; h = +m[4]; mi = +m[5]; se = +(m[6] || 0); ap = m[7];
    }
    if (ap) h = (h % 12) + (/p/i.test(ap) ? 12 : 0);
    if (mo < 1 || mo > 12 || d < 1 || d > 31 || h > 23 || mi > 59 || se > 59) return null;
    return { key: dayKey(y, mo, d), sec: h * 3600 + mi * 60 + se };
  }

  var DT_HEADERS = ['date/time', 'datetime', 'date time', 'timestamp', 'punch time', 'check time'];

  async function readLog(buffer) {
    var wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    var punches = [], skipped = 0;
    wb.eachSheet(function (ws) {
      var hdr = -1, nameCol = -1, dtCol = -1;
      ws.eachRow({ includeEmpty: false }, function (row, rn) {
        if (hdr > 0 || rn > 20) return;
        var n = -1, d = -1;
        row.eachCell(function (cell, cn) {
          var t = cellText(cell.value).trim().toLowerCase();
          if (t === 'name') n = cn;
          if (DT_HEADERS.indexOf(t) >= 0) d = cn;
        });
        if (n > 0 && d > 0) { hdr = rn; nameCol = n; dtCol = d; }
      });
      if (hdr < 0) { hdr = 1; nameCol = 2; dtCol = 4; } // typical export layout
      ws.eachRow({ includeEmpty: false }, function (row, rn) {
        if (rn <= hdr) return;
        var name = cellText(row.getCell(nameCol).value).replace(/\s+/g, ' ').trim();
        var st = parseStamp(row.getCell(dtCol).value);
        if (!name || !st) { skipped++; return; }
        punches.push({ name: name, key: st.key, sec: st.sec });
      });
    });
    return { punches: punches, skipped: skipped };
  }

  /* ---------- grouping and period ---------- */

  function groupEmployees(punches, o) {
    var map = new Map(), cutoff = (o && o.overnight) ? (o.cutoffHour || 10) * 3600 : 0;
    punches.forEach(function (p) {
      var id = p.name.toUpperCase();
      var e = map.get(id);
      if (!e) { e = { id: id, rawName: p.name, days: new Map() }; map.set(id, e); }
      var key = p.key, sec = p.sec;
      if (cutoff && sec < cutoff) { key = addDays(key, -1); sec += DAY; }
      if (!e.days.has(key)) e.days.set(key, []);
      e.days.get(key).push(sec);
    });
    return Array.from(map.values());
  }

  function displayName(raw) {
    var parts = raw.replace(/\s+/g, ' ').trim().split(' ');
    if (parts.length > 2) parts = parts.filter(function (t) { return !/^[A-Za-z]\.?$/.test(t); });
    return parts.join(' ').toUpperCase();
  }

  // Cut-offs are the 9th-23rd and the 24th-8th, chosen from the latest punch in the file.
  function detectPeriod(punches) {
    var last = punches.reduce(function (a, p) { return p.key > a ? p.key : a; }, punches[0].key);
    var p = last.split('-').map(Number), y = p[0], m = p[1], d = p[2];
    if (d >= 9 && d <= 23) return { start: dayKey(y, m, 9), end: dayKey(y, m, 23) };
    if (d >= 24) return { start: dayKey(y, m, 24), end: m === 12 ? dayKey(y + 1, 1, 8) : dayKey(y, m + 1, 8) };
    return { start: m === 1 ? dayKey(y - 1, 12, 24) : dayKey(y, m - 1, 24), end: dayKey(y, m, 8) };
  }

  function fullRange(punches) {
    var keys = punches.map(function (p) { return p.key; }).sort();
    return { start: keys[0], end: keys[keys.length - 1] };
  }

  /* ---------- daily rows ---------- */

  function buildRows(emp, period, o) {
    var rows = [], outside = [];
    for (var k = period.start; k <= period.end; k = addDays(k, 1)) {
      var wd = weekday(k), label = wd === 6 ? 'Saturday' : wd === 0 ? 'Sunday' : '';
      var row = { key: k, label: label, inSec: null, outSec: null, hrs: null, late: null, status: label ? 'weekend' : 'absent', scans: 0 };
      var secs = emp.days.get(k);
      if (secs && secs.length) {
        secs = secs.slice().sort(function (a, b) { return a - b; });
        var i = secs[0], out = secs[secs.length - 1];
        row.scans = secs.length;
        if (o.truncate) { i = Math.floor(i / 60) * 60; out = Math.floor(out / 60) * 60; }
        row.inSec = i;
        if (secs.length < 2 || secs[secs.length - 1] - secs[0] < MIN_SHIFT_SEC) {
          row.status = 'incomplete';
        } else {
          row.outSec = out;
          row.status = 'ok';
          row.hrs = (out - i) / 3600;
          row.late = row.hrs - o.duty;
        }
      }
      rows.push(row);
    }
    emp.days.forEach(function (_, key) { if (key < period.start || key > period.end) outside.push(key); });
    outside.sort();
    return { rows: rows, outside: outside };
  }

  function summarize(rows, o) {
    var s = { worked: 0, hours: 0, net: 0, incomplete: 0, absent: 0 };
    rows.forEach(function (r) {
      if (r.status === 'ok') { s.worked++; s.hours += r.hrs; s.net += r.late; }
      else if (r.status === 'incomplete') s.incomplete++;
      else if (r.status === 'absent') s.absent++;
    });
    s.amount = s.hours * (o.rate || 0);
    return s;
  }

  /* ---------- workbook ---------- */

  var THIN = { style: 'thin' }, MED = { style: 'medium' };
  function border(l, r, t, b) { var o = {}; if (l) o.left = l; if (r) o.right = r; if (t) o.top = t; if (b) o.bottom = b; return o; }

  function sheetName(name, used) {
    var base = (name.replace(/[\[\]:*?\/\\]/g, ' ').trim() || 'DTR').slice(0, 31), n = base, i = 2;
    while (used.has(n.toLowerCase())) { var suf = ' (' + i++ + ')'; n = base.slice(0, 31 - suf.length) + suf; }
    used.add(n.toLowerCase());
    return n;
  }

  function addSheet(wb, used, emp, period, o, genStamp) {
    var built = buildRows(emp, period, o), rows = built.rows;
    var ws = wb.addWorksheet(sheetName(emp.name, used), {
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1,
        margins: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } }
    });
    ws.properties.defaultRowHeight = 14.4;
    [['A', 9.11], ['B', 34.11], ['C', 9.55], ['D', 14.66], ['E', 18.89], ['F', 10.11], ['G', 9.11], ['H', 9.11], ['I', 15.33], ['J', 14.5], ['K', 16.66]]
      .forEach(function (c) { ws.getColumn(c[0]).width = c[1]; });

    ws.mergeCells('B2:K2');
    var t = ws.getCell('B2');
    t.value = '  Payroll for pay period ' + fmtUS(period.start) + ' - ' + fmtUS(period.end);
    t.font = { name: 'Calibri', size: 12, bold: true };
    t.alignment = { horizontal: 'center' };
    ws.getRow(2).height = 15.6; ws.getRow(4).height = 15; ws.getRow(5).height = 16.2; ws.getRow(6).height = 15.6;

    var heads = ['Name', 'Date', 'Time-In', 'Time-Out', 'No. of hrs', 'Total hrs', 'Rate', 'Total Amount', 'Duty Schedule', 'Late/Undertime'];
    heads.forEach(function (h, i) {
      var c = ws.getCell(5, 2 + i);
      c.value = h;
      c.font = { name: 'Calibri', size: 12, bold: true };
      c.alignment = { horizontal: 'center' };
      c.border = i === 0 ? border(THIN, THIN, THIN, THIN) : i === 1 ? border(null, MED, MED, MED) : border(MED, MED, MED, MED);
    });

    // name row
    var nm = ws.getCell('B6');
    nm.value = emp.name;
    nm.font = { name: 'Arial', size: 10, bold: true };
    nm.alignment = { horizontal: 'center' };
    nm.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    nm.border = border(THIN, THIN, THIN, THIN);
    ws.getCell('C6').border = border(null, THIN, null, THIN);
    for (var c6 = 4; c6 <= 11; c6++) ws.getCell(6, c6).border = border(THIN, c6 === 11 ? MED : THIN, null, THIN);

    var first = 7, n = rows.length, spare = 2, totalRow = first + n + spare;

    function frame(r, weekendLabel) {
      var B = ws.getCell(r, 2);
      B.font = { name: 'Calibri', size: 10, bold: true };
      B.alignment = { horizontal: 'center' };
      B.border = border(THIN, THIN, THIN, THIN);
      if (weekendLabel) { B.value = weekendLabel; B.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9DC3E6' } }; }
      var C = ws.getCell(r, 3);
      C.font = { name: 'Calibri', size: 10, bold: true };
      C.alignment = { horizontal: 'center' };
      C.border = border(null, THIN, THIN, THIN);
      C.numFmt = 'd-mmm';
      for (var c = 4; c <= 11; c++) {
        var x = ws.getCell(r, c);
        x.border = border(THIN, c === 11 ? MED : THIN, THIN, THIN);
        x.font = { name: 'Calibri', size: 10, bold: c !== 11 };
        if (c <= 9) x.alignment = { horizontal: 'center', vertical: c <= 5 ? 'center' : undefined };
        if (c === 10) x.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFF0000' } };
        x.numFmt = c <= 5 ? 'h:mm AM/PM' : c === 6 ? 'h:mm' : (c === 10 ? 'General' : ACC);
      }
    }

    var lateSum = 0;
    rows.forEach(function (row, idx) {
      var r = first + idx;
      frame(r, row.label);
      ws.getCell(r, 3).value = new Date(keyToUTC(row.key));
      if (row.inSec != null) ws.getCell(r, 4).value = (row.inSec % DAY) / DAY;
      if (row.status === 'ok') {
        var diff = row.outSec - row.inSec;
        ws.getCell(r, 5).value = (row.outSec % DAY) / DAY;
        ws.getCell(r, 6).value = { formula: '+E' + r + '-D' + r + '+(D' + r + '>E' + r + ')', result: diff / DAY };
        ws.getCell(r, 7).value = { formula: '+F' + r + '*24', result: diff / 3600 };
        ws.getCell(r, 8).value = o.rate || 0;
        ws.getCell(r, 9).value = { formula: '+G' + r + '*H' + r, result: (diff / 3600) * (o.rate || 0) };
        ws.getCell(r, 10).value = o.duty;
        ws.getCell(r, 11).value = { formula: '+G' + r + '-J' + r, result: row.late };
        lateSum += row.late;
      }
    });
    for (var s = 0; s < spare; s++) frame(first + n + s, null);

    for (var c = 2; c <= 11; c++) {
      var tc = ws.getCell(totalRow, c);
      tc.border = border(c === 2 ? THIN : THIN, c === 11 ? MED : THIN, THIN, MED);
      tc.font = { name: 'Calibri', size: 10, bold: c === 10, color: c === 10 ? { argb: 'FFFF0000' } : undefined };
    }
    var tk = ws.getCell(totalRow, 11);
    tk.value = { formula: 'SUM(K' + first + ':K' + (totalRow - 1) + ')', result: lateSum };
    tk.numFmt = ACC;
    ws.getRow(totalRow).height = 15;

    var f = ws.getCell(totalRow + 3, 2);
    f.value = emp.name + ' - ' + genStamp;
    f.font = { name: 'Calibri', size: 11 };
    return ws;
  }

  async function buildWorkbook(employees, period, o, genStamp) {
    var wb = new ExcelJS.Workbook(), used = new Set();
    wb.creator = 'DTR Generator';
    employees.forEach(function (e) { addSheet(wb, used, e, period, o, genStamp); });
    return wb.xlsx.writeBuffer();
  }

  function fileName(period) {
    var s = period.start.split('-').map(Number), e = period.end.split('-').map(Number);
    var a = MONTHS[s[1] - 1] + ' ' + pad(s[2]);
    var b = (s[1] === e[1] ? '' : MONTHS[e[1] - 1] + ' ') + pad(e[2]);
    return 'DTR COMPUTATION ' + a + ' - ' + b + ' ' + e[0] + '.xlsx';
  }

  root.DTR = {
    DAY: DAY, DUTY_DEFAULT: DUTY_DEFAULT, MIN_SHIFT_SEC: MIN_SHIFT_SEC,
    parseStamp: parseStamp, readLog: readLog, groupEmployees: groupEmployees, displayName: displayName,
    detectPeriod: detectPeriod, fullRange: fullRange, buildRows: buildRows, summarize: summarize,
    buildWorkbook: buildWorkbook, fileName: fileName, addDays: addDays, weekday: weekday, fmtUS: fmtUS,
    dayKey: dayKey, keyToUTC: keyToUTC, pad: pad
  };
})(typeof window !== 'undefined' ? window : globalThis);

(function () {
  'use strict';
  var D = window.DTR, DAY = D.DAY, pad = D.pad;
  var $ = function (id) { return document.getElementById(id); };
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  var state = { punches: [], employees: [], names: {}, active: 0, source: 'example', fileName: '', skipped: 0, busy: false };

  /* ---------- example data (invented person, shown until a file is uploaded) ---------- */
  function exampleLog() {
    var out = [], seed = 11;
    function rnd() { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; }
    for (var k = '2026-10-08'; k <= '2026-10-23'; k = D.addDays(k, 1)) {
      var wd = D.weekday(k);
      if (wd === 0 || wd === 6) continue;
      out.push({ name: 'Juan Dela Cruz', key: k, sec: 8 * 3600 });
      out.push({ name: 'Juan Dela Cruz', key: k, sec: 17 * 3600 });
    }
    return out;
  }

  /* ---------- formatting ---------- */
  function fmtTime(sec) {
    sec = sec % DAY; var h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60);
    return (((h + 11) % 12) + 1) + ':' + pad(m) + ' ' + (h < 12 ? 'AM' : 'PM');
  }
  function fmtHM(diff) {
    var m = Math.round(diff / 60); return Math.floor(m / 60) + ':' + pad(m % 60);
  }
  function acc(n) {
    if (Math.abs(n) < 1e-9) return '-';
    return n < 0 ? '(' + Math.abs(n).toFixed(2) + ')' : n.toFixed(2);
  }
  function shortDate(key) { var p = key.split('-'); return MON[+p[1] - 1] + ' ' + (+p[2]); }
  function sheetDate(key) { var p = key.split('-'); return (+p[2]) + '-' + MON[+p[1] - 1]; }
  function stamp() { var d = new Date(); return pad(d.getMonth() + 1) + '/' + pad(d.getDate()) + '/' + d.getFullYear(); }
  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }

  /* ---------- inputs ---------- */
  function opts() {
    var duty = parseFloat($('optDuty').value), rate = parseFloat($('optRate').value);
    return {
      duty: duty > 0 ? duty : D.DUTY_DEFAULT, rate: rate > 0 ? rate : 0,
      truncate: $('optTrunc').checked, overnight: $('optNight').checked, cutoffHour: +$('optCut').value
    };
  }
  function period() {
    var s = $('optStart').value, e = $('optEnd').value;
    if (!s || !e) return { error: 'Choose the first and last day of the pay period.' };
    if (s > e) return { error: 'The pay period ends before it starts.' };
    if ((D.keyToUTC(e) - D.keyToUTC(s)) / 86400000 > 92) return { error: 'The pay period is longer than 93 days.' };
    return { start: s, end: e };
  }
  function setPeriod(p) { $('optStart').value = p.start; $('optEnd').value = p.end; }
  function showMsg(text) { var m = $('msg'); m.hidden = !text; m.textContent = text || ''; }
  function setStatus(text) { $('status').textContent = text || ''; }
  function updateUploadAvailability() {
    var enabled = !period().error;
    $('file').disabled = !enabled;
    $('drop').setAttribute('aria-disabled', enabled ? 'false' : 'true');
    $('btnAuto').disabled = !state.punches.length;
    $('btnWhole').disabled = !state.punches.length;
    $('uploadHint').textContent = enabled
      ? 'or drop it here. Needs the Name and Date/Time columns.'
      : 'Enter a valid pay period above to enable upload.';
  }

  function load(punches, meta) {
    meta = meta || {};
    state.punches = punches || [];
    state.source = meta.source || 'file';
    state.fileName = meta.fileName || '';
    state.skipped = meta.skipped || 0;
    state.names = {};
    state.active = 0;
    if (meta.source !== 'file') setPeriod(D.detectPeriod(state.punches));
    regroup();
    render();
  }

  /* ---------- data flow ---------- */
  function regroup() {
    var o = opts();
    state.employees = D.groupEmployees(state.punches, o).map(function (e) {
      e.name = state.names[e.id] != null ? state.names[e.id] : D.displayName(e.rawName);
      return e;
    });
    if (state.active >= state.employees.length) state.active = 0;
  }

  function clearAll() {
    state.punches = []; state.names = {}; state.employees = []; state.active = 0;
    state.source = ''; state.fileName = ''; state.skipped = 0;
    $('file').value = '';
    $('optStart').value = ''; $('optEnd').value = '';
    $('optDuty').value = '8.5'; $('optRate').value = '0';
    $('optTrunc').checked = true; $('optNight').checked = false;
    $('optCut').value = '10'; $('cutWrap').hidden = true;
    $('nameInput').value = '';
    updateUploadAvailability(); render(); showMsg(''); setStatus('');
  }

  async function onFile(file) {
    if (!file) return;
    if (period().error) {
      showMsg('Enter a valid pay period before uploading a punch log.');
      $('file').value = '';
      return;
    }
    showMsg(''); setStatus('Reading ' + file.name + '…');
    try {
      var buf = await file.arrayBuffer();
      var res = await D.readLog(buf);
      if (!res.punches.length) {
        setStatus('');
        showMsg('No punches found. The file needs a Name column and a Date/Time column, such as 10/8/2026 3:09:56 PM.');
        return;
      }
      setStatus('');
      load(res.punches, { source: 'file', fileName: file.name, skipped: res.skipped });
    } catch (err) {
      setStatus('');
      showMsg('That file could not be read. Save the log as .xlsx (not .xls or .csv) and try again.');
    }
  }

  /* ---------- render ---------- */
  function render() {
    updateUploadAvailability();
    $('result').hidden = !state.source;
    $('emptyPreview').hidden = !!state.source;
    renderSource(); renderTabs(); renderNameInput(); renderSheet();
  }

  function renderSource() {
    var s = $('source'); s.textContent = '';
    var emps = state.employees.length;
    if (!state.source) return;
    if (state.source === 'example') {
      s.appendChild(el('span', 'pill', 'Sample'));
      s.appendChild(el('span', null, 'Sample data only. Upload your file to replace it.'));
    } else {
      s.appendChild(el('span', 'pill ok', 'Loaded'));
      s.appendChild(el('span', null, state.fileName + ' · ' + state.punches.length + ' punches · ' + emps + (emps === 1 ? ' employee' : ' employees')));
    }
  }

  function renderTabs() {
    var t = $('tabs'); t.textContent = '';
    t.hidden = state.employees.length < 2;
    state.employees.forEach(function (e, i) {
      var b = el('button', 'tab', e.name || '(no name)');
      b.type = 'button'; b.setAttribute('role', 'tab'); b.setAttribute('aria-selected', i === state.active ? 'true' : 'false');
      b.addEventListener('click', function () { state.active = i; render(); });
      t.appendChild(b);
    });
  }

  function renderNameInput() {
    var e = state.employees[state.active];
    $('nameInput').value = e ? e.name : '';
  }

  function renderSheet() {
    var wrap = $('sheetwrap'), stats = $('stats'), flags = $('flags');
    wrap.textContent = ''; stats.textContent = ''; flags.textContent = '';
    var e = state.employees[state.active], p = period(), o = opts();
    var canDownload = state.source === 'file' && !!e && !p.error;
    $('btnDownload').disabled = !canDownload;
    var n = state.employees.length;
    $('btnDownload').textContent = n > 1 ? 'Download .xlsx (' + n + ' sheets)' : 'Download .xlsx';
    $('dlHint').textContent = state.source === 'example' ? 'Upload your own file to enable the download.'
      : (n > 1 ? 'One sheet per employee in a single workbook.' : '');
    if (!e) return;
    if (p.error) { wrap.appendChild(el('p', 'err', p.error)); return; }

    var built = D.buildRows(e, p, o), rows = built.rows, sum = D.summarize(rows, o);

    // stats
    function stat(label, value) { var d = el('div'); d.appendChild(el('dt', null, label)); d.appendChild(el('dd', null, value)); stats.appendChild(d); }
    stat('Days worked', String(sum.worked));
    stat('Total hours', sum.hours.toFixed(2));
    stat('Vs. duty hours', (sum.net > 0 ? '+' : sum.net < 0 ? '−' : '') + Math.abs(sum.net).toFixed(2));
    if (o.rate > 0) stat('Total amount', sum.amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }));

    // flags
    function flag(text, info) { flags.appendChild(el('li', info ? 'info' : null, text)); }
    rows.forEach(function (r) {
      if (r.status === 'incomplete') flag(shortDate(r.key) + ': only one scan (' + fmtTime(r.inSec) + '). Time-Out is left blank.');
      if (r.status === 'absent') flag(shortDate(r.key) + ': no scans on a weekday. The row is left blank.', true);
      if (r.label && r.inSec != null) flag(shortDate(r.key) + ' (' + r.label + '): scans found, so hours are counted.', true);
    });
    if (built.outside.length) flag('Scans on ' + built.outside.map(shortDate).join(', ') + ' fall outside the pay period and are left out.', true);
    if (state.skipped) flag(state.skipped + ' row' + (state.skipped === 1 ? '' : 's') + ' in the file had no readable name or date/time and were skipped.', true);

    // sheet
    wrap.appendChild(el('p', 'sheettitle', 'Payroll for pay period ' + D.fmtUS(p.start) + ' - ' + D.fmtUS(p.end)));
    var tb = el('table', 'sheet');
    var cg = document.createElement('colgroup');
    [16, 7, 10, 10, 9, 8, 6, 12, 11, 13].forEach(function (w) { var c = document.createElement('col'); c.style.width = w + '%'; cg.appendChild(c); });
    tb.appendChild(cg);
    var th = document.createElement('thead'), hr = document.createElement('tr');
    ['Name', 'Date', 'Time-In', 'Time-Out', 'No. of hrs', 'Total hrs', 'Rate', 'Total Amount', 'Duty Schedule', 'Late/Undertime']
      .forEach(function (h) { hr.appendChild(el('th', null, h)); });
    th.appendChild(hr); tb.appendChild(th);
    var body = document.createElement('tbody');
    function tr(cells, cls) { var r = document.createElement('tr'); if (cls) r.className = cls; cells.forEach(function (c) { r.appendChild(c); }); body.appendChild(r); }
    function td(text, cls) { return el('td', cls || null, text); }

    var nameRow = [td(e.name, 'nm')]; for (var i = 0; i < 9; i++) nameRow.push(td(''));
    tr(nameRow);
    var late = 0;
    rows.forEach(function (r) {
      var c = [td(r.label, r.label ? 'wk' : null), td(sheetDate(r.key), r.status === 'absent' ? 'abs' : null)];
      if (r.inSec != null) c.push(td(fmtTime(r.inSec))); else c.push(td(''));
      if (r.status === 'ok') {
        var diff = r.outSec - r.inSec;
        c.push(td(fmtTime(r.outSec)), td(fmtHM(diff)), td(acc(r.hrs), 'n'), td(acc(o.rate), 'n'), td(acc(r.hrs * o.rate), 'n'), td(String(o.duty), 'red'), td(acc(r.late), 'n'));
        late += r.late;
      } else {
        c.push(td(r.status === 'incomplete' ? 'missing' : '', r.status === 'incomplete' ? 'miss' : null));
        for (var j = 0; j < 6; j++) c.push(td(''));
      }
      tr(c);
    });
    for (var s = 0; s < 2; s++) { var blank = []; for (var k = 0; k < 10; k++) blank.push(td('')); tr(blank); }
    var tot = []; for (var q = 0; q < 9; q++) tot.push(td('')); tot.push(td(acc(late), 'n'));
    tr(tot, 'total');
    tb.appendChild(body); wrap.appendChild(tb);
    wrap.appendChild(el('p', 'sheetfoot', e.name + ' - ' + stamp()));
  }

  /* ---------- download ---------- */
  async function download() {
    var p = period(), o = opts();
    if (p.error || state.source !== 'file') return;
    var btn = $('btnDownload'); btn.disabled = true; setStatus('Building the workbook…');
    try {
      var buf = await D.buildWorkbook(state.employees, p, o, stamp());
      var blob = new Blob([buf], { type: XLSX_TYPE }), name = D.fileName(p), dl = null;
      try { dl = window.claude && window.claude.use ? await window.claude.use('downloads') : null; } catch (x) { dl = null; }
      if (dl) {
        try { await dl.save({ filename: name, data: blob }); setStatus('Saved ' + name); }
        catch (err) { setStatus(err && err.code === 'declined' ? '' : 'The file could not be saved here. Open this page in its own browser tab and try again.'); }
      } else {
        var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
        document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
        setStatus('Downloaded ' + name);
      }
    } catch (err) {
      setStatus('The workbook could not be built. Check the pay period and try again.');
    }
    btn.disabled = false; renderSheet();
  }

  /* ---------- events ---------- */
  $('file').addEventListener('change', function (ev) { onFile(ev.target.files[0]); ev.target.value = ''; });
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (t) { drop.addEventListener(t, function (ev) { ev.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (t) { drop.addEventListener(t, function (ev) { ev.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (ev) {
    ev.preventDefault(); drop.classList.remove('over');
    if (period().error) { showMsg('Enter a valid pay period before uploading a punch log.'); return; }
    onFile(ev.dataTransfer && ev.dataTransfer.files[0]);
  });

  ['optStart', 'optEnd', 'optDuty', 'optRate', 'optTrunc'].forEach(function (id) { $(id).addEventListener('input', function () { updateUploadAvailability(); renderSheet(); }); });
  ['optNight', 'optCut'].forEach(function (id) {
    $(id).addEventListener('change', function () { $('cutWrap').hidden = !$('optNight').checked; regroup(); render(); });
  });
  $('btnAuto').addEventListener('click', function () { setPeriod(D.detectPeriod(state.punches)); renderSheet(); });
  $('btnWhole').addEventListener('click', function () { setPeriod(D.fullRange(state.punches)); renderSheet(); });
  $('nameInput').addEventListener('input', function (ev) {
    var e = state.employees[state.active]; if (!e) return;
    e.name = ev.target.value.toUpperCase(); state.names[e.id] = e.name;
    ev.target.value = e.name; renderTabs(); renderSheet();
  });
  $('btnDownload').addEventListener('click', download);
  $('btnClear').addEventListener('click', clearAll);
  $('btnSample').addEventListener('click', function () { load(exampleLog(), { source: 'example' }); showMsg(''); setStatus('Sample data loaded.'); });

  function setTheme(theme) {
    var dark = theme === 'dark';
    document.documentElement.setAttribute('data-theme', dark ? 'dark' : 'light');
    $('themeToggle').setAttribute('aria-checked', dark ? 'true' : 'false');
    try { localStorage.setItem('dtr-theme', dark ? 'dark' : 'light'); } catch (err) { /* storage may be unavailable */ }
  }
  $('themeToggle').addEventListener('click', function () {
    setTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
  });
  try { setTheme(localStorage.getItem('dtr-theme') === 'dark' ? 'dark' : 'light'); }
  catch (err) { setTheme('light'); }

  clearAll();
})();
