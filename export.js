// Builds the Excel workbook from the adminExport response.
// Uses xlsx-js-style (a SheetJS build that writes cell colours). Works in the browser and in Node (tests).

(function (root) {
  const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const CAT_ORDER = { R1: 1, R2: 2, R3: 3, R4: 4, ROTATOR: 5 };
  const ABBR = { DAY12: 'D12', NIGHT12: 'N12', MORNING8: 'M8', EVENING8: 'E8', NIGHT8: 'N8' };
  const REASON = {
    OUTSIDE_GEOFENCE: 'Outside the hospital area', LOCATION_INACCURATE: 'GPS accuracy too low',
    LOCATION_MISSING: 'No location', LOCATION_UNRELIABLE: 'No location accuracy', LOCATION_STALE: 'Old location reading',
    QR_INVALID: 'No / expired screen QR', SHIFT_NOT_CURRENT: 'Shift not open at that time', SHIFT_INVALID: 'No shift chosen',
    ALREADY_RECORDED: 'Already checked in', OVERLAPPING_SHIFT: 'Overlapping shift', DEVICE_USED_RECENTLY: 'Phone used by someone else',
    NAME_INVALID: 'Name not triple', CATEGORY_INVALID: 'No category', SPECIALTY_INVALID: 'No specialty',
    DEVICE_INVALID: 'Device not identified', RATE_LIMITED: 'Too many attempts'
  };

  const C = { green: 'C6EFCE', greenInk: '006100', red: 'FFC7CE', redInk: '9C0006', amber: 'FFEB9C', amberInk: '9C5700',
              grey: 'E7E6E6', head: '0B5F73', white: 'FFFFFF' };
  const ST = {
    title: { font: { bold: true, sz: 14 } },
    sub: { font: { italic: true, color: { rgb: '595959' } } },
    head: { font: { bold: true, color: { rgb: C.white } }, fill: { fgColor: { rgb: C.head } }, alignment: { vertical: 'center', wrapText: true } },
    ok: { fill: { fgColor: { rgb: C.green } }, font: { color: { rgb: C.greenInk } } },
    bad: { fill: { fgColor: { rgb: C.red } }, font: { color: { rgb: C.redInk } } },
    warn: { fill: { fgColor: { rgb: C.amber } }, font: { color: { rgb: C.amberInk } } },
    void: { fill: { fgColor: { rgb: C.grey } }, font: { color: { rgb: '7F7F7F' }, strike: true } },
    bold: { font: { bold: true } },
    center: { alignment: { horizontal: 'center' } }
  };

  function weekday(date) { return WEEKDAYS[new Date(date + 'T00:00:00Z').getUTCDay()]; }
  function catLabel(c) { return c === 'ROTATOR' ? 'Rotator' : c; }
  function merge(a, b) {
    const out = JSON.parse(JSON.stringify(a || {}));
    Object.keys(b || {}).forEach(k => { out[k] = Object.assign({}, out[k] || {}, b[k]); });
    return out;
  }

  /** Writes an array of rows and applies styles: styles[r] = style for the whole row, or a function(c) per cell. */
  function sheet(XLSX, rows, styles, cols) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    rows.forEach((row, r) => {
      const st = styles[r];
      if (!st) return;
      for (let c = 0; c < row.length; c++) {
        const s = typeof st === 'function' ? st(c) : st;
        if (!s) continue;
        const ref = XLSX.utils.encode_cell({ r: r, c: c });
        if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        ws[ref].s = merge(ws[ref].s, s);
      }
    });
    ws['!cols'] = cols.map(w => ({ wch: w }));
    return ws;
  }

  function buildWorkbook(XLSX, data) {
    const rows = (data.rows || []).slice().sort((x, y) =>
      (x.shiftDate + x.shiftStart + x.fullName).localeCompare(y.shiftDate + y.shiftStart + y.fullName));
    const attempts = data.attempts || [];
    const confirmed = rows.filter(r => r.status === 'CONFIRMED');
    const title = 'ER Attendance — ' + (data.hospitalName || 'Emergency Department');
    const period = 'Period: ' + data.from + ' to ' + data.to + '   (shift date = the day the shift starts)';

    // People (by normalised name) with their confirmed check-ins.
    const people = new Map();
    confirmed.forEach(r => {
      const p = people.get(r.nameKey) || { key: r.nameKey, name: r.fullName, category: r.category, specialty: r.specialty, recs: [], outside: 0 };
      p.name = r.fullName; p.category = r.category; p.specialty = r.specialty || p.specialty;
      p.recs.push(r);
      people.set(r.nameKey, p);
    });
    attempts.filter(a => a.outside).forEach(a => {
      const p = people.get(a.nameKey) || { key: a.nameKey, name: a.fullName, category: a.category, specialty: a.specialty, recs: [], outside: 0 };
      p.outside++;
      people.set(a.nameKey, p);
    });
    const list = Array.from(people.values()).sort((a, b) =>
      (CAT_ORDER[a.category] || 9) - (CAT_ORDER[b.category] || 9) || String(a.name).localeCompare(String(b.name)));

    /* ---- Sheet 1: Summary ---- */
    const s1 = [[title], [period], [],
      ['#', 'Full name', 'Category', 'Specialty (Rotators)', 'Shifts checked in', '12h shifts', '8h shifts', 'Attempts from outside', 'Dates & shifts']];
    const st1 = { 0: ST.title, 1: ST.sub, 3: ST.head };
    list.forEach((p, i) => {
      const n12 = p.recs.filter(r => r.shiftType === '12H').length;
      const dates = p.recs.slice().sort((x, y) => x.shiftStart.localeCompare(y.shiftStart))
        .map(r => r.shiftDate + ' ' + (ABBR[r.shift] || r.shift)).join(', ');
      s1.push([i + 1, p.name, catLabel(p.category), p.specialty || '', p.recs.length, n12, p.recs.length - n12, p.outside, dates]);
      st1[s1.length - 1] = c => (c === 4 && p.recs.length ? ST.ok : c === 7 && p.outside ? ST.bad : null);
    });
    s1.push([]);
    s1.push(['', 'Totals', '', '', confirmed.length, confirmed.filter(r => r.shiftType === '12H').length,
             confirmed.filter(r => r.shiftType === '8H').length, attempts.filter(a => a.outside).length, '']);
    st1[s1.length - 1] = ST.bold;
    ['R1', 'R2', 'R3', 'R4', 'ROTATOR'].forEach(c => {
      s1.push(['', catLabel(c) + ' — doctors', list.filter(p => p.category === c && p.recs.length).length, '',
               confirmed.filter(r => r.category === c).length]);
    });
    s1.push([]);
    s1.push(['', 'Shift codes: D12 = Day 12h 08–20 · N12 = Night 12h 20–08 · M8 = Morning 08–16 · E8 = Evening 16–00 · N8 = Night 00–08']);
    st1[s1.length - 1] = ST.sub;

    /* ---- Sheet 2: Check-ins ---- */
    const s2 = [[title], [period], [],
      ['Shift date', 'Day', 'Shift', 'Shift start', 'Check-in', 'Min. from start', 'Full name', 'Category', 'Specialty',
       'Status', 'Distance (m)', 'Accuracy (m)', 'Record ID', 'Notes']];
    const st2 = { 0: ST.title, 1: ST.sub, 3: ST.head };
    rows.forEach(r => {
      s2.push([r.shiftDate, weekday(r.shiftDate), r.shiftLabel, r.shiftStart, r.checkIn,
        r.minutesFromStart === '' ? '' : Number(r.minutesFromStart), r.fullName, catLabel(r.category), r.specialty || '',
        r.status === 'CONFIRMED' ? 'Checked in on site' : 'Voided',
        r.distance === '' ? '' : Number(r.distance), r.accuracy === '' ? '' : Number(r.accuracy), r.attendanceId,
        [r.flags, r.correctionRef].filter(Boolean).join(' ')]);
      st2[s2.length - 1] = r.status === 'CONFIRMED' ? ST.ok : ST.void;
    });

    /* ---- Sheet 3: Daily grid ---- */
    const dates = [];
    for (let d = data.from; d <= data.to; d = addDays(d, 1)) { dates.push(d); if (dates.length > 400) break; }
    const s3 = [[title], [period], [],
      ['Full name', 'Category'].concat(dates.map(d => d.slice(5) + ' ' + weekday(d).slice(0, 3))).concat(['Total'])];
    const st3 = { 0: ST.title, 1: ST.sub, 3: ST.head };
    list.forEach(p => {
      const cells = dates.map(d => {
        const recs = p.recs.filter(r => r.shiftDate === d).map(r => ABBR[r.shift] || r.shift);
        if (recs.length) return { v: recs.join(' + '), k: 'ok' };
        if (attempts.some(a => a.outside && a.nameKey === p.key && a.shiftDate === d)) return { v: '✗ outside', k: 'bad' };
        return { v: '', k: '' };
      });
      s3.push([p.name, catLabel(p.category) + (p.specialty ? ' · ' + p.specialty : '')].concat(cells.map(x => x.v)).concat([p.recs.length]));
      st3[s3.length - 1] = c => (c >= 2 && c < 2 + dates.length && cells[c - 2].k ? merge(ST[cells[c - 2].k], ST.center) : null);
    });
    s3.push(['Checked in', ''].concat(dates.map(d => confirmed.filter(r => r.shiftDate === d).length)).concat([confirmed.length]));
    st3[s3.length - 1] = ST.bold;
    s3.push([]);
    s3.push(['Green = checked in on site (shift code) · Red = tried from outside, not recorded']);
    st3[s3.length - 1] = ST.sub;

    /* ---- Sheet 4: Failed attempts ---- */
    const s4 = [[title], [period], [],
      ['Time', 'Shift date', 'Full name', 'Category', 'Specialty', 'Shift', 'Reason', 'Distance (m)', 'Accuracy (m)']];
    const st4 = { 0: ST.title, 1: ST.sub, 3: ST.head };
    attempts.forEach(a => {
      s4.push([a.checkIn, a.shiftDate, a.fullName, catLabel(a.category), a.specialty || '', a.shiftLabel,
        REASON[a.reason] || a.reason, a.distance === '' ? '' : Number(a.distance), a.accuracy === '' ? '' : Number(a.accuracy)]);
      st4[s4.length - 1] = a.outside ? ST.bad : ST.warn;
    });
    if (!attempts.length) s4.push(['No failed attempts in this period']);

    const wb = XLSX.utils.book_new();
    const w1 = sheet(XLSX, s1, st1, [5, 32, 10, 22, 16, 11, 10, 20, 70]);
    const w2 = sheet(XLSX, s2, st2, [11, 11, 26, 17, 17, 14, 32, 10, 20, 18, 12, 12, 28, 30]);
    const w3 = sheet(XLSX, s3, st3, [32, 22].concat(dates.map(() => 11)).concat([8]));
    const w4 = sheet(XLSX, s4, st4, [17, 11, 32, 10, 20, 26, 26, 12, 12]);
    w2['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: Math.max(3, s2.length - 1), c: 13 } }) };
    w4['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: Math.max(3, s4.length - 1), c: 8 } }) };
    XLSX.utils.book_append_sheet(wb, w1, 'Summary');
    XLSX.utils.book_append_sheet(wb, w2, 'Check-ins');
    XLSX.utils.book_append_sheet(wb, w3, 'Daily grid');
    XLSX.utils.book_append_sheet(wb, w4, 'Failed attempts');
    return wb;
  }

  function addDays(dateKey, n) {
    const d = new Date(dateKey + 'T12:00:00Z');
    d.setUTCDate(d.getUTCDate() + n);
    return d.toISOString().slice(0, 10);
  }

  const api = { buildWorkbook: buildWorkbook, weekday: weekday };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.AttendanceExport = api;
})(this);
