// Shared helpers for all pages. No security decision is made in this file —
// the server re-validates everything on every request.

const DEVICE_KEY = 'ed_att_device_id';

async function api(action, payload) {
  try {
    // text/plain keeps this a "simple" CORS request (no preflight), which Apps Script supports.
    const res = await fetch(window.APP_CONFIG.API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ action: action }, payload || {}))
    });
    if (!res.ok) return { ok: false, code: 'NETWORK' };
    return await res.json();
  } catch (e) {
    return { ok: false, code: 'NETWORK' };
  }
}

const store = {
  get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } },
  del(k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }
};

/** Random per-phone identifier (used only to stop one phone checking in several people). */
function deviceId() {
  let id = store.get(DEVICE_KEY);
  if (!id || !/^[a-f0-9]{32}$/.test(id)) {
    const b = new Uint8Array(16);
    crypto.getRandomValues(b);
    id = Array.from(b, x => x.toString(16).padStart(2, '0')).join('');
    store.set(DEVICE_KEY, id);
  }
  return id;
}

const MESSAGES = {
  CONFIRMED:           'Attendance confirmed.',
  OUTSIDE_GEOFENCE:    'Attendance denied: You are outside the authorized hospital area.',
  LOCATION_MISSING:    'Attendance denied: Your location could not be verified.',
  LOCATION_UNRELIABLE: 'Attendance denied: Your location could not be verified.',
  LOCATION_INACCURATE: 'Attendance denied: Location accuracy is too low. Turn on Precise Location and move near an entrance or window.',
  LOCATION_STALE:      'Attendance denied: Location reading is too old. Please try again.',
  LOCATION_DENIED:     'Attendance denied: Location verification failed. Please enable location services and try again.',
  LOCATION_TIMEOUT:    'Attendance denied: Location could not be obtained in time. Please try again.',
  NO_GEOLOCATION:      'This browser does not support location services.',
  NAME_INVALID:        'Enter your full name (at least three words, letters only).',
  CATEGORY_INVALID:    'Choose your category: R1, R2, R3, R4 or Rotator.',
  SPECIALTY_INVALID:   'Rotators: choose your specialty.',
  SHIFT_INVALID:       'Choose your shift type and shift.',
  SHIFT_NOT_CURRENT:   'This shift is not open for check-in now. Check-in opens 2 hours before the shift starts and closes when it ends.',
  OVERLAPPING_SHIFT:   'You are already checked in for a shift that overlaps this one.',
  DEVICE_INVALID:      'This device could not be identified. Do not use private browsing.',
  DEVICE_USED_RECENTLY:'This phone was used to check in someone else within the last 12 hours. Each doctor must use their own phone.',
  NOT_AUTHORIZED:      'Not authorized.',
  ALREADY_RECORDED:    'You are already checked in for this shift.',
  QR_INVALID:          'Attendance denied: Scan the QR code currently shown on the ED screen.',
  RATE_LIMITED:        'Too many attempts. Please wait a few minutes.',
  BUSY:                'System busy. Please try again.',
  CONFIG_ERROR:        'System configuration error. Contact the administrator.',
  SERVER_ERROR:        'Server error. Please try again.',
  NETWORK:             'Network error. Check your internet connection.',
  BAD_REQUEST:         'Invalid request.',
  REASON_REQUIRED:     'A reason of at least 10 characters is required.'
};

function msg(code) {
  return '<div>' + esc(MESSAGES[code] || MESSAGES.SERVER_ERROR) + '</div>';
}

function esc(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function $(id) { return document.getElementById(id); }

/* ---- Shift helpers (display/auto-select only — the server decides). Riyadh is fixed UTC+3. ---- */
const RIYADH_OFFSET_MS = 3 * 3600 * 1000;
function riyadhDateKey(ms) { return new Date(ms + RIYADH_OFFSET_MS).toISOString().slice(0, 10); }
function riyadhStartMs(dateKey, hhmm) { return Date.parse(dateKey + 'T' + hhmm + ':00+03:00'); }
function addDaysKey(dateKey, n) { const d = new Date(dateKey + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function shiftEnd(s) { return String((Number(s.start.slice(0, 2)) + s.hours) % 24).padStart(2, '0') + ':00'; }

/** The occurrence of shift s whose start is closest to nowMs, plus whether it is open (early window .. end). */
function shiftInstance(s, nowMs, earlyMin) {
  const today = riyadhDateKey(nowMs);
  let best = null;
  for (let off = -1; off <= 1; off++) {
    const day = addDaysKey(today, off);
    const start = riyadhStartMs(day, s.start), end = start + s.hours * 3600000;
    const open = nowMs >= start - earlyMin * 60000 && nowMs < end;
    const dist = Math.abs(nowMs - start);
    if (open) return { shiftDate: day, start: start, end: end, open: true, dist: dist };
    if (!best || dist < best.dist) best = { shiftDate: day, start: start, end: end, open: false, dist: dist };
  }
  return best;
}

/** Pick the shift of a type whose start time is nearest to now (open shifts preferred). */
function suggestShift(list, nowMs, earlyMin) {
  let pick = null;
  list.forEach(function (s) {
    const inst = shiftInstance(s, nowMs, earlyMin);
    const score = (inst.open ? 0 : 1e12) + inst.dist;
    if (!pick || score < pick.score) pick = { code: s.code, score: score };
  });
  return pick ? pick.code : '';
}
