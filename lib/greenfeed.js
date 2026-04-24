export const GAS_COLS = ['CH4GramsPerDay', 'CO2GramsPerDay', 'O2GramsPerDay', 'H2GramsPerDay'];

export function digitsOnly(value) {
  return String(value ?? '').replace(/\D+/g, '');
}

export function normalizeId(value) {
  if (value === null || value === undefined) return '';
  const s = String(value).trim();
  if (!s || s.toLowerCase() === 'nan') return '';
  const n = Number(s);
  if (!Number.isNaN(n) && Number.isFinite(n)) return String(Math.trunc(n));
  let out = digitsOnly(s);
  if (s.endsWith('.0') && out.endsWith('0')) out = out.slice(0, -1);
  return out;
}

export function normalizeGreenfeedRfid(value) {
  const s = digitsOnly(value);
  return s.startsWith('000000000') ? s.slice(9) : s;
}

export function durationToSeconds(value) {
  if (value === null || value === undefined) return NaN;
  const s = String(value).trim();
  if (!s || s.toLowerCase() === 'nan') return NaN;
  if (/^\d+(\.\d+)?$/.test(s)) return Number(s);
  const m = s.match(/^(\d{1,3}):(\d{2}):(\d{2})$/);
  if (!m) return NaN;
  const [, hh, mm, ss] = m;
  return Number(hh) * 3600 + Number(mm) * 60 + Number(ss);
}

export function parseUnits(text) {
  return [...new Set(String(text || '').split(/[\s,]+/).map((x) => x.trim()).filter(Boolean))];
}

export function safeNumericAndDuration(rows) {
  return rows.map((row) => {
    const next = { ...row };
    for (const col of GAS_COLS) next[col] = toNumber(row[col]);
    next.GoodDataDuration = durationToSeconds(row.GoodDataDuration);
    next.RunTime = durationToSeconds(row.RunTime);
    next.StartTime = row.StartTime ? new Date(row.StartTime) : null;
    next.EndTime = row.EndTime ? new Date(row.EndTime) : null;
    next.day_start = next.StartTime ? dateOnly(next.StartTime) : null;
    next.day_end = next.EndTime ? dateOnly(next.EndTime) : null;
    next.HourOfDay = next.StartTime
      ? next.StartTime.getHours() + next.StartTime.getMinutes() / 60 + next.StartTime.getSeconds() / 3600
      : NaN;
    next.RFID_norm = normalizeGreenfeedRfid(row.RFID);
    return next;
  });
}

export function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

export function dateOnly(d) {
  return d.toISOString().slice(0, 10);
}

export function chooseGroupCol(rows) {
  if (rows.some((r) => r.EART)) return 'EART';
  if (rows.some((r) => r.FarmName)) return 'FarmName';
  return 'RFID_norm';
}

export function weightedMeanWithFallback(values, weights) {
  let weightedSum = 0;
  let weightSum = 0;
  const valid = [];
  values.forEach((value, i) => {
    const x = toNumber(value);
    const w = toNumber(weights[i]);
    if (Number.isFinite(x)) valid.push(x);
    if (Number.isFinite(x) && Number.isFinite(w) && w > 0) {
      weightedSum += x * w;
      weightSum += w;
    }
  });
  if (weightSum > 0) return weightedSum / weightSum;
  if (!valid.length) return NaN;
  return valid.reduce((a, b) => a + b, 0) / valid.length;
}

export function computeDailyGases(rows, groupCol) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = `${String(row[groupCol] ?? '')}__${String(row.day_end ?? '')}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });

  return Array.from(groups.values()).map((sub) => ({
    [groupCol]: String(sub[0][groupCol] ?? ''),
    day: sub[0].day_end,
    n: sub.length,
    daily_CH4: weightedMeanWithFallback(sub.map((r) => r.CH4GramsPerDay), sub.map((r) => r.GoodDataDuration)),
    daily_CO2: weightedMeanWithFallback(sub.map((r) => r.CO2GramsPerDay), sub.map((r) => r.GoodDataDuration)),
    daily_O2: weightedMeanWithFallback(sub.map((r) => r.O2GramsPerDay), sub.map((r) => r.GoodDataDuration)),
    daily_H2: weightedMeanWithFallback(sub.map((r) => r.H2GramsPerDay), sub.map((r) => r.GoodDataDuration))
  }));
}

export function computeOrder(dailyRows, groupCol) {
  const agg = new Map();
  dailyRows.forEach((row) => {
    const key = String(row[groupCol] ?? '');
    if (!agg.has(key)) agg.set(key, { n: 0, vals: [] });
    const entry = agg.get(key);
    entry.n += Number(row.n || 0);
    if (Number.isFinite(row.daily_CH4)) entry.vals.push(row.daily_CH4);
  });
  return [...agg.entries()]
    .map(([key, value]) => ({ key, mean: value.vals.length ? value.vals.reduce((a, b) => a + b, 0) / value.vals.length : -Infinity }))
    .sort((a, b) => b.mean - a.mean)
    .map((x) => x.key);
}

export function unitBreakdownTable(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const unit = String(row.Unit ?? '');
    if (!map.has(unit)) map.set(unit, { Unit: unit, Readings: 0, CH4_nonmissing: 0, CO2_nonmissing: 0, O2_nonmissing: 0, H2_nonmissing: 0 });
    const entry = map.get(unit);
    entry.Readings += 1;
    if (Number.isFinite(row.CH4GramsPerDay)) entry.CH4_nonmissing += 1;
    if (Number.isFinite(row.CO2GramsPerDay)) entry.CO2_nonmissing += 1;
    if (Number.isFinite(row.O2GramsPerDay)) entry.O2_nonmissing += 1;
    if (Number.isFinite(row.H2GramsPerDay)) entry.H2_nonmissing += 1;
  });
  return [...map.values()].sort((a, b) => b.Readings - a.Readings || a.Unit.localeCompare(b.Unit));
}

export function zscore(arr) {
  const nums = arr.filter(Number.isFinite);
  if (!nums.length) return arr.map(() => NaN);
  const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
  const variance = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / nums.length;
  const sd = Math.sqrt(variance);
  if (!sd) return arr.map(() => NaN);
  return arr.map((x) => (Number.isFinite(x) ? (x - mean) / sd : NaN));
}

export function buildWindowsEveryNHours(stepHours) {
  const labels = [];
  for (let s = 0; s < 24; s += stepHours) labels.push(`${hhLabel(s)}-${hhLabel((s + stepHours) % 24)}`);
  return labels;
}

export function hhLabel(h) {
  const hour = ((h % 24) + 24) % 24;
  const suffix = hour < 12 ? 'AM' : 'PM';
  const h12 = hour % 12 || 12;
  return `${h12}${suffix}`;
}

export function assignWindowByStep(hour, stepHours) {
  if (!Number.isFinite(hour)) return 'Missing';
  const idx = Math.floor(hour / stepHours);
  const start = idx * stepHours;
  return `${hhLabel(start)}-${hhLabel((start + stepHours) % 24)}`;
}

export function csvFromRows(rows) {
  if (!rows.length) return '';
  const headers = Object.keys(rows[0]);
  const esc = (v) => {
    const s = v instanceof Date ? v.toISOString() : String(v ?? '');
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return [headers.join(','), ...rows.map((r) => headers.map((h) => esc(r[h])).join(','))].join('\n');
}
