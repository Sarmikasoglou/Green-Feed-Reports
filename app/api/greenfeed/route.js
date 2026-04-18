import { parseUnits } from '@/lib/greenfeed';

const LOGIN_URL = 'https://portal.c-lockinc.com/api/login';
const EMISSIONS_URL = 'https://portal.c-lockinc.com/api/getemissions';
const COLUMNS = [
  'Unit', 'AnimalName', 'RFID', 'StartTime', 'EndTime', 'GoodDataDuration',
  'CO2GramsPerDay', 'CH4GramsPerDay', 'O2GramsPerDay', 'H2GramsPerDay', 'H2SGramsPerDay',
  'AirflowLitersPerSec', 'AirflowCf', 'WindSpeedMetersPerSec', 'WindDirDeg', 'WindCf',
  'WasInterrupted', 'InterruptingTags', 'TempPipeDegreesCelsius', 'IsPreliminary', 'RunTime'
];

function buildUrl(unit, startDate, endDate) {
  const et = `${endDate} 12:00:00`;
  return `${EMISSIONS_URL}?d=visits&fids=${encodeURIComponent(unit)}&st=${encodeURIComponent(startDate)}&et=${encodeURIComponent(et)}`;
}

function parseCsvLike(rawText) {
  const text = String(rawText || '').trim();
  if (!text) return [];
  const low = text.toLowerCase();
  if (low.includes('no data') || low.includes('no records')) return [];
  if (low.includes('unauthorized') || low.includes('forbidden')) {
    throw new Error(`API unauthorized/forbidden. First 300 chars: ${text.slice(0, 300)}`);
  }
  const lines = text.split(/\r?\n/).filter((x) => x.trim());
  if (lines.length < 3) throw new Error(`Unexpected API response format. First 300 chars: ${text.slice(0, 300)}`);
  const dataLines = lines.slice(2);
  const rows = dataLines
    .map((line) => parseCsvLine(line))
    .filter((row) => row.length >= COLUMNS.length)
    .map((row) => Object.fromEntries(COLUMNS.map((col, i) => [col, row[i] ?? ''])));
  return rows;
}

function parseCsvLine(line) {
  const out = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      out.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  out.push(current);
  return out;
}

async function authenticate(username, password) {
  const resp = await fetch(LOGIN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ user: username, pass: password }),
    cache: 'no-store'
  });
  const text = (await resp.text()).trim();
  if (!resp.ok) throw new Error(`Login failed: HTTP ${resp.status}`);
  if (!text) throw new Error('Authentication succeeded but token was empty.');
  return text;
}

async function fetchUnit(token, unit, startDate, endDate) {
  const url = buildUrl(unit, startDate, endDate);
  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ token }),
    cache: 'no-store'
  });
  const text = await resp.text();
  if (!resp.ok) throw new Error(`Unit ${unit}: HTTP ${resp.status}. First 300 chars: ${text.slice(0, 300)}`);
  return parseCsvLike(text).map((row) => ({ ...row, Unit: String(unit) }));
}

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

export async function POST(request) {
  try {
    const { username, password, unitsText, startDate, endDate } = await request.json();
    const units = parseUnits(unitsText);
    if (!username || !password) return Response.json({ error: 'Username and password are required.' }, { status: 400 });
    if (!units.length) return Response.json({ error: 'At least one unit is required.' }, { status: 400 });
    const token = await authenticate(username, password);
    const results = await Promise.all(units.map(async (unit) => ({ unit, rows: await fetchUnit(token, unit, startDate, endDate) })));
    const rows = results.flatMap((x) => x.rows);
    return Response.json({ units, rows, perUnitCounts: results.map((x) => ({ unit: x.unit, rows: x.rows.length })) });
  } catch (error) {
    return Response.json({ error: error.message || 'Unknown error' }, { status: 500 });
  }
}
