'use client';

import dynamic from 'next/dynamic';
import { useMemo, useRef, useState } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { toPng } from 'html-to-image';
import jsPDF from 'jspdf';
import {
  assignWindowByStep,
  buildWindowsEveryNHours,
  chooseGroupCol,
  computeDailyGases,
  computeOrder,
  csvFromRows,
  normalizeId,
  safeNumericAndDuration,
  unitBreakdownTable,
  zscore
} from '@/lib/greenfeed';

const Plot = dynamic(() => import('react-plotly.js'), {
  ssr: false
});

const STEP_OPTIONS = [1, 2, 3, 4, 6, 8, 12, 24];
const initialState = {
  username: '',
  password: '',
  unitsText: '682',
  startDate: '2026-02-11',
  endDate: new Date().toISOString().slice(0, 10),
  stepHours: 3,
  reportName: '26ES1_Set1Report.pdf',
  csvName: 'GFdata_merged.csv',
  unitCsvName: 'GF_unit_summary.csv'
};

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => resolve(results.data),
      error: reject
    });
  });
}

function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(firstSheet, { defval: '' }));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function metricCard(label, value) {
  return (
    <div style={{ background: 'white', borderRadius: 12, padding: 16, boxShadow: '0 2px 10px rgba(0,0,0,0.08)' }}>
      <div style={{ fontSize: 13, color: '#5c677d' }}>{label}</div>
      <div style={{ fontSize: 24, fontWeight: 700, marginTop: 4 }}>{value}</div>
    </div>
  );
}

function tablePreview(title, rows) {
  if (!rows?.length) return null;
  const cols = Object.keys(rows[0]).slice(0, 8);
  return (
    <div style={{ background: 'white', borderRadius: 12, padding: 16, boxShadow: '0 2px 10px rgba(0,0,0,0.08)' }}>
      <h3 style={{ marginTop: 0 }}>{title}</h3>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
          <thead>
            <tr>{cols.map((col) => <th key={col} style={thStyle}>{col}</th>)}</tr>
          </thead>
          <tbody>
            {rows.slice(0, 10).map((row, i) => (
              <tr key={i}>{cols.map((col) => <td key={col} style={tdStyle}>{String(row[col] ?? '')}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

const thStyle = { textAlign: 'left', padding: '8px 10px', borderBottom: '1px solid #e3e8f0', background: '#f8fafc' };
const tdStyle = { padding: '8px 10px', borderBottom: '1px solid #eef2f7', fontSize: 13 };
const sectionCard = { background: 'white', borderRadius: 14, padding: 18, boxShadow: '0 2px 12px rgba(0,0,0,0.08)' };

function buildChartSpecs({ matchedRows, dailyRows, groupCol, order, stepHours, treatmentAvailable }) {
  const byDay = aggregateCount(matchedRows, 'day_start');
  const dayLabels = byDay.map((x) => x.key);
  const charts = [];

  charts.push({
    title: 'Total Records Per Day',
    data: [{ type: 'bar', x: dayLabels, y: byDay.map((x) => x.count) }],
    layout: { margin: { t: 50, l: 50, r: 20, b: 80 }, height: 420 }
  });

  const gases = [
    ['CH4', 'CH4GramsPerDay'],
    ['CO2', 'CO2GramsPerDay'],
    ['O2', 'O2GramsPerDay'],
    ['H2', 'H2GramsPerDay']
  ];
  const normTraces = [];
  for (const [label, col] of gases) {
    const values = zscore(matchedRows.map((r) => Number(r[col])));
    normTraces.push({
      type: 'box',
      name: label,
      x: matchedRows.map((r) => r.day_start),
      y: values,
      boxpoints: false
    });
  }
  charts.push({
    title: 'Normalized Gas Production Per Day',
    data: normTraces,
    layout: { margin: { t: 50, l: 50, r: 20, b: 80 }, height: 500, boxmode: 'group' }
  });

  const totalRecords = order.map((animal) => ({ animal, count: dailyRows.filter((r) => String(r[groupCol]) === animal).reduce((a, b) => a + Number(b.n || 0), 0) }));
  charts.push({
    title: 'Total Records Per Animal',
    data: [{ type: 'bar', x: totalRecords.map((x) => x.animal), y: totalRecords.map((x) => x.count) }],
    layout: { margin: { t: 50, l: 50, r: 20, b: 120 }, height: 420 }
  });

  const windows = buildWindowsEveryNHours(Number(stepHours));
  const windowMap = new Map(order.map((animal) => [animal, Object.fromEntries(windows.map((w) => [w, 0]))]));
  matchedRows.forEach((row) => {
    const animal = String(row[groupCol] ?? '');
    if (!windowMap.has(animal)) return;
    const label = assignWindowByStep(Number(row.HourOfDay), Number(stepHours));
    if (label !== 'Missing' && windowMap.get(animal)[label] !== undefined) windowMap.get(animal)[label] += 1;
  });
  const stackedTraces = windows.map((window) => ({
    type: 'bar',
    name: window,
    x: order,
    y: order.map((animal) => {
      const vals = windowMap.get(animal) || {};
      const total = Object.values(vals).reduce((a, b) => a + b, 0);
      return total ? vals[window] / total : 0;
    })
  }));
  charts.push({
    title: `Daily Records Distribution (every ${stepHours}h)`,
    data: stackedTraces,
    layout: { margin: { t: 50, l: 50, r: 20, b: 120 }, height: 420, barmode: 'stack', yaxis: { tickformat: ',.0%' } }
  });

  const dailyGasCharts = [
    ['daily_CH4', 'Methane (CH4) Production Per Animal'],
    ['daily_CO2', 'Carbon Dioxide (CO2) Production Per Animal'],
    ['daily_O2', 'Oxygen (O2) Production Per Animal'],
    ['daily_H2', 'Hydrogen (H2) Production Per Animal']
  ];
  dailyGasCharts.forEach(([col, title]) => {
    const hasValues = dailyRows.some((r) => Number.isFinite(Number(r[col])) && Number(r[col]) !== 0);
    if (!hasValues && col === 'daily_H2') return;
    charts.push({
      title,
      data: order.map((animal) => ({
        type: 'box',
        name: animal,
        y: dailyRows.filter((r) => String(r[groupCol]) === animal).map((r) => Number(r[col])).filter(Number.isFinite),
        boxpoints: false
      })),
      layout: { margin: { t: 50, l: 50, r: 20, b: 140 }, height: 420, showlegend: false }
    });
  });

  if (treatmentAvailable) {
    const treatments = [...new Set(dailyRows.map((r) => String(r.Treatment || 'Missing')))];
    charts.push({
      title: 'Total Records Per Treatment',
      data: [{
        type: 'bar',
        x: treatments,
        y: treatments.map((t) => dailyRows.filter((r) => String(r.Treatment || 'Missing') === t).reduce((a, b) => a + Number(b.n || 0), 0))
      }],
      layout: { margin: { t: 50, l: 50, r: 20, b: 100 }, height: 400 }
    });

    [['daily_CH4', 'Daily Methane (CH4) by Treatment'], ['daily_CO2', 'Daily Carbon Dioxide (CO2) by Treatment'], ['daily_O2', 'Daily Oxygen (O2) by Treatment'], ['daily_H2', 'Daily Hydrogen (H2) by Treatment']]
      .forEach(([col, title]) => {
        const hasValues = dailyRows.some((r) => Number.isFinite(Number(r[col])) && Number(r[col]) !== 0);
        if (!hasValues && col === 'daily_H2') return;
        charts.push({
          title,
          data: treatments.map((t) => ({
            type: 'box',
            name: t,
            y: dailyRows.filter((r) => String(r.Treatment || 'Missing') === t).map((r) => Number(r[col])).filter(Number.isFinite),
            boxpoints: false
          })),
          layout: { margin: { t: 50, l: 50, r: 20, b: 100 }, height: 400, showlegend: false }
        });
      });
  }

  return charts;
}

function aggregateCount(rows, key) {
  const map = new Map();
  rows.forEach((row) => {
    const val = String(row[key] ?? '');
    map.set(val, (map.get(val) || 0) + 1);
  });
  return [...map.entries()].map(([k, count]) => ({ key: k, count })).sort((a, b) => a.key.localeCompare(b.key));
}

export default function Home() {
  const [form, setForm] = useState(initialState);
  const [mvhFile, setMvhFile] = useState(null);
  const [treatFile, setTreatFile] = useState(null);
  const [useTreatments, setUseTreatments] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState(null);
  const chartRefs = useRef([]);

  const charts = useMemo(() => {
    if (!result?.matchedRows?.length) return [];
    return buildChartSpecs(result);
  }, [result]);

  async function handleProcess() {
    setLoading(true);
    setError('');
    try {
      if (!mvhFile) throw new Error('Upload MVH Research.csv first.');
      const gfResp = await fetch('/api/greenfeed', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form)
      });
      const gfJson = await gfResp.json();
      if (!gfResp.ok) throw new Error(gfJson.error || 'GreenFeed request failed.');

      const gfRows = safeNumericAndDuration(gfJson.rows || []);
      const mvhRows = await parseCsvFile(mvhFile);
      const mvhCols = Object.fromEntries(Object.keys(mvhRows[0] || {}).map((c) => [c.toLowerCase(), c]));
      if (!mvhCols.eid || !mvhCols.pen || !mvhCols.eart) throw new Error('MVH file must include EID, Pen, and EART columns.');

      const mvh = mvhRows
        .map((row) => ({
          ...row,
          EID_norm: normalizeId(row[mvhCols.eid]),
          EART: String(row[mvhCols.eart] ?? '').trim(),
          EART_norm: normalizeId(row[mvhCols.eart]),
          Pen: String(row[mvhCols.pen] ?? '').trim()
        }))
        .filter((row) => row.EID_norm || row.EART_norm);

      const mvhMap = new Map(mvh.map((row) => [String(row.EID_norm), row]));
      let matchedRows = gfRows
        .map((row) => ({ ...row, ...(mvhMap.get(String(row.RFID_norm)) || null) }))
        .filter((row) => row.EID_norm);

      let treatmentAvailable = false;
      if (useTreatments && treatFile) {
        const trtRows = await parseExcelFile(treatFile);
        const trtCols = Object.fromEntries(Object.keys(trtRows[0] || {}).map((c) => [c.toLowerCase(), c]));
        if (!trtCols.eart || !trtCols.treatment) throw new Error('Treatments file must include EART and Treatment columns.');
        const trtMap = new Map(
          trtRows
            .map((row) => [normalizeId(row[trtCols.eart]), String(row[trtCols.treatment] ?? '').trim()])
            .filter(([key]) => key)
        );
        matchedRows = matchedRows.map((row) => ({ ...row, Treatment: trtMap.get(String(row.EART_norm)) || '' }));
        treatmentAvailable = matchedRows.some((row) => row.Treatment);
      }

      if (!matchedRows.length) throw new Error('0 rows remained after MVH matching. Check GF RFID vs MVH EID.');

      const groupCol = chooseGroupCol(matchedRows);
      let dailyRows = computeDailyGases(matchedRows, groupCol);
      if (treatmentAvailable) {
        const treatmentMap = new Map();
        matchedRows.forEach((row) => {
          if (!treatmentMap.has(String(row[groupCol]))) treatmentMap.set(String(row[groupCol]), row.Treatment || '');
        });
        dailyRows = dailyRows.map((row) => ({ ...row, Treatment: treatmentMap.get(String(row[groupCol])) || '' }));
      }
      const order = computeOrder(dailyRows, groupCol);
      const unitSummary = unitBreakdownTable(matchedRows);
      const matchedAnimalIds = new Set(matchedRows.map((row) => String(row.EID_norm)));
      const matchedMvh = mvh.filter((row) => matchedAnimalIds.has(String(row.EID_norm)));

      setResult({
        gfRows,
        mvh,
        matchedRows,
        matchedMvh,
        dailyRows,
        groupCol,
        order,
        unitSummary,
        treatmentAvailable,
        stepHours: Number(form.stepHours)
      });
    } catch (err) {
      setError(err.message || 'Unexpected error');
      setResult(null);
    } finally {
      setLoading(false);
    }
  }

  function downloadText(filename, text, mime) {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function exportPdf() {
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    for (let i = 0; i < chartRefs.current.length; i += 1) {
      const node = chartRefs.current[i];
      if (!node) continue;
      const img = await toPng(node, { cacheBust: true, pixelRatio: 2 });
      if (i > 0) pdf.addPage();
      const width = 780;
      const height = 430;
      pdf.text(charts[i]?.title || `Chart ${i + 1}`, 40, 35);
      pdf.addImage(img, 'PNG', 30, 50, width, height);
    }
    pdf.save(form.reportName || 'GreenFeed_Report.pdf');
  }

  return (
    <main style={{ maxWidth: 1400, margin: '0 auto', padding: 20 }}>
      <div style={{ marginBottom: 20 }}>
        <h1 style={{ marginBottom: 6 }}>GreenFeed QC Dashboard + MVH + Treatments</h1>
        <div style={{ color: '#5c677d' }}>JavaScript/Next.js version for Vercel deployment.</div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '360px 1fr', gap: 20, alignItems: 'start' }}>
        <div style={{ ...sectionCard, position: 'sticky', top: 20 }}>
          <h2 style={{ marginTop: 0 }}>Inputs</h2>
          {[
            ['C-LOCK username', 'username'],
            ['C-LOCK password', 'password'],
            ['Unit IDs', 'unitsText'],
            ['Start date', 'startDate'],
            ['End date', 'endDate'],
            ['PDF filename', 'reportName'],
            ['Merged CSV filename', 'csvName'],
            ['Unit summary CSV filename', 'unitCsvName']
          ].map(([label, key]) => (
            <div key={key} style={{ marginBottom: 12 }}>
              <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>{label}</label>
              <input
                type={key.includes('date') ? 'date' : key === 'password' ? 'password' : 'text'}
                value={form[key]}
                onChange={(e) => setForm((s) => ({ ...s, [key]: e.target.value }))}
                style={inputStyle}
              />
            </div>
          ))}

          <div style={{ marginBottom: 12 }}>
            <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>Plot 4 window size</label>
            <select value={form.stepHours} onChange={(e) => setForm((s) => ({ ...s, stepHours: Number(e.target.value) }))} style={inputStyle}>
              {STEP_OPTIONS.map((n) => <option key={n} value={n}>{n} hours</option>)}
            </select>
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>MVH Research.csv</label>
            <input type="file" accept=".csv" onChange={(e) => setMvhFile(e.target.files?.[0] || null)} />
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input type="checkbox" checked={useTreatments} onChange={(e) => setUseTreatments(e.target.checked)} />
              Use Treatments.xlsx
            </label>
          </div>

          {useTreatments && (
            <div style={{ marginBottom: 12 }}>
              <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>Treatments.xlsx</label>
              <input type="file" accept=".xlsx,.xls" onChange={(e) => setTreatFile(e.target.files?.[0] || null)} />
            </div>
          )}

          <button onClick={handleProcess} disabled={loading} style={buttonStyle}>
            {loading ? 'Processing…' : 'Process + Preview'}
          </button>

          {error && <div style={{ marginTop: 12, color: '#b42318', fontSize: 14 }}>{error}</div>}
        </div>

        <div style={{ display: 'grid', gap: 20 }}>
          {result && (
            <>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 14 }}>
                {metricCard('GF rows', result.gfRows.length.toLocaleString())}
                {metricCard('Matched rows', result.matchedRows.length.toLocaleString())}
                {metricCard('Matched cows', new Set(result.matchedRows.map((r) => r.EID_norm)).size.toLocaleString())}
                {metricCard('Grouping variable', result.groupCol)}
              </div>

              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                <button style={buttonStyle} onClick={() => downloadText(form.csvName, csvFromRows(result.matchedRows), 'text/csv')}>Download merged CSV</button>
                <button style={buttonStyle} onClick={() => downloadText(form.unitCsvName, csvFromRows(result.unitSummary), 'text/csv')}>Download unit summary CSV</button>
                <button style={buttonStyle} onClick={exportPdf}>Download PDF report</button>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
                {tablePreview('Matched MVH preview', result.matchedMvh)}
                {tablePreview('Unit contribution preview', result.unitSummary)}
              </div>

              {charts.map((chart, i) => (
                <div key={chart.title + i} ref={(node) => { chartRefs.current[i] = node; }} style={sectionCard}>
                  <h3 style={{ marginTop: 0 }}>{chart.title}</h3>
                  <Plot
                    data={chart.data}
                    layout={{ ...chart.layout, autosize: true }}
                    style={{ width: '100%' }}
                    useResizeHandler
                    config={{ responsive: true, displaylogo: false }}
                  />
                </div>
              ))}
            </>
          )}
        </div>
      </div>
    </main>
  );
}

const inputStyle = {
  width: '100%',
  padding: '10px 12px',
  borderRadius: 10,
  border: '1px solid #d0d7e2',
  boxSizing: 'border-box'
};

const buttonStyle = {
  background: '#184a8b',
  color: 'white',
  border: 'none',
  padding: '10px 14px',
  borderRadius: 10,
  cursor: 'pointer'
};
