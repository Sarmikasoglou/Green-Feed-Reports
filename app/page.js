'use client';

import dynamic from 'next/dynamic';
import { useEffect, useMemo, useRef, useState } from 'react';
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

const STORAGE_KEY = 'greenfeed-dashboard-preferences-v1';
const STEP_OPTIONS = [1, 2, 3, 4, 6, 8, 12, 24];
const MSU_GREEN = '#18453B';
const MSU_GREEN_DARK = '#12372f';
const MSU_GREEN_MID = '#2f6b5c';
const MSU_GREEN_LIGHT = '#e9f2ef';
const MSU_CREAM = '#f7f4ed';
const MSU_GOLD = '#c8a96a';
const MSU_TEXT = '#1f2e2a';

const initialState = {
  username: '',
  password: '',
  unitsText: '682',
  startDate: '2026-02-11',
  endDate: new Date().toISOString().slice(0, 10),
  stepHours: 3,
  feedPeriodDurationSec: 10800,
  minFeedPeriods: 1,
  reportName: '26ES1_Set1Report.pdf',
  csvName: 'GFdata_merged.csv',
  unitCsvName: 'GF_unit_summary.csv',
  dailyCsvName: 'GF_daily_summary.csv',
  unmatchedCsvName: 'GF_unmatched_rows.csv',
  cupDropsCsvName: 'GF_cupdrops_daily_utilization.csv'
};

const emptyFilters = {
  animalSearch: '',
  selectedAnimals: [],
  selectedUnit: 'All',
  selectedTreatments: [],
  selectedDay: 'All',
  startDate: '',
  endDate: ''
};

const inputStyle = {
  width: '100%',
  padding: '10px 12px',
  borderRadius: 10,
  border: '1px solid #c8d7d0',
  background: '#ffffff',
  boxSizing: 'border-box'
};

const buttonStyle = {
  border: 'none',
  borderRadius: 10,
  background: MSU_GREEN,
  color: '#ffffff',
  padding: '10px 14px',
  fontWeight: 600,
  cursor: 'pointer'
};

const secondaryButtonStyle = {
  ...buttonStyle,
  background: MSU_GREEN_LIGHT,
  color: MSU_GREEN
};

const tabButtonStyle = {
  border: '1px solid #c8d7d0',
  borderRadius: 999,
  background: '#ffffff',
  color: MSU_GREEN_MID,
  padding: '10px 14px',
  fontWeight: 600,
  cursor: 'pointer'
};

const sectionCard = {
  background: '#ffffff',
  borderRadius: 18,
  padding: 18,
  boxShadow: '0 12px 28px rgba(24, 69, 59, 0.08)',
  border: '1px solid #e2ebe6'
};

const thStyle = {
  textAlign: 'left',
  padding: '9px 10px',
  borderBottom: '1px solid #dbe4ee',
  background: '#f7f9fc',
  whiteSpace: 'nowrap'
};

const tdStyle = {
  padding: '9px 10px',
  borderBottom: '1px solid #edf2f7',
  fontSize: 13,
  verticalAlign: 'top'
};

function categoryAxis(extra = {}) {
  return {
    type: 'category',
    categoryorder: 'array',
    tickangle: -45,
    automargin: true,
    ...extra
  };
}

function dayAxis(dayLabels, extra = {}) {
  return categoryAxis({
    categoryarray: dayLabels,
    tickmode: 'array',
    tickvals: dayLabels,
    ticktext: dayLabels,
    tickangle: -60,
    tickfont: { size: 10 },
    ...extra
  });
}

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
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(firstSheet, { defval: '' }));
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function metricCard(label, value, note) {
  return (
    <div style={{ ...sectionCard, padding: 16 }}>
      <div style={{ fontSize: 13, color: '#5f726b' }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 700, marginTop: 6 }}>{value}</div>
      {note ? <div style={{ marginTop: 6, fontSize: 12, color: '#6f7f79' }}>{note}</div> : null}
    </div>
  );
}

function metricSection(title, description, items) {
  return (
    <div style={sectionCard}>
      <div style={{ marginBottom: 14 }}>
        <h2 style={{ margin: 0, fontSize: 20 }}>{title}</h2>
        {description ? <div style={{ marginTop: 6, color: '#64746f', fontSize: 14 }}>{description}</div> : null}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 14 }}>
        {items.map((item, index) => (
          <div key={`${title}-${index}`}>{item}</div>
        ))}
      </div>
    </div>
  );
}

function downloadCard(title, description, buttonLabel, onClick) {
  return (
    <div style={{ ...sectionCard, padding: 16 }}>
      <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 6 }}>{title}</div>
      <div style={{ color: '#64746f', fontSize: 14, lineHeight: 1.5, marginBottom: 14 }}>{description}</div>
      <button style={buttonStyle} onClick={onClick} type="button">
        {buttonLabel}
      </button>
    </div>
  );
}

function formatNumber(value, digits = 1) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 'NA';
  return num.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: digits
  });
}

function aggregateCount(rows, key) {
  const map = new Map();
  rows.forEach((row) => {
    const value = String(row[key] ?? '');
    map.set(value, (map.get(value) || 0) + 1);
  });
  return [...map.entries()]
    .map(([itemKey, count]) => ({ key: itemKey, count }))
    .sort((a, b) => a.key.localeCompare(b.key));
}

function getOptionList(values) {
  return ['All', ...[...new Set(values.filter(Boolean).map(String))].sort((a, b) => a.localeCompare(b))];
}

function tablePreview(title, rows, description) {
  if (!rows?.length) {
    return (
      <div style={sectionCard}>
        <h3 style={{ marginTop: 0, marginBottom: 6 }}>{title}</h3>
        <div style={{ color: '#728197', fontSize: 14 }}>{description || 'No rows available.'}</div>
      </div>
    );
  }

  const columns = Object.keys(rows[0]).slice(0, 8);
  return (
    <div style={sectionCard}>
      <h3 style={{ marginTop: 0, marginBottom: 6 }}>{title}</h3>
      {description ? <div style={{ color: '#728197', fontSize: 14, marginBottom: 12 }}>{description}</div> : null}
      <div style={{ overflowX: 'auto' }}>
        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
          <thead>
            <tr>
              {columns.map((column) => (
                <th key={column} style={thStyle}>
                  {column}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.slice(0, 10).map((row, index) => (
              <tr key={index}>
                {columns.map((column) => (
                  <td key={column} style={tdStyle}>
                    {String(row[column] ?? '')}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function buildTopAnimalTable(dailyRows, groupCol) {
  const map = new Map();
  dailyRows.forEach((row) => {
    const animalId = String(row[groupCol] ?? '');
    if (!animalId) return;
    if (!map.has(animalId)) {
      map.set(animalId, {
        Animal: animalId,
        Days: 0,
        Records: 0,
        Mean_CH4: [],
        Mean_CO2: [],
        Treatment: row.Treatment || '',
        Unit: row.Unit || ''
      });
    }
    const item = map.get(animalId);
    item.Days += 1;
    item.Records += Number(row.n || 0);
    if (Number.isFinite(Number(row.daily_CH4))) item.Mean_CH4.push(Number(row.daily_CH4));
    if (Number.isFinite(Number(row.daily_CO2))) item.Mean_CO2.push(Number(row.daily_CO2));
    if (!item.Treatment && row.Treatment) item.Treatment = row.Treatment;
    if (!item.Unit && row.Unit) item.Unit = row.Unit;
  });

  return [...map.values()]
    .map((row) => ({
      Animal: row.Animal,
      Treatment: row.Treatment || 'Unassigned',
      Unit: row.Unit || 'Unknown',
      Days: row.Days,
      Records: row.Records,
      Mean_CH4: row.Mean_CH4.length ? formatNumber(row.Mean_CH4.reduce((a, b) => a + b, 0) / row.Mean_CH4.length, 2) : 'NA',
      Mean_CO2: row.Mean_CO2.length ? formatNumber(row.Mean_CO2.reduce((a, b) => a + b, 0) / row.Mean_CO2.length, 2) : 'NA'
    }))
    .sort((a, b) => Number(b.Mean_CH4) - Number(a.Mean_CH4))
    .slice(0, 10);
}

function buildDailySummaryRows(dailyRows, groupCol) {
  return dailyRows
    .map((row) => ({
      Animal: String(row[groupCol] ?? ''),
      Day: row.day || '',
      Unit: row.Unit || '',
      Treatment: row.Treatment || '',
      Records: Number(row.n || 0),
      CH4: formatNumber(row.daily_CH4, 2),
      CO2: formatNumber(row.daily_CO2, 2),
      O2: formatNumber(row.daily_O2, 2),
      H2: formatNumber(row.daily_H2, 2)
    }))
    .sort((a, b) => a.Day.localeCompare(b.Day) || a.Animal.localeCompare(b.Animal));
}

function buildTreatmentSummaryRows(dailyRows) {
  const map = new Map();
  dailyRows.forEach((row) => {
    const treatment = String(row.Treatment || 'Unassigned');
    if (!map.has(treatment)) {
      map.set(treatment, { Treatment: treatment, Animals: new Set(), Days: 0, Records: 0, Ch4: [] });
    }
    const item = map.get(treatment);
    if (row.EART || row.Animal || row.RFID_norm) {
      item.Animals.add(String(row.EART || row.Animal || row.RFID_norm || ''));
    }
    item.Days += 1;
    item.Records += Number(row.n || 0);
    if (Number.isFinite(Number(row.daily_CH4))) item.Ch4.push(Number(row.daily_CH4));
  });

  return [...map.values()]
    .map((item) => ({
      Treatment: item.Treatment,
      Animals: item.Animals.size,
      Days: item.Days,
      Records: item.Records,
      Mean_CH4: item.Ch4.length ? formatNumber(item.Ch4.reduce((a, b) => a + b, 0) / item.Ch4.length, 2) : 'NA'
    }))
    .sort((a, b) => a.Treatment.localeCompare(b.Treatment));
}

function buildTreatmentUsageRows({ matchedRows, cupDropRows, mvhRows }) {
  const treatmentAnimalMap = new Map();
  const usingAnimals = new Set();

  mvhRows.forEach((row) => {
    const animal = String(row.EID_norm || '').trim();
    const treatment = String(row.Treatment || '').trim();
    if (!animal || !treatment) return;
    if (!treatmentAnimalMap.has(treatment)) treatmentAnimalMap.set(treatment, new Set());
    treatmentAnimalMap.get(treatment).add(animal);
  });

  cupDropRows.forEach((row) => {
    const animal = String(row.EID_norm || '').trim();
    const treatment = String(row.Treatment || '').trim();
    if (!animal || !treatment) return;
    usingAnimals.add(`${treatment}__${animal}`);
  });

  if (!treatmentAnimalMap.size) {
    matchedRows.forEach((row) => {
      const animal = String(row.EID_norm || '').trim();
      const treatment = String(row.Treatment || '').trim();
      if (!animal || !treatment) return;
      if (!treatmentAnimalMap.has(treatment)) treatmentAnimalMap.set(treatment, new Set());
      treatmentAnimalMap.get(treatment).add(animal);
    });
  }

  return [...treatmentAnimalMap.entries()]
    .map(([treatment, animals]) => {
      const totalAnimals = animals.size;
      const usingCount = [...animals].filter((animal) => usingAnimals.has(`${treatment}__${animal}`)).length;
      return {
        Treatment: treatment,
        AnimalsAssigned: totalAnimals,
        UsingGreenFeed: usingCount,
        NotUsingGreenFeed: Math.max(0, totalAnimals - usingCount),
        UtilizationPct: totalAnimals ? formatNumber((usingCount / totalAnimals) * 100, 1) : '0'
      };
    })
    .sort((a, b) => a.Treatment.localeCompare(b.Treatment));
}

function buildTreatmentDiagnostics({ cupDropRows, mvhRows }) {
  const cupDropMatchedCows = new Set(
    cupDropRows
      .filter((row) => String(row.EID_norm || '').trim() && String(row.EART || '').trim())
      .map((row) => String(row.EID_norm || '').trim())
  );

  const treatmentAssignedCows = new Set(
    mvhRows
      .filter((row) => String(row.EID_norm || '').trim() && String(row.Treatment || '').trim())
      .map((row) => String(row.EID_norm || '').trim())
  );

  const missingFromCupDrops = [...treatmentAssignedCows].filter((animal) => !cupDropMatchedCows.has(animal));

  return {
    cupDropMatchedCows: cupDropMatchedCows.size,
    treatmentAssignedCows: treatmentAssignedCows.size,
    treatmentAssignedMissingFromCupDrops: missingFromCupDrops.length
  };
}

function getUsageDay(row) {
  return String(row.day_start || row.day_end || '');
}

function animalValueForGroup(row, groupCol) {
  if (!row) return '';
  const direct = String(row[groupCol] ?? '').trim();
  if (direct) return direct;
  if (groupCol === 'EART') return String(row.EART || '').trim();
  return String(row.EID_norm || row.RFID_norm || '').trim();
}

function datePart(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const match = raw.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : '';
}

function normalizeCupDropRows(rows, mvhMap, treatmentMap) {
  return rows
    .map((row) => {
      const eid = normalizeId(row.RFIDTag);
      const mvhRow = mvhMap.get(String(eid)) || {};
      const day = datePart(row.FeedTime);
      return {
        ...row,
        EID_norm: eid,
        FeedDate: day,
        Unit: String(row.FeederID || '').trim(),
        CurrentCup_num: Number(row.CurrentCup || 0),
        MaxCups_num: Number(row.MaxCups || 0),
        CurrentPeriod_num: Number(row.CurrentPeriod || 0),
        MaxPeriods_num: Number(row.MaxPeriods || 0),
        CupDelay_num: Number(row.CupDelay || 0),
        PeriodDelay_num: Number(row.PeriodDelay || 0),
        DropMass_num: Number(row.DropMass || 0),
        Treatment: treatmentMap.get(String(mvhRow.EART_norm || '')) || '',
        EART: mvhRow.EART || '',
        Pen: mvhRow.Pen || ''
      };
    })
    .filter((row) => row.EID_norm && row.FeedDate);
}

function buildDailyUtilizationRows(cupDropRows, cohortAnimalIds) {
  if (!cohortAnimalIds.length) return [];
  const minimumFeedPeriods = Math.max(1, Number(arguments[2]) || 1);
  const cohortSet = new Set(cohortAnimalIds);
  const byDayAnimalPeriods = new Map();
  cupDropRows.forEach((row) => {
      const day = String(row.FeedDate || '');
      const animal = String(row.EID_norm || '');
      const period = Number(row.CurrentPeriod_num || 0);
      if (!day || !animal || !period) return;
      if (!cohortSet.has(animal)) return;
      const key = `${day}__${animal}`;
      if (!byDayAnimalPeriods.has(key)) byDayAnimalPeriods.set(key, new Set());
      byDayAnimalPeriods.get(key).add(period);
  });

  const byDay = new Map();
  byDayAnimalPeriods.forEach((periods, key) => {
    const [day, animal] = key.split('__');
    if (periods.size < minimumFeedPeriods) return;
    if (!byDay.has(day)) byDay.set(day, new Set());
    byDay.get(day).add(animal);
  });

  return [...byDay.entries()]
    .map(([day, animals]) => ({
      Day: day,
      CowsUsingGreenFeed: animals.size,
      CohortCows: cohortAnimalIds.length,
      UtilizationPct: cohortAnimalIds.length ? (animals.size / cohortAnimalIds.length) * 100 : 0,
      MinimumFeedPeriods: minimumFeedPeriods
    }))
    .sort((a, b) => a.Day.localeCompare(b.Day));
}

function buildFeedPeriodDistributionRows(cupDropRows, cohortAnimalIds, maxPeriods = 8) {
  if (!cohortAnimalIds.length) return [];
  const dayPeriodMap = new Map();

  cupDropRows.forEach((row) => {
    const day = String(row.FeedDate || '');
    const animal = String(row.EID_norm || '');
    const period = Number(row.CurrentPeriod_num || 0);
    if (!day || !animal || !period) return;
    const key = `${day}__${animal}`;
    if (!dayPeriodMap.has(key)) dayPeriodMap.set(key, new Set());
    dayPeriodMap.get(key).add(period);
  });

  const days = [...new Set(cupDropRows.map((row) => String(row.FeedDate || '')).filter(Boolean))].sort();
  return days.map((day) => {
    const counts = Object.fromEntries(Array.from({ length: maxPeriods + 1 }, (_, idx) => [String(idx), 0]));
    cohortAnimalIds.forEach((animal) => {
      const key = `${day}__${animal}`;
      const periodsUsed = dayPeriodMap.has(key) ? dayPeriodMap.get(key).size : 0;
      const bucket = String(Math.min(periodsUsed, maxPeriods));
      counts[bucket] += 1;
    });
    return {
      Day: day,
      ...counts
    };
  });
}

function buildCowPeriodRowsForSelectedDay(cupDropRows, cohortAnimalIds, selectedDay) {
  if (!selectedDay || selectedDay === 'All') return [];
  const periodMap = new Map();

  cupDropRows
    .filter((row) => String(row.FeedDate || '') === selectedDay)
    .forEach((row) => {
      const animal = String(row.EID_norm || '');
      const period = Number(row.CurrentPeriod_num || 0);
      if (!animal || !period) return;
      if (!periodMap.has(animal)) periodMap.set(animal, new Set());
      periodMap.get(animal).add(period);
    });

  return cohortAnimalIds
    .map((animal) => ({
      Animal: animal,
      PeriodsUsed: periodMap.has(animal) ? periodMap.get(animal).size : 0
    }))
    .sort((a, b) => b.PeriodsUsed - a.PeriodsUsed || a.Animal.localeCompare(b.Animal));
}

function buildDailyPelletIntakeRows(cupDropRows) {
  const byDay = new Map();

  cupDropRows.forEach((row) => {
    const day = String(row.FeedDate || '');
    const dropMass = Number(row.DropMass_num || 0);
    if (!day || !Number.isFinite(dropMass)) return;
    byDay.set(day, (byDay.get(day) || 0) + dropMass);
  });

  return [...byDay.entries()]
    .map(([day, grams]) => ({
      Day: day,
      PelletIntakeGrams: grams
    }))
    .sort((a, b) => a.Day.localeCompare(b.Day));
}

function buildCupDropChartSpecs({ utilizationRows, distributionRows, cowPeriodRows, selectedDay, minimumFeedPeriods }) {
  const charts = [];

  if (utilizationRows.length) {
    charts.push({
      id: 'cupdrops-daily-utilization',
      title: `Daily GreenFeed Utilization (% of cows with at least ${minimumFeedPeriods} feed period${minimumFeedPeriods === 1 ? '' : 's'})`,
      controls: 'feed-period-threshold',
      data: [{
        type: 'bar',
        x: utilizationRows.map((row) => row.Day),
        y: utilizationRows.map((row) => row.UtilizationPct),
        text: utilizationRows.map((row) => `${row.CowsUsingGreenFeed}/${row.CohortCows}`),
        textposition: 'outside',
        marker: { color: '#0f766e' }
      }],
      layout: {
        margin: { t: 50, l: 60, r: 20, b: 110 },
        height: 420,
        yaxis: { title: '% of cows', range: [0, 100] },
        xaxis: dayAxis(utilizationRows.map((row) => row.Day))
      },
      minWidth: Math.max(1000, utilizationRows.length * 70)
    });
  }

  if (distributionRows.length) {
    const buckets = ['0', '1', '2', '3', '4', '5', '6', '7', '8'];
    charts.push({
      id: 'cupdrops-feed-period-distribution',
      title: 'Feed Period Use Distribution (% of cows by periods used per day)',
      data: buckets.map((bucket, index) => ({
        type: 'bar',
        name: `${bucket} period${bucket === '1' ? '' : 's'}`,
        x: distributionRows.map((row) => row.Day),
        y: distributionRows.map((row) => {
          const total = buckets.reduce((sum, key) => sum + Number(row[key] || 0), 0);
          return total ? (Number(row[bucket] || 0) / total) * 100 : 0;
        }),
        marker: { color: ['#cbd5e1', '#1d4ed8', '#0ea5e9', '#10b981', '#84cc16', '#f59e0b', '#f97316', '#ef4444', '#7c3aed'][index] }
      })),
      layout: {
        margin: { t: 50, l: 60, r: 20, b: 110 },
        height: 440,
        barmode: 'stack',
        yaxis: { title: '% of cows', range: [0, 100] },
        xaxis: dayAxis(distributionRows.map((row) => row.Day))
      },
      minWidth: Math.max(1000, distributionRows.length * 70)
    });
  }

  if (cowPeriodRows.length) {
    charts.push({
      id: 'cupdrops-periods-per-cow',
      title: `Feed Periods Used Per Cow${selectedDay && selectedDay !== 'All' ? ` on ${selectedDay}` : ''}`,
      data: [{
        type: 'bar',
        x: cowPeriodRows.map((row) => row.Animal),
        y: cowPeriodRows.map((row) => row.PeriodsUsed),
        marker: { color: '#2563eb' }
      }],
      layout: {
        margin: { t: 50, l: 60, r: 20, b: 160 },
        height: 420,
        yaxis: { title: 'Feed periods used', dtick: 1, range: [0, 8.5] },
        xaxis: categoryAxis({
          categoryarray: cowPeriodRows.map((row) => row.Animal),
          tickmode: 'array',
          tickvals: cowPeriodRows.map((row) => row.Animal),
          ticktext: cowPeriodRows.map((row) => row.Animal),
          tickfont: { size: 11 }
        })
      },
      minWidth: Math.max(1200, cowPeriodRows.length * 42)
    });
  }

  return charts;
}

function buildDailyCowUsageRows(rows, groupCol) {
  const map = new Map();

  rows.forEach((row) => {
    const day = getUsageDay(row);
    const animal = String(row[groupCol] ?? '').trim();
    if (!day || !animal) return;

    if (!map.has(day)) {
      map.set(day, {
        Day: day,
        CowsUsingGreenFeed: new Set()
      });
    }

    map.get(day).CowsUsingGreenFeed.add(animal);
  });

  return [...map.values()]
    .map((row) => {
      const cows = [...row.CowsUsingGreenFeed].sort((a, b) => a.localeCompare(b));
      return {
        Day: row.Day,
        CowCount: cows.length,
        Cows: cows.join(', ')
      };
    })
    .sort((a, b) => a.Day.localeCompare(b.Day));
}

function countCowsUsingFeedPeriod(rows, groupCol, targetDurationSec) {
  const target = Number(targetDurationSec);
  if (!Number.isFinite(target) || target <= 0) return 0;

  const qualifyingCows = new Set();
  const byCowDay = new Map();

  rows.forEach((row) => {
    const day = getUsageDay(row);
    const animal = String(row[groupCol] ?? '').trim();
    if (!day || !animal || !(row.StartTime instanceof Date)) return;
    const key = `${day}__${animal}`;
    if (!byCowDay.has(key)) byCowDay.set(key, []);
    byCowDay.get(key).push(row);
  });

  byCowDay.forEach((cowRows, key) => {
    const [, animal] = key.split('__');
    const ordered = [...cowRows]
      .filter((row) => row.StartTime instanceof Date)
      .sort((a, b) => a.StartTime - b.StartTime);

    for (let i = 1; i < ordered.length; i += 1) {
      const gapSeconds = (ordered[i].StartTime - ordered[i - 1].StartTime) / 1000;
      if (gapSeconds >= target) {
        qualifyingCows.add(animal);
        break;
      }
    }
  });

  return qualifyingCows.size;
}

function buildChartSpecs({ matchedRows, dailyRows, groupCol, order, stepHours, treatmentAvailable, selectedDay, cupDropRows }) {
  const byDay = aggregateCount(matchedRows, 'day_start');
  const dayLabels = byDay.map((entry) => entry.key);
  const charts = [];
  const pelletIntakeByDay = new Map(buildDailyPelletIntakeRows(cupDropRows || []).map((row) => [row.Day, row.PelletIntakeGrams]));
  const pelletIntakeKgLabels = dayLabels.map((day) => {
    const grams = pelletIntakeByDay.get(day);
    return Number.isFinite(grams) ? `${(grams / 1000).toFixed(2)} kg` : '';
  });

  charts.push({
    id: 'gf-total-records-per-day',
    title: pelletIntakeByDay.size ? 'Total Records Per Day (daily pellet intake shown above bars)' : 'Total Records Per Day',
    data: [
      {
        type: 'bar',
        name: 'Record count',
        marker: { color: '#1f5eff' },
        x: dayLabels,
        y: byDay.map((entry) => entry.count),
        text: pelletIntakeKgLabels,
        textposition: 'outside',
        textfont: { color: '#8b6b2f', size: 11 },
        cliponaxis: false,
        hovertemplate: pelletIntakeByDay.size
          ? 'Day: %{x}<br>Records: %{y}<br>Pellet intake: %{text}<extra></extra>'
          : 'Day: %{x}<br>Records: %{y}<extra></extra>'
      },
      ...(pelletIntakeByDay.size ? [{
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Pellet intake (g)',
        line: { color: '#c8a96a', width: 4, dash: 'solid' },
        marker: { size: 9, color: '#c8a96a', symbol: 'diamond' },
        x: dayLabels,
        y: dayLabels.map((day) => pelletIntakeByDay.get(day) ?? 0),
        yaxis: 'y2',
        hovertemplate: 'Day: %{x}<br>Pellet intake: %{y:.0f} g<extra></extra>'
      }] : [])
    ],
    layout: {
      margin: { t: 80, l: 60, r: 90, b: 110 },
      height: 420,
      xaxis: dayAxis(dayLabels),
      yaxis: { title: 'Record count' },
      ...(pelletIntakeByDay.size ? {
        yaxis2: {
          title: 'Pellet intake (g)',
          overlaying: 'y',
          side: 'right',
          showgrid: false,
          tickfont: { color: '#8b6b2f' },
          titlefont: { color: '#8b6b2f' },
          zeroline: false
        },
        legend: {
          orientation: 'h',
          y: 1.12,
          x: 0
        },
        annotations: [
          {
            xref: 'paper',
            yref: 'paper',
            x: 1,
            y: 1.16,
            xanchor: 'right',
            showarrow: false,
            font: { size: 12, color: '#8b6b2f' },
            text: 'Gold line = daily pellet intake'
          }
        ]
      } : {})
    },
    minWidth: Math.max(1000, dayLabels.length * 70)
  });

  const gases = [
    ['CH4', 'CH4GramsPerDay', '#0f766e'],
    ['CO2', 'CO2GramsPerDay', '#1d4ed8'],
    ['O2', 'O2GramsPerDay', '#a16207'],
    ['H2', 'H2GramsPerDay', '#9333ea']
  ];

  charts.push({
    id: 'gf-normalized-gas-per-day',
    title: 'Normalized Gas Production Per Day',
    data: gases.map(([label, column, color]) => ({
      type: 'box',
      name: label,
      marker: { color },
      x: matchedRows.map((row) => String(row.day_start)),
      y: zscore(matchedRows.map((row) => Number(row[column]))),
      boxpoints: false
    })),
    layout: {
      margin: { t: 50, l: 50, r: 20, b: 140 },
      height: 500,
      boxmode: 'group',
      xaxis: dayAxis(dayLabels)
    },
    minWidth: Math.max(1200, dayLabels.length * 95)
  });

  const animalOrder = order.map(String);
  const totalsByAnimal = animalOrder.map((animal) => ({
    animal,
    count: dailyRows
      .filter((row) => String(row[groupCol]) === animal)
      .reduce((sum, row) => sum + Number(row.n || 0), 0)
  }));

  charts.push({
    id: 'gf-total-records-per-animal',
    title: 'Total Records Per Animal',
    data: [
      {
        type: 'bar',
        marker: { color: '#2563eb' },
        x: totalsByAnimal.map((entry) => entry.animal),
        y: totalsByAnimal.map((entry) => entry.count)
      }
    ],
    layout: {
      margin: { t: 50, l: 50, r: 20, b: 140 },
      height: 420,
      xaxis: categoryAxis({ categoryarray: animalOrder })
    }
  });

  const windows = buildWindowsEveryNHours(Number(stepHours));
  const windowMap = new Map(
    animalOrder.map((animal) => [animal, Object.fromEntries(windows.map((window) => [window, 0]))])
  );

  matchedRows.forEach((row) => {
    const animal = String(row[groupCol] ?? '');
    if (!windowMap.has(animal)) return;
    const label = assignWindowByStep(Number(row.HourOfDay), Number(stepHours));
    if (label !== 'Missing' && windowMap.get(animal)[label] !== undefined) {
      windowMap.get(animal)[label] += 1;
    }
  });

  charts.push({
    id: 'gf-daily-records-distribution',
    title: `Daily Records Distribution${selectedDay && selectedDay !== 'All' ? ` for ${selectedDay}` : ''} (every ${stepHours}h, n=${animalOrder.length} cows)`,
    controls: 'window-size',
    data: windows.map((window, index) => ({
      type: 'bar',
      name: window,
      x: animalOrder,
      y: animalOrder.map((animal) => {
        const values = windowMap.get(animal) || {};
        const total = Object.values(values).reduce((sum, value) => sum + value, 0);
        return total ? values[window] / total : 0;
      }),
      marker: {
        color: ['#1f5eff', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#84cc16', '#f97316'][index % 8]
      }
    })),
    layout: {
      margin: { t: 50, l: 50, r: 20, b: 140 },
      height: 420,
      barmode: 'stack',
      yaxis: { tickformat: ',.0%' },
      xaxis: categoryAxis({
        categoryarray: animalOrder,
        tickmode: 'array',
        tickvals: animalOrder,
        ticktext: animalOrder,
        tickfont: { size: 11 }
      })
    },
    minWidth: Math.max(1200, animalOrder.length * 42)
  });

  [
    ['daily_CH4', 'Methane (CH4) Production Per Animal'],
    ['daily_CO2', 'Carbon Dioxide (CO2) Production Per Animal'],
    ['daily_O2', 'Oxygen (O2) Production Per Animal'],
    ['daily_H2', 'Hydrogen (H2) Production Per Animal']
  ].forEach(([column, title]) => {
    const hasValues = dailyRows.some((row) => Number.isFinite(Number(row[column])) && Number(row[column]) !== 0);
    if (!hasValues && column === 'daily_H2') return;

    charts.push({
      id: `gf-${column}-per-animal`,
      title,
      data: animalOrder.map((animal) => ({
        type: 'box',
        name: animal,
        x: dailyRows
          .filter((row) => String(row[groupCol]) === animal)
          .map(() => animal),
        y: dailyRows
          .filter((row) => String(row[groupCol]) === animal)
          .map((row) => Number(row[column]))
          .filter(Number.isFinite),
        boxpoints: false
      })),
      layout: {
        margin: { t: 50, l: 50, r: 20, b: 160 },
        height: 420,
        showlegend: false,
        xaxis: categoryAxis({ categoryarray: animalOrder })
      }
    });
  });

  if (treatmentAvailable) {
    const treatments = getOptionList(dailyRows.map((row) => row.Treatment)).filter((value) => value !== 'All');

    charts.push({
      id: 'gf-total-records-per-treatment',
      title: 'Total Records Per Treatment',
      data: [
        {
          type: 'bar',
          marker: { color: '#0f766e' },
          x: treatments,
          y: treatments.map((treatment) =>
            dailyRows
              .filter((row) => String(row.Treatment || 'Unassigned') === treatment)
              .reduce((sum, row) => sum + Number(row.n || 0), 0)
          )
        }
      ],
      layout: {
        margin: { t: 50, l: 50, r: 20, b: 100 },
        height: 400,
        xaxis: categoryAxis({ categoryarray: treatments })
      }
    });

    [
      ['daily_CH4', 'Daily Methane (CH4) by Treatment'],
      ['daily_CO2', 'Daily Carbon Dioxide (CO2) by Treatment'],
      ['daily_O2', 'Daily Oxygen (O2) by Treatment'],
      ['daily_H2', 'Daily Hydrogen (H2) by Treatment']
    ].forEach(([column, title]) => {
      const hasValues = dailyRows.some((row) => Number.isFinite(Number(row[column])) && Number(row[column]) !== 0);
      if (!hasValues && column === 'daily_H2') return;

      charts.push({
        id: `gf-${column}-by-treatment`,
        title,
        data: treatments.map((treatment) => ({
          type: 'box',
          name: treatment,
          x: dailyRows
            .filter((row) => String(row.Treatment || 'Unassigned') === treatment)
            .map(() => treatment),
          y: dailyRows
            .filter((row) => String(row.Treatment || 'Unassigned') === treatment)
            .map((row) => Number(row[column]))
            .filter(Number.isFinite),
          boxpoints: false
        })),
        layout: {
          margin: { t: 50, l: 50, r: 20, b: 100 },
          height: 400,
          showlegend: false,
          xaxis: categoryAxis({ categoryarray: treatments })
        }
      });
    });
  }

  return charts;
}

export default function Home() {
  const [form, setForm] = useState(initialState);
  const [filters, setFilters] = useState(emptyFilters);
  const [mvhFile, setMvhFile] = useState(null);
  const [treatFile, setTreatFile] = useState(null);
  const [cupDropsFile, setCupDropsFile] = useState(null);
  const [useTreatments, setUseTreatments] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState(null);
  const [selectedCharts, setSelectedCharts] = useState([]);
  const [activeTab, setActiveTab] = useState('overview');
  const chartRefs = useRef({});

  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const saved = JSON.parse(raw);
      setForm((current) => ({
        ...current,
        ...saved.form,
        password: ''
      }));
      setUseTreatments(Boolean(saved.useTreatments));
    } catch {
      // Ignore malformed browser storage and continue with defaults.
    }
  }, []);

  useEffect(() => {
    const { password, ...persistedForm } = form;
    window.localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        form: persistedForm,
        useTreatments
      })
    );
  }, [form, useTreatments]);

  const filterOptions = useMemo(() => {
    if (!result) {
      return {
        animals: [],
        units: ['All'],
        treatments: [],
        days: ['All'],
        minDate: '',
        maxDate: ''
      };
    }

    const dates = [
      ...new Set([
        ...result.matchedRows.map((row) => String(row.day_start || '')).filter(Boolean),
        ...(result.cupDropRows || []).map((row) => String(row.FeedDate || '')).filter(Boolean)
      ])
    ].sort();
    return {
      animals: [
        ...new Set([
          ...result.matchedRows.map((row) => animalValueForGroup(row, result.groupCol)).filter(Boolean),
          ...(result.cupDropRows || []).map((row) => animalValueForGroup(row, result.groupCol)).filter(Boolean)
        ])
      ].sort((a, b) => a.localeCompare(b)),
      units: getOptionList([...
        result.matchedRows.map((row) => row.Unit),
        ...(result.cupDropRows || []).map((row) => row.Unit)
      ]),
      treatments: [...new Set([
        ...result.matchedRows.map((row) => row.Treatment || 'Unassigned'),
        ...(result.cupDropRows || []).map((row) => row.Treatment || 'Unassigned')
      ])].sort((a, b) => a.localeCompare(b)),
      days: ['All', ...dates],
      minDate: dates[0] || '',
      maxDate: dates[dates.length - 1] || ''
    };
  }, [result]);

  const filteredMatchedRows = useMemo(() => {
    if (!result?.matchedRows?.length) return [];
    return result.matchedRows.filter((row) => {
      const animalValue = String(row[result.groupCol] ?? '');
      const animalSearch = filters.animalSearch.trim().toLowerCase();
      const treatmentValue = String(row.Treatment || 'Unassigned');
      const rowDate = String(row.day_start || '');
      const selectedAnimalSet = new Set(filters.selectedAnimals || []);
      const selectedTreatmentSet = new Set(filters.selectedTreatments || []);

      if (selectedAnimalSet.size && !selectedAnimalSet.has(animalValue)) return false;
      if (selectedTreatmentSet.size && !selectedTreatmentSet.has(treatmentValue)) return false;
      if (animalSearch && !animalValue.toLowerCase().includes(animalSearch)) return false;
      if (filters.selectedUnit !== 'All' && String(row.Unit || '') !== filters.selectedUnit) return false;
      if (filters.selectedDay !== 'All' && rowDate !== filters.selectedDay) return false;
      if (filters.startDate && rowDate && rowDate < filters.startDate) return false;
      if (filters.endDate && rowDate && rowDate > filters.endDate) return false;
      return true;
    });
  }, [filters, result]);

  const filteredDailyRows = useMemo(() => {
    if (!result?.groupCol || !filteredMatchedRows.length) return [];
    const rows = computeDailyGases(filteredMatchedRows, result.groupCol).map((row) => {
      const matching = filteredMatchedRows.find(
        (item) =>
          String(item[result.groupCol] ?? '') === String(row[result.groupCol] ?? '') &&
          String(item.day_end || '') === String(row.day || '')
      );
      return {
        ...row,
        Treatment: matching?.Treatment || '',
        Unit: matching?.Unit || ''
      };
    });

    return rows;
  }, [filteredMatchedRows, result]);

  const baseFilteredCupDropRows = useMemo(() => {
    if (!result?.cupDropRows?.length) return [];
    return result.cupDropRows.filter((row) => {
      const animalValue = animalValueForGroup(row, result.groupCol);
      const animalSearch = filters.animalSearch.trim().toLowerCase();
      const treatmentValue = String(row.Treatment || 'Unassigned');
      const rowDate = String(row.FeedDate || '');
      const selectedAnimalSet = new Set(filters.selectedAnimals || []);
      const selectedTreatmentSet = new Set(filters.selectedTreatments || []);

      if (selectedAnimalSet.size && !selectedAnimalSet.has(animalValue)) return false;
      if (selectedTreatmentSet.size && !selectedTreatmentSet.has(treatmentValue)) return false;
      if (animalSearch && !animalValue.toLowerCase().includes(animalSearch)) return false;
      if (filters.selectedUnit !== 'All' && String(row.Unit || '') !== filters.selectedUnit) return false;
      if (filters.startDate && rowDate && rowDate < filters.startDate) return false;
      if (filters.endDate && rowDate && rowDate > filters.endDate) return false;
      return true;
    });
  }, [filters, result]);

  const filteredCupDropRows = useMemo(() => {
    if (!baseFilteredCupDropRows.length) return [];
    if (filters.selectedDay === 'All') return baseFilteredCupDropRows;
    return baseFilteredCupDropRows.filter((row) => String(row.FeedDate || '') === filters.selectedDay);
  }, [baseFilteredCupDropRows, filters.selectedDay]);

  const filteredOrder = useMemo(() => {
    if (!result?.groupCol || !filteredDailyRows.length) return [];
    return computeOrder(filteredDailyRows, result.groupCol).map(String);
  }, [filteredDailyRows, result]);

  const filteredUnitSummary = useMemo(() => unitBreakdownTable(filteredMatchedRows), [filteredMatchedRows]);
  const dailySummaryRows = useMemo(
    () => (result?.groupCol ? buildDailySummaryRows(filteredDailyRows, result.groupCol) : []),
    [filteredDailyRows, result]
  );
  const topAnimalRows = useMemo(
    () => (result?.groupCol ? buildTopAnimalTable(filteredDailyRows, result.groupCol) : []),
    [filteredDailyRows, result]
  );
  const treatmentSummaryRows = useMemo(() => buildTreatmentSummaryRows(filteredDailyRows), [filteredDailyRows]);
  const dailyCowUsageRows = useMemo(
    () => (result?.groupCol ? buildDailyCowUsageRows(filteredMatchedRows, result.groupCol) : []),
    [filteredMatchedRows, result]
  );
  const cowsUsingGreenFeedInView = useMemo(() => {
    if (baseFilteredCupDropRows.length) {
      return new Set(
        baseFilteredCupDropRows
          .filter((row) => String(row.EID_norm || '').trim() && String(row.EART || '').trim())
          .map((row) => String(row.EID_norm || '').trim())
      ).size;
    }

    return new Set(filteredMatchedRows.map((row) => String(row[result?.groupCol] ?? '')).filter(Boolean)).size;
  }, [baseFilteredCupDropRows, filteredMatchedRows, result]);
  const apiAnimalsFound = useMemo(
    () => new Set(filteredMatchedRows.map((row) => String(row.EID_norm || row[result?.groupCol] || '')).filter(Boolean)).size,
    [filteredMatchedRows, result]
  );
  const cupDropsAnimalsFound = useMemo(
    () => new Set(baseFilteredCupDropRows.map((row) => String(row.EID_norm || '')).filter(Boolean)).size,
    [baseFilteredCupDropRows]
  );
  const cowsAtFeedDurationInView = useMemo(() => {
    if (!result?.groupCol) return 0;
    return countCowsUsingFeedPeriod(filteredMatchedRows, result.groupCol, form.feedPeriodDurationSec);
  }, [filteredMatchedRows, form.feedPeriodDurationSec, result]);

  const cupDropCohortAnimalIds = useMemo(() => {
    return [
      ...new Set(
        baseFilteredCupDropRows
          .filter((row) => String(row.EID_norm || '').trim() && String(row.EART || '').trim())
          .map((row) => String(row.EID_norm || '').trim())
      )
    ].sort((a, b) => a.localeCompare(b));
  }, [baseFilteredCupDropRows]);

  const cupDropUtilizationRows = useMemo(
    () => buildDailyUtilizationRows(baseFilteredCupDropRows, cupDropCohortAnimalIds, form.minFeedPeriods).filter((row) => filters.selectedDay === 'All' || row.Day === filters.selectedDay),
    [baseFilteredCupDropRows, cupDropCohortAnimalIds, filters.selectedDay, form.minFeedPeriods]
  );

  const cupDropDistributionRows = useMemo(
    () => buildFeedPeriodDistributionRows(baseFilteredCupDropRows, cupDropCohortAnimalIds).filter((row) => filters.selectedDay === 'All' || row.Day === filters.selectedDay),
    [baseFilteredCupDropRows, cupDropCohortAnimalIds, filters.selectedDay]
  );

  const cupDropCowPeriodRows = useMemo(
    () => buildCowPeriodRowsForSelectedDay(filteredCupDropRows, cupDropCohortAnimalIds, filters.selectedDay),
    [cupDropCohortAnimalIds, filteredCupDropRows, filters.selectedDay]
  );
  const treatmentUsageRows = useMemo(
    () => buildTreatmentUsageRows({
      matchedRows: filteredMatchedRows,
      cupDropRows: baseFilteredCupDropRows,
      mvhRows: (result?.mvh || []).filter((row) => {
        const treatmentValue = String(row.Treatment || '').trim();
        if (!treatmentValue) return false;
        if (filters.selectedTreatments?.length && !filters.selectedTreatments.includes(treatmentValue)) return false;
        if (filters.selectedAnimals?.length && !filters.selectedAnimals.includes(animalValueForGroup(row, result.groupCol))) return false;
        return true;
      })
    }),
    [baseFilteredCupDropRows, filteredMatchedRows, filters.selectedAnimals, filters.selectedTreatments, result]
  );
  const treatmentDiagnostics = useMemo(
    () => buildTreatmentDiagnostics({
      cupDropRows: baseFilteredCupDropRows,
      mvhRows: (result?.mvh || []).filter((row) => {
        const treatmentValue = String(row.Treatment || '').trim();
        if (!treatmentValue) return false;
        if (filters.selectedTreatments?.length && !filters.selectedTreatments.includes(treatmentValue)) return false;
        if (filters.selectedAnimals?.length && !filters.selectedAnimals.includes(String(row.EID_norm || ''))) return false;
        return true;
      })
    }),
    [baseFilteredCupDropRows, filters.selectedAnimals, filters.selectedTreatments, result]
  );

  const charts = useMemo(() => {
    const nextCharts = [];
    if (result?.groupCol && filteredMatchedRows.length && filteredDailyRows.length) {
      nextCharts.push(...buildChartSpecs({
        matchedRows: filteredMatchedRows,
        dailyRows: filteredDailyRows,
        groupCol: result.groupCol,
        order: filteredOrder,
        stepHours: Number(form.stepHours),
        treatmentAvailable: filteredMatchedRows.some((row) => row.Treatment),
        selectedDay: filters.selectedDay,
        cupDropRows: filteredCupDropRows
      }));
    }
    if (result?.cupDropRows?.length && cupDropCohortAnimalIds.length) {
      nextCharts.push(...buildCupDropChartSpecs({
        utilizationRows: cupDropUtilizationRows,
        distributionRows: cupDropDistributionRows,
        cowPeriodRows: cupDropCowPeriodRows,
        selectedDay: filters.selectedDay,
        minimumFeedPeriods: Number(form.minFeedPeriods)
      }));
    }
    if ((filters.selectedTreatments?.length || 0) > 1) {
      const allowedChartIds = new Set([
        'gf-total-records-per-treatment',
        'gf-daily_CH4-by-treatment',
        'gf-daily_CO2-by-treatment',
        'gf-daily_O2-by-treatment',
        'gf-daily_H2-by-treatment'
      ]);
      return nextCharts.filter((chart) => allowedChartIds.has(chart.id));
    }

    return nextCharts;
  }, [cupDropCohortAnimalIds, cupDropCowPeriodRows, cupDropDistributionRows, cupDropUtilizationRows, filteredDailyRows, filteredMatchedRows, filteredOrder, filters.selectedDay, filters.selectedTreatments, form.minFeedPeriods, form.stepHours, result]);

  useEffect(() => {
    setSelectedCharts((current) => {
      const ids = charts.map((chart) => chart.id || chart.title);
      if (!ids.length) return [];
      const kept = current.filter((id) => ids.includes(id));
      return kept.length ? kept : ids;
    });
  }, [charts]);

  const visibleCharts = useMemo(
    () => charts.filter((chart) => selectedCharts.includes(chart.id || chart.title)),
    [charts, selectedCharts]
  );

  const unmatchedPreviewRows = useMemo(() => {
    if (!result?.unmatchedRows?.length) return [];
    return result.unmatchedRows.slice(0, 25).map((row) => ({
      Unit: row.Unit || '',
      RFID: row.RFID || '',
      RFID_norm: row.RFID_norm || '',
      StartTime: row.StartTime instanceof Date ? row.StartTime.toISOString() : '',
      AnimalName: row.AnimalName || ''
    }));
  }, [result]);

  const matchRate = useMemo(() => {
    if (!result?.gfRows?.length) return 0;
    return (result.matchedRows.length / result.gfRows.length) * 100;
  }, [result]);

  async function handleProcess() {
    setLoading(true);
    setError('');
    try {
      if (!mvhFile) throw new Error('Upload MVH Research.csv first.');

      const gfResponse = await fetch('/api/greenfeed', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form)
      });

      const gfJson = await gfResponse.json();
      if (!gfResponse.ok) throw new Error(gfJson.error || 'GreenFeed request failed.');

      const gfRows = safeNumericAndDuration(gfJson.rows || []);
      const mvhRows = await parseCsvFile(mvhFile);
      const mvhColumns = Object.fromEntries(Object.keys(mvhRows[0] || {}).map((column) => [column.toLowerCase(), column]));

      if (!mvhColumns.eid || !mvhColumns.pen || !mvhColumns.eart) {
        throw new Error('MVH file must include EID, Pen, and EART columns.');
      }

      const mvh = mvhRows
        .map((row) => ({
          ...row,
          EID_norm: normalizeId(row[mvhColumns.eid]),
          EART: String(row[mvhColumns.eart] ?? '').trim(),
          EART_norm: normalizeId(row[mvhColumns.eart]),
          Pen: String(row[mvhColumns.pen] ?? '').trim()
        }))
        .filter((row) => row.EID_norm || row.EART_norm);

      const mvhMap = new Map(mvh.map((row) => [String(row.EID_norm), row]));

      let matchedRows = gfRows
        .map((row) => ({ ...row, ...(mvhMap.get(String(row.RFID_norm)) || null) }))
        .filter((row) => row.EID_norm);

      const unmatchedRows = gfRows.filter((row) => !mvhMap.has(String(row.RFID_norm)));

      let treatmentMap = new Map();

      let mvhWithTreatments = mvh;

      if (useTreatments && treatFile) {
        const treatmentRows = await parseExcelFile(treatFile);
        const treatmentColumns = Object.fromEntries(
          Object.keys(treatmentRows[0] || {}).map((column) => [column.toLowerCase(), column])
        );

        if (!treatmentColumns.eart || !treatmentColumns.treatment) {
          throw new Error('Treatments file must include EART and Treatment columns.');
        }

        treatmentMap = new Map(
          treatmentRows
            .map((row) => [
              normalizeId(row[treatmentColumns.eart]),
              String(row[treatmentColumns.treatment] ?? '').trim()
            ])
            .filter(([key]) => key)
        );

        matchedRows = matchedRows.map((row) => ({
          ...row,
          Treatment: treatmentMap.get(String(row.EART_norm)) || ''
        }));

        mvhWithTreatments = mvh.map((row) => ({
          ...row,
          Treatment: treatmentMap.get(String(row.EART_norm)) || ''
        }));
      }

      if (!matchedRows.length) {
        throw new Error('0 rows remained after MVH matching. Check GF RFID vs MVH EID.');
      }

      const groupCol = chooseGroupCol(matchedRows);
      const dailyRows = computeDailyGases(matchedRows, groupCol).map((row) => {
        const matching = matchedRows.find(
          (item) =>
            String(item[groupCol] ?? '') === String(row[groupCol] ?? '') &&
            String(item.day_end || '') === String(row.day || '')
        );
        return {
          ...row,
          Treatment: matching?.Treatment || '',
          Unit: matching?.Unit || ''
        };
      });

      const order = computeOrder(dailyRows, groupCol).map(String);
      const unitSummary = unitBreakdownTable(matchedRows);
      const dates = matchedRows.map((row) => String(row.day_start || '')).filter(Boolean).sort();
      const cupDropRows = cupDropsFile
        ? normalizeCupDropRows(await parseCsvFile(cupDropsFile), mvhMap, treatmentMap)
        : [];
      const allDates = [...new Set([...dates, ...cupDropRows.map((row) => String(row.FeedDate || '')).filter(Boolean)])].sort();

      setFilters({
        animalSearch: '',
        selectedAnimals: [],
        selectedUnit: 'All',
        selectedTreatments: [],
        selectedDay: 'All',
        startDate: allDates[0] || dates[0] || '',
        endDate: allDates[allDates.length - 1] || dates[dates.length - 1] || ''
      });

      setResult({
        gfRows,
        mvh: mvhWithTreatments,
        matchedRows,
        unmatchedRows,
        dailyRows,
        cupDropRows,
        groupCol,
        order,
        unitSummary
      });
      setActiveTab('overview');
    } catch (processError) {
      setError(processError.message || 'Unexpected error');
      setResult(null);
    } finally {
      setLoading(false);
    }
  }

  function updateFormField(key, value) {
    setForm((current) => ({ ...current, [key]: value }));
  }

  function updateFilter(key, value) {
    setFilters((current) => ({ ...current, [key]: value }));
  }

  function downloadText(filename, text, mime) {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  async function exportPdf() {
    if (!visibleCharts.length) {
      setError('Select at least one chart before exporting a PDF report.');
      return;
    }

    const pdf = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    let pageIndex = 0;

    for (const chart of visibleCharts) {
      const node = chartRefs.current[chart.title];
      if (!node) continue;
      const image = await toPng(node, { cacheBust: true, pixelRatio: 2 });
      if (pageIndex > 0) pdf.addPage();
      pdf.setFontSize(14);
      pdf.text(chart.title, 34, 34);
      pdf.addImage(image, 'PNG', 24, 48, 790, 420);
      pageIndex += 1;
    }

    pdf.save(form.reportName || 'GreenFeed_Report.pdf');
  }

  return (
    <main style={{ maxWidth: 1540, margin: '0 auto', padding: 24 }}>
      <div
        className="dashboard-hero"
        style={{
          marginBottom: 24,
          background: `linear-gradient(135deg, ${MSU_CREAM} 0%, ${MSU_GREEN_LIGHT} 100%)`,
          borderRadius: 24,
          padding: 24,
          border: '1px solid #d6e3dd'
        }}
      >
        <div style={{ fontSize: 12, letterSpacing: 1.4, textTransform: 'uppercase', color: MSU_GREEN_MID, marginBottom: 8 }}>
          GreenFeed dashboard
        </div>
        <h1 style={{ margin: 0, fontSize: 34 }}>GreenFeed QC Dashboard</h1>
        <p style={{ margin: '10px 0 0', color: '#506760', maxWidth: 760, lineHeight: 1.5 }}>
          Pull GreenFeed data, link with RFIDs, compare treatments, and review summaries.
        </p>
      </div>

      <div className="dashboard-layout" style={{ display: 'grid', gridTemplateColumns: '360px minmax(0, 1fr)', gap: 20, alignItems: 'start' }}>
        <div className="dashboard-sidebar" style={{ ...sectionCard, position: 'sticky', top: 20 }}>
          <h2 style={{ marginTop: 0, marginBottom: 14 }}>Inputs</h2>
          {[
            ['C-LOCK username', 'username'],
            ['C-LOCK password', 'password'],
            ['Unit IDs', 'unitsText'],
            ['Start date', 'startDate'],
            ['End date', 'endDate']
          ].map(([label, key]) => (
            <div key={key} style={{ marginBottom: 12 }}>
              <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>{label}</label>
              <input
                type={key.includes('date') ? 'date' : key === 'password' ? 'password' : 'text'}
                value={form[key]}
                onChange={(event) => updateFormField(key, event.target.value)}
                style={inputStyle}
              />
            </div>
          ))}

          <div style={{ marginBottom: 12 }}>
            <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>MVH Research.csv</label>
            <input type="file" accept=".csv" onChange={(event) => setMvhFile(event.target.files?.[0] || null)} />
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>CupDrops CSV (optional)</label>
            <input type="file" accept=".csv" onChange={(event) => setCupDropsFile(event.target.files?.[0] || null)} />
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input type="checkbox" checked={useTreatments} onChange={(event) => setUseTreatments(event.target.checked)} />
              Use Treatments.xlsx
            </label>
          </div>

          {useTreatments ? (
            <div style={{ marginBottom: 12 }}>
              <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>Treatments.xlsx</label>
              <input type="file" accept=".xlsx,.xls" onChange={(event) => setTreatFile(event.target.files?.[0] || null)} />
            </div>
          ) : null}

          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 16 }}>
            <button onClick={handleProcess} disabled={loading} style={buttonStyle}>
              {loading ? 'Processing...' : 'Process + Preview'}
            </button>
            <button
              onClick={() => {
                setForm(initialState);
                setFilters(emptyFilters);
                setUseTreatments(false);
                setMvhFile(null);
                setTreatFile(null);
                setCupDropsFile(null);
                setResult(null);
                setError('');
              }}
              style={secondaryButtonStyle}
              type="button"
            >
              Reset
            </button>
          </div>

          <div style={{ marginTop: 16, color: '#6a7a8f', fontSize: 13, lineHeight: 1.5 }}>
            Preferences except password are saved in this browser so repeated runs are quicker.
          </div>

          {error ? <div style={{ marginTop: 14, color: '#b42318', fontSize: 14 }}>{error}</div> : null}
        </div>

        <div style={{ display: 'grid', gap: 20 }}>
          {!result ? (
            <div style={sectionCard}>
              <h2 style={{ marginTop: 0 }}>What changed in this version</h2>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 14 }}>
                <div>
                  <div style={{ fontWeight: 700, marginBottom: 6 }}>Interactive filtering</div>
                  <div style={{ color: '#64748b', lineHeight: 1.5 }}>Slice the matched dataset by unit, treatment, date window, and animal ID search after processing.</div>
                </div>
                <div>
                  <div style={{ fontWeight: 700, marginBottom: 6 }}>Coverage review</div>
                  <div style={{ color: '#64748b', lineHeight: 1.5 }}>See unmatched GreenFeed RFID rows so it is easier to audit MVH linkage gaps before exporting reports.</div>
                </div>
                <div>
                  <div style={{ fontWeight: 700, marginBottom: 6 }}>More exports</div>
                  <div style={{ color: '#64748b', lineHeight: 1.5 }}>Download daily summaries, unmatched rows, filtered merged data, unit summaries, and chart PDFs from the same run.</div>
                </div>
              </div>
            </div>
          ) : null}

          {result ? (
            <>
              <div style={sectionCard}>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                  {[
                    ['overview', 'Overview'],
                    ['filters', 'Filters'],
                    ['downloads', 'Downloads'],
                    ['charts', 'Charts']
                  ].map(([key, label]) => (
                    <button
                      key={key}
                      type="button"
                      onClick={() => setActiveTab(key)}
                      style={{
                        ...tabButtonStyle,
                        background: activeTab === key ? MSU_GREEN : '#ffffff',
                        color: activeTab === key ? '#ffffff' : '#49627e',
                        borderColor: activeTab === key ? MSU_GREEN : '#c8d7d0'
                      }}
                    >
                      {label}
                    </button>
                  ))}
                </div>
              </div>

              {activeTab === 'overview' ? (
              <>
              <div style={{ display: 'grid', gap: 20 }}>
                {metricSection('Source data', 'How many records came from each file or API source.', [
                  metricCard('GF rows', result.gfRows.length.toLocaleString(), 'All rows returned from the API'),
                  metricCard('Matched rows', filteredMatchedRows.length.toLocaleString(), `${formatNumber(matchRate, 1)}% of total GF rows matched MVH`),
                  result.cupDropRows?.length ? metricCard('CupDrops rows', result.cupDropRows.length.toLocaleString(), 'Pellet-drop records from the feeder export') : null
                ].filter(Boolean))}

                {metricSection('Animal coverage', 'How many unique animals each source contributes in the current filtered view.', [
                  metricCard('Animals in view', new Set(filteredMatchedRows.map((row) => String(row[result.groupCol] ?? ''))).size.toLocaleString(), 'Based on current filters'),
                  metricCard('API animals found', apiAnimalsFound.toLocaleString(), 'Unique matched animals from the API data'),
                  metricCard('Cows using GF', cowsUsingGreenFeedInView.toLocaleString(), 'Unique cows with at least one filtered GreenFeed visit'),
                  result.cupDropRows?.length ? metricCard('CupDrops animals found', cupDropsAnimalsFound.toLocaleString(), 'Unique animals found in the CupDrops file') : null,
                  result.cupDropRows?.length ? metricCard('CupDrops cohort cows', cupDropCohortAnimalIds.length.toLocaleString(), 'Unique GreenFeed RFIDs with an MVH EID-to-EART match') : null
                ].filter(Boolean))}

                {metricSection('Context', 'Quick context for the currently filtered dataset.', [
                  metricCard('Units in view', new Set(filteredMatchedRows.map((row) => String(row.Unit || ''))).size.toLocaleString(), 'From the filtered matched data'),
                  metricCard('Grouping variable', result.groupCol, 'Used for per-animal summaries')
                ])}
              </div>
              {treatmentUsageRows.length ? (
                <div style={{ ...sectionCard, marginTop: 20 }}>
                  <h2 style={{ marginTop: 0, marginBottom: 8 }}>Treatment summary</h2>
                    <div style={{ color: '#64746f', fontSize: 14, marginBottom: 14 }}>
                      Animals assigned to each treatment, how many are using GreenFeed, and how many are not using it in the current filtered view.
                    </div>
                  <div style={{ color: MSU_GREEN_MID, fontSize: 13, marginBottom: 14 }}>
                    CupDrops matched cows: {treatmentDiagnostics.cupDropMatchedCows.toLocaleString()} | Treatment-assigned cows: {treatmentDiagnostics.treatmentAssignedCows.toLocaleString()} | Treatment-assigned cows missing from CupDrops: {treatmentDiagnostics.treatmentAssignedMissingFromCupDrops.toLocaleString()}
                  </div>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ borderCollapse: 'collapse', width: '100%' }}>
                      <thead>
                        <tr>
                          {['Treatment', 'Animals assigned', 'Using GreenFeed', 'Not using GreenFeed', 'Utilization %'].map((column) => (
                            <th key={column} style={thStyle}>
                              {column}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {treatmentUsageRows.map((row) => (
                          <tr key={row.Treatment}>
                            <td style={tdStyle}>{row.Treatment}</td>
                            <td style={tdStyle}>{row.AnimalsAssigned}</td>
                            <td style={tdStyle}>{row.UsingGreenFeed}</td>
                            <td style={tdStyle}>{row.NotUsingGreenFeed}</td>
                            <td style={tdStyle}>{row.UtilizationPct}%</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : null}
              </>
              ) : null}

              {activeTab === 'filters' ? (
              <div style={sectionCard}>
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
                  <div>
                    <h2 style={{ margin: 0 }}>Filters</h2>
                    <div style={{ marginTop: 6, color: '#64748b', fontSize: 14 }}>
                      Current filters update charts, tables, and downloads together.
                    </div>
                  </div>
                  <button type="button" style={secondaryButtonStyle} onClick={() => setFilters({
                    animalSearch: '',
                    selectedAnimals: [],
                    selectedUnit: 'All',
                    selectedTreatments: [],
                    selectedDay: 'All',
                    startDate: filterOptions.minDate,
                    endDate: filterOptions.maxDate
                  })}>
                    Reset filters
                  </button>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12, marginTop: 16 }}>
                  <div>
                    <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>Animal search</label>
                    <input
                      value={filters.animalSearch}
                      onChange={(event) => updateFilter('animalSearch', event.target.value)}
                      placeholder={`Search ${result.groupCol}`}
                      style={inputStyle}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>
                      Pick cows {filters.selectedAnimals.length ? `(${filters.selectedAnimals.length})` : ''}
                    </label>
                    <select
                      multiple
                      value={filters.selectedAnimals}
                      onChange={(event) => updateFilter('selectedAnimals', Array.from(event.target.selectedOptions).map((option) => option.value))}
                      style={{ ...inputStyle, minHeight: 140 }}
                    >
                      {filterOptions.animals.map((animal) => (
                        <option key={animal} value={animal}>
                          {animal}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>Unit</label>
                    <select value={filters.selectedUnit} onChange={(event) => updateFilter('selectedUnit', event.target.value)} style={inputStyle}>
                      {filterOptions.units.map((unit) => (
                        <option key={unit} value={unit}>
                          {unit}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>
                      Compare treatments {filters.selectedTreatments.length ? `(${filters.selectedTreatments.length})` : ''}
                    </label>
                    <select
                      multiple
                      value={filters.selectedTreatments}
                      onChange={(event) => updateFilter('selectedTreatments', Array.from(event.target.selectedOptions).map((option) => option.value))}
                      style={{ ...inputStyle, minHeight: 140 }}
                    >
                      {filterOptions.treatments.map((treatment) => (
                        <option key={treatment} value={treatment}>
                          {treatment}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>Day to view</label>
                    <select
                      value={filters.selectedDay}
                      onChange={(event) => updateFilter('selectedDay', event.target.value)}
                      style={inputStyle}
                    >
                      {filterOptions.days.map((day) => (
                        <option key={day} value={day}>
                          {day}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>Start day</label>
                    <input type="date" value={filters.startDate} onChange={(event) => updateFilter('startDate', event.target.value)} style={inputStyle} />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: 13, marginBottom: 4 }}>End day</label>
                    <input type="date" value={filters.endDate} onChange={(event) => updateFilter('endDate', event.target.value)} style={inputStyle} />
                  </div>
                </div>
              </div>
              ) : null}

              {activeTab === 'downloads' ? (
              <div style={sectionCard}>
                <h2 style={{ marginTop: 0, marginBottom: 8 }}>Downloads</h2>
                <div style={{ color: '#64746f', fontSize: 14, marginBottom: 18 }}>
                  Export raw data, summaries, feeder utilization results, or a PDF version of the visible charts.
                </div>

                <div style={{ marginBottom: 18 }}>
                  <div style={{ fontWeight: 700, marginBottom: 10 }}>Data tables</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: 14 }}>
                    {downloadCard(
                      'Filtered merged CSV',
                      'All matched GreenFeed rows after your current filters are applied.',
                      'Download filtered merged CSV',
                      () => downloadText(form.csvName, csvFromRows(filteredMatchedRows), 'text/csv')
                    )}
                    {downloadCard(
                      'Unit summary CSV',
                      'Counts and gas-data availability summarized by GreenFeed unit.',
                      'Download unit summary CSV',
                      () => downloadText(form.unitCsvName, csvFromRows(filteredUnitSummary), 'text/csv')
                    )}
                    {downloadCard(
                      'Daily summary CSV',
                      'Per-animal daily gas summaries built from the filtered matched data.',
                      'Download daily summary CSV',
                      () => downloadText(form.dailyCsvName, csvFromRows(dailySummaryRows), 'text/csv')
                    )}
                    {downloadCard(
                      'Unmatched rows CSV',
                      'GreenFeed rows that did not match an MVH animal by RFID/EID.',
                      'Download unmatched rows CSV',
                      () => downloadText(form.unmatchedCsvName, csvFromRows(result.unmatchedRows), 'text/csv')
                    )}
                  </div>
                </div>

                <div style={{ marginBottom: 18 }}>
                  <div style={{ fontWeight: 700, marginBottom: 10 }}>Feeder use</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: 14 }}>
                    {cupDropUtilizationRows.length ? downloadCard(
                      'Daily utilization CSV',
                      'Daily percent of cows using GreenFeed based on the CupDrops feeder file.',
                      'Download daily utilization CSV',
                      () => downloadText(form.cupDropsCsvName, csvFromRows(cupDropUtilizationRows), 'text/csv')
                    ) : null}
                    {downloadCard(
                      'Daily GF cow list',
                      'Which cows had at least one GreenFeed visit on each day in the current view.',
                      'Download daily GF cow list',
                      () => downloadText('GF_daily_cows_using_greenfeed.csv', csvFromRows(dailyCowUsageRows), 'text/csv')
                    )}
                  </div>
                </div>

                <div>
                  <div style={{ fontWeight: 700, marginBottom: 10 }}>Reports</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: 14 }}>
                    {downloadCard(
                      'PDF report',
                      'A PDF export of the charts currently visible in the Charts tab.',
                      'Download PDF report',
                      exportPdf
                    )}
                  </div>
                </div>
              </div>
              ) : null}

              {activeTab === 'charts' ? (
              <>
              <div style={sectionCard}>
                <h2 style={{ marginTop: 0, marginBottom: 8 }}>Chart selection</h2>
                <div style={{ color: '#64746f', fontSize: 14, marginBottom: 14 }}>
                  Choose which charts stay visible and which charts are included in the PDF export.
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 10 }}>
                  {charts.map((chart) => (
                    <label
                      key={chart.id || chart.title}
                      style={{
                        border: '1px solid #d7e1ec',
                        borderRadius: 12,
                        padding: '10px 12px',
                        display: 'flex',
                        gap: 10,
                        alignItems: 'flex-start'
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={selectedCharts.includes(chart.id || chart.title)}
                        onChange={(event) => {
                          if (event.target.checked) {
                            setSelectedCharts((current) => [...current, chart.id || chart.title]);
                          } else {
                            setSelectedCharts((current) => current.filter((id) => id !== (chart.id || chart.title)));
                          }
                        }}
                      />
                      <span>{chart.title}</span>
                    </label>
                  ))}
                </div>
              </div>

              <div style={{ display: 'grid', gap: 20 }}>
                {visibleCharts.length ? (
                  visibleCharts.map((chart) => (
                    <div key={chart.id || chart.title} style={sectionCard}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap', marginBottom: 12 }}>
                        <div style={{ fontWeight: 700 }}>{chart.title}</div>
                        {chart.controls === 'feed-period-threshold' ? (
                          <div style={{ minWidth: 220 }}>
                            <label style={{ display: 'block', fontSize: 12, marginBottom: 4, color: '#64746f' }}>
                              Minimum feed periods
                            </label>
                            <select
                              value={form.minFeedPeriods}
                              onChange={(event) => updateFormField('minFeedPeriods', Number(event.target.value))}
                              style={inputStyle}
                            >
                              {[1, 2, 3, 4, 5, 6, 7, 8].map((count) => (
                                <option key={count} value={count}>
                                  At least {count}
                                </option>
                              ))}
                            </select>
                          </div>
                        ) : null}
                        {chart.controls === 'window-size' ? (
                          <div style={{ minWidth: 220 }}>
                            <label style={{ display: 'block', fontSize: 12, marginBottom: 4, color: '#64746f' }}>
                              Window size
                            </label>
                            <select
                              value={form.stepHours}
                              onChange={(event) => updateFormField('stepHours', Number(event.target.value))}
                              style={inputStyle}
                            >
                              {STEP_OPTIONS.map((hours) => (
                                <option key={hours} value={hours}>
                                  {hours} hours
                                </option>
                              ))}
                            </select>
                          </div>
                        ) : null}
                      </div>
                      <div style={{ overflowX: 'auto' }}>
                        <div
                          ref={(node) => { chartRefs.current[chart.title] = node; }}
                          style={{ minWidth: chart.minWidth || '100%' }}
                        >
                        <Plot
                          data={chart.data}
                          layout={{
                            paper_bgcolor: '#ffffff',
                            plot_bgcolor: '#ffffff',
                            font: { family: 'Arial, sans-serif' },
                            ...chart.layout
                          }}
                          config={{ responsive: true, displaylogo: false }}
                          style={{ width: '100%', minWidth: chart.minWidth || '100%' }}
                        />
                        </div>
                      </div>
                    </div>
                  ))
                ) : (
                  <div style={sectionCard}>Select at least one chart to display the analysis plots.</div>
                )}
              </div>
              </>
              ) : null}
            </>
          ) : null}
        </div>
      </div>
      <style jsx global>{`
        @media (max-width: 1100px) {
          .dashboard-layout {
            grid-template-columns: minmax(0, 1fr) !important;
          }

          .dashboard-sidebar {
            position: static !important;
          }
        }

        @media (max-width: 700px) {
          .dashboard-hero h1 {
            font-size: 28px !important;
          }

          main {
            padding: 16px !important;
          }
        }
      `}</style>
    </main>
  );
}
