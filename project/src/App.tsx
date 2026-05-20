import { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';
import type { SalesRecord, DumpRecord, DateRange, DrillLevel, BreadcrumbItem } from './types';
import {
  filterSalesByDate,
  filterDumpByDate,
  aggregateStoreMetrics,
  aggregateFamilyMetrics,
  aggregateItemMetrics,
  aggregateArticleMetrics,
  formatCurrency,
  formatPct,
  excelSerialToDateString,
} from './dataUtils';

type SortField = 'name' | 'currentSales' | 'compSales' | 'netSalesGrowth' | 'growthPct' | 'currentDump' | 'currentDumpPct' | 'compDump' | 'compDumpPct' | 'deltaDump';
type SortDir = 'asc' | 'desc';

function App() {
  const [salesData, setSalesData] = useState<SalesRecord[]>([]);
  const [dumpData, setDumpData] = useState<DumpRecord[]>([]);
  const [salesFile, setSalesFile] = useState<string>('');
  const [dumpFile, setDumpFile] = useState<string>('');
  const [salesRows, setSalesRows] = useState<number>(0);
  const [dumpRows, setDumpRows] = useState<number>(0);
  const [dataLoaded, setDataLoaded] = useState(false);

  const [currentPeriod, setCurrentPeriod] = useState<DateRange>({ start: '', end: '' });
  const [compPeriod, setCompPeriod] = useState<DateRange>({ start: '', end: '' });

  const [drill, setDrill] = useState<DrillLevel>({ type: 'area' });
  const [breadcrumbs, setBreadcrumbs] = useState<BreadcrumbItem[]>([
    { label: 'Area Overview', level: { type: 'area' } },
  ]);

  const [sortField, setSortField] = useState<SortField>('currentSales');
  const [sortDir, setSortDir] = useState<SortDir>('desc');

  const barChartRef = useRef<HTMLCanvasElement>(null);
  const donutChartRef = useRef<HTMLCanvasElement>(null);
  const barChartInstance = useRef<Chart | null>(null);
  const donutChartInstance = useRef<Chart | null>(null);

  const handleSalesUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setSalesFile(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target?.result;
      const wb = XLSX.read(data, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<SalesRecord>(ws);
      const filtered = json.filter((r) => r.family_name === 'Fruit' || r.family_name === 'Vegetables');
      setSalesData(filtered);
      setSalesRows(filtered.length);
      if (dumpData.length > 0) setDataLoaded(true);
    };
    reader.readAsBinaryString(file);
  }, [dumpData.length]);

  const handleDumpUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setDumpFile(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target?.result;
      const wb = XLSX.read(data, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<DumpRecord>(ws);
      const filtered = json.filter((r) => r.family_name === 'Fruit' || r.family_name === 'Vegetables');
      setDumpData(filtered);
      setDumpRows(filtered.length);
      if (salesData.length > 0) setDataLoaded(true);
    };
    reader.readAsBinaryString(file);
  }, [salesData.length]);

  useEffect(() => {
    if (salesData.length > 0 && dumpData.length > 0 && !dataLoaded) {
      setDataLoaded(true);
    }
  }, [salesData, dumpData, dataLoaded]);

  useEffect(() => {
    if (salesData.length > 0) {
      const dates = salesData.map((r) => r.b_date);
      const max = Math.max(...dates);
      const maxDate = new Date(excelSerialToDateString(max));
      const weekAgo = new Date(maxDate);
      weekAgo.setDate(weekAgo.getDate() - 6);
      const twoWeeksAgo = new Date(weekAgo);
      twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 6);

      const fmt = (d: Date) => d.toISOString().split('T')[0];
      setCurrentPeriod({ start: fmt(weekAgo), end: fmt(maxDate) });
      setCompPeriod({ start: fmt(twoWeeksAgo), end: new Date(weekAgo.getTime() - 86400000).toISOString().split('T')[0] });
    }
  }, [salesData.length]);

  const storeMetrics = useMemo(
    () => aggregateStoreMetrics(
      filterSalesByDate(salesData, currentPeriod),
      filterSalesByDate(salesData, compPeriod),
      filterDumpByDate(dumpData, currentPeriod),
      filterDumpByDate(dumpData, compPeriod)
    ),
    [salesData, dumpData, currentPeriod, compPeriod]
  );

  const familyMetrics = useMemo(() => {
    if (!drill.store) return [];
    const storeSales = salesData.filter((r) => r.loc_name === drill.store);
    const storeDump = dumpData.filter((r) => r.loc_name === drill.store);
    return aggregateFamilyMetrics(
      filterSalesByDate(storeSales, currentPeriod),
      filterSalesByDate(storeSales, compPeriod),
      filterDumpByDate(storeDump, currentPeriod),
      filterDumpByDate(storeDump, compPeriod)
    );
  }, [salesData, dumpData, drill.store, currentPeriod, compPeriod]);

  const itemMetrics = useMemo(() => {
    if (!drill.store || !drill.family) return [];
    const filtered = salesData.filter((r) => r.loc_name === drill.store && r.family_name === drill.family);
    const filteredDump = dumpData.filter((r) => r.loc_name === drill.store && r.family_name === drill.family);
    return aggregateItemMetrics(
      filterSalesByDate(filtered, currentPeriod),
      filterSalesByDate(filtered, compPeriod),
      filterDumpByDate(filteredDump, currentPeriod),
      filterDumpByDate(filteredDump, compPeriod)
    );
  }, [salesData, dumpData, drill.store, drill.family, currentPeriod, compPeriod]);

  const articleMetrics = useMemo(() => {
    if (!drill.store || !drill.family || !drill.item) return [];
    const filtered = salesData.filter((r) => r.loc_name === drill.store && r.family_name === drill.family && r.item_desc === drill.item);
    const filteredDump = dumpData.filter((r) => r.loc_name === drill.store && r.family_name === drill.family && r.item_desc === drill.item);
    return aggregateArticleMetrics(
      filterSalesByDate(filtered, currentPeriod),
      filterSalesByDate(filtered, compPeriod),
      filterDumpByDate(filteredDump, currentPeriod),
      filterDumpByDate(filteredDump, compPeriod)
    );
  }, [salesData, dumpData, drill, currentPeriod, compPeriod]);

  const currentGridData = useMemo(() => {
    if (drill.type === 'area') return storeMetrics.map((m) => ({ key: m.name, label: m.name, ...m }));
    if (drill.type === 'store') return familyMetrics.map((m) => ({ key: m.family, label: m.family, ...m }));
    if (drill.type === 'family') return itemMetrics.map((m) => ({ key: m.itemDesc, label: m.itemDesc, ...m }));
    return articleMetrics.map((m) => ({ key: `${m.itemNo}`, label: m.itemDesc, ...m }));
  }, [drill.type, storeMetrics, familyMetrics, itemMetrics, articleMetrics]);

  const sortedGridData = useMemo(() => {
    const sorted = [...currentGridData].sort((a, b) => {
      const aVal = a[sortField as keyof typeof a] ?? 0;
      const bVal = b[sortField as keyof typeof b] ?? 0;
      if (typeof aVal === 'string' && typeof bVal === 'string') {
        return sortDir === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
      }
      return sortDir === 'asc' ? (aVal as number) - (bVal as number) : (bVal as number) - (aVal as number);
    });
    return sorted;
  }, [currentGridData, sortField, sortDir]);

  const totals = useMemo(() => {
    const cs = currentGridData.reduce((s, r) => s + (r.currentSales || 0), 0);
    const ps = currentGridData.reduce((s, r) => s + (r.compSales || 0), 0);
    const cd = currentGridData.reduce((s, r) => s + (r.currentDump || 0), 0);
    const pd = currentGridData.reduce((s, r) => s + (r.compDump || 0), 0);
    return {
      currentSales: cs,
      compSales: ps,
      netSalesGrowth: cs - ps,
      growthPct: ps > 0 ? ((cs - ps) / ps) * 100 : 0,
      currentDump: cd,
      currentDumpPct: cs > 0 ? (cd / cs) * 100 : 0,
      compDump: pd,
      compDumpPct: ps > 0 ? (pd / ps) * 100 : 0,
      deltaDump: cd - pd,
    };
  }, [currentGridData]);

  const kpiData = useMemo(() => {
    if (storeMetrics.length === 0) return { best: null, worst: null, highDump: null };
    const best = storeMetrics.reduce((a, b) => (b.growthPct > a.growthPct ? b : a), storeMetrics[0]);
    const worst = storeMetrics.reduce((a, b) => (b.growthPct < a.growthPct ? b : a), storeMetrics[0]);
    const highDump = storeMetrics.reduce((a, b) => (b.currentDumpPct > a.currentDumpPct ? b : a), storeMetrics[0]);
    return { best, worst, highDump };
  }, [storeMetrics]);

  const handleSort = useCallback((field: SortField) => {
    if (sortField === field) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortField(field);
      setSortDir('desc');
    }
  }, [sortField]);

  const handleRowClick = useCallback(
    (row: { key: string; label: string }) => {
      if (drill.type === 'area') {
        const newDrill: DrillLevel = { type: 'store', store: row.key };
        const newCrumbs: BreadcrumbItem[] = [
          { label: 'Area Overview', level: { type: 'area' } },
          { label: row.label, level: newDrill },
        ];
        setDrill(newDrill);
        setBreadcrumbs(newCrumbs);
      } else if (drill.type === 'store') {
        const newDrill: DrillLevel = { type: 'family', store: drill.store, family: row.key };
        const newCrumbs: BreadcrumbItem[] = [
          { label: 'Area Overview', level: { type: 'area' } },
          { label: drill.store!, level: { type: 'store', store: drill.store } },
          { label: row.label, level: newDrill },
        ];
        setDrill(newDrill);
        setBreadcrumbs(newCrumbs);
      } else if (drill.type === 'family') {
        const newDrill: DrillLevel = { type: 'item', store: drill.store, family: drill.family, item: row.key };
        const newCrumbs: BreadcrumbItem[] = [
          { label: 'Area Overview', level: { type: 'area' } },
          { label: drill.store!, level: { type: 'store', store: drill.store } },
          { label: drill.family!, level: { type: 'family', store: drill.store, family: drill.family } },
          { label: row.label, level: newDrill },
        ];
        setDrill(newDrill);
        setBreadcrumbs(newCrumbs);
      }
    },
    [drill]
  );

  const handleBreadcrumbClick = useCallback((level: DrillLevel) => {
    setDrill(level);
    const idx = level.type === 'area' ? 0 : level.type === 'store' ? 1 : level.type === 'family' ? 2 : 3;
    setBreadcrumbs((prev) => prev.slice(0, idx + 1));
  }, []);

  useEffect(() => {
    if (!barChartRef.current || !donutChartRef.current) return;
    if (!dataLoaded) return;

    const labels = sortedGridData.map((r) => {
      const l = r.label || r.key;
      return l.length > 20 ? l.substring(0, 18) + '..' : l;
    });
    const currentSalesVals = sortedGridData.map((r) => Math.round(r.currentSales || 0));
    const compSalesVals = sortedGridData.map((r) => Math.round(r.compSales || 0));
    const dumpVals = sortedGridData.map((r) => Math.round(r.currentDump || 0));

    if (barChartInstance.current) barChartInstance.current.destroy();
    barChartInstance.current = new Chart(barChartRef.current, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          {
            label: 'Current Sales',
            data: currentSalesVals,
            backgroundColor: '#0d9488',
            borderRadius: 4,
            barPercentage: 0.7,
            categoryPercentage: 0.7,
          },
          {
            label: 'Comp Sales',
            data: compSalesVals,
            backgroundColor: '#94a3b8',
            borderRadius: 4,
            barPercentage: 0.7,
            categoryPercentage: 0.7,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { font: { size: 11 }, usePointStyle: true, pointStyle: 'rectRounded' } },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${formatCurrency(ctx.raw as number)}`,
            },
          },
        },
        scales: {
          x: { ticks: { font: { size: 10 }, maxRotation: 45 }, grid: { display: false } },
          y: {
            ticks: { font: { size: 10 }, callback: (v) => formatCurrency(v as number) },
            grid: { color: '#f1f5f9' },
          },
        },
      },
    });

    if (donutChartInstance.current) donutChartInstance.current.destroy();
    const positiveIndices = dumpVals.reduce((acc, v, i) => { if (v > 0) acc.push(i); return acc; }, [] as number[]);
    const positiveLabels = positiveIndices.map((i) => labels[i]);
    const positiveDumpVals = positiveIndices.map((i) => dumpVals[i]);

    const donutColors = [
      '#0d9488', '#f59e0b', '#dc2626', '#6366f1', '#ec4899', '#8b5cf6',
      '#06b6d4', '#84cc16', '#f97316', '#14b8a6', '#a855f7', '#ef4444',
    ];

    donutChartInstance.current = new Chart(donutChartRef.current, {
      type: 'doughnut',
      data: {
        labels: positiveLabels,
        datasets: [{
          data: positiveDumpVals,
          backgroundColor: donutColors.slice(0, positiveLabels.length),
          borderWidth: 2,
          borderColor: '#fff',
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '55%',
        plugins: {
          legend: { position: 'right', labels: { font: { size: 10 }, padding: 8, usePointStyle: true, pointStyle: 'circle' } },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const total = positiveDumpVals.reduce((a, b) => a + b, 0);
                const pct = total > 0 ? ((ctx.raw as number) / total) * 100 : 0;
                return `${ctx.label}: ${formatCurrency(ctx.raw as number)} (${pct.toFixed(1)}%)`;
              },
            },
          },
        },
      },
    });

    return () => {
      if (barChartInstance.current) barChartInstance.current.destroy();
      if (donutChartInstance.current) donutChartInstance.current.destroy();
    };
  }, [sortedGridData, dataLoaded]);

  const SortIcon = ({ field }: { field: SortField }) => (
    <span className="ml-1 text-[10px] opacity-70">
      {sortField === field ? (sortDir === 'asc' ? '▲' : '▼') : '⇅'}
    </span>
  );

  const ValueCell = ({ value, isCurrency, isPct, invertColor }: { value: number; isCurrency?: boolean; isPct?: boolean; invertColor?: boolean }) => {
    const display = isCurrency ? formatCurrency(value) : isPct ? formatPct(value) : String(value);
    let color = '';
    if (invertColor) {
      if (value < 0) color = 'text-green-700';
      else if (value > 0) color = 'text-red-600';
    } else if (isCurrency || isPct) {
      if (value > 0) color = 'text-green-700';
      else if (value < 0) color = 'text-red-600';
    }
    return <span className={color}>{display}</span>;
  };

  return (
    <div className="min-h-screen bg-[#f8fafc]">
      {/* Header */}
      <header className="bg-[#0f172a] text-white px-6 py-4 flex items-center justify-between shadow-lg">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-teal-500 flex items-center justify-center">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M3 3v18h18" /><path d="M7 16l4-8 4 5 5-9" />
            </svg>
          </div>
          <div>
            <h1 className="text-lg font-bold tracking-tight">More Retail - F&V Area Manager Dashboard</h1>
            <p className="text-xs text-slate-400">12 Stores | Telangana Region</p>
          </div>
        </div>
        {dataLoaded && (
          <div className="flex items-center gap-4 text-xs text-slate-300">
            <span>Sales: {salesRows.toLocaleString()} rows</span>
            <span className="text-slate-600">|</span>
            <span>Dump: {dumpRows.toLocaleString()} rows</span>
          </div>
        )}
      </header>

      {!dataLoaded ? (
        /* Upload Screen */
        <div className="max-w-3xl mx-auto mt-20 px-6">
          <div className="text-center mb-10">
            <h2 className="text-2xl font-bold text-slate-800 mb-2">Upload Your Data Files</h2>
            <p className="text-slate-500 text-sm">Upload both F&V Sales and F&V Dump files to get started. Accepts .xlsx and .csv formats.</p>
          </div>
          <div className="grid grid-cols-2 gap-6">
            {/* Sales Upload */}
            <div className="upload-zone rounded-xl p-8 text-center bg-white">
              <div className="w-14 h-14 mx-auto mb-4 rounded-full bg-teal-50 flex items-center justify-center">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="12" y1="18" x2="12" y2="12" /><polyline points="9 15 12 12 15 15" />
                </svg>
              </div>
              <h3 className="font-semibold text-slate-800 mb-1">F&V Sales File</h3>
              <p className="text-xs text-slate-400 mb-4">b_date, loc_name, family_name, sales_value...</p>
              <label className="inline-block px-5 py-2.5 bg-teal-600 text-white text-sm font-medium rounded-lg cursor-pointer hover:bg-teal-700 transition-colors">
                Choose File
                <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleSalesUpload} />
              </label>
              {salesFile && (
                <div className="mt-3 text-xs text-teal-700 font-medium">
                  {salesFile} ({salesRows.toLocaleString()} rows)
                </div>
              )}
            </div>

            {/* Dump Upload */}
            <div className="upload-zone rounded-xl p-8 text-center bg-white">
              <div className="w-14 h-14 mx-auto mb-4 rounded-full bg-amber-50 flex items-center justify-center">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="12" y1="18" x2="12" y2="12" /><polyline points="9 15 12 12 15 15" />
                </svg>
              </div>
              <h3 className="font-semibold text-slate-800 mb-1">F&V Dump File</h3>
              <p className="text-xs text-slate-400 mb-4">dump_date, loc_name, family_name, dump_value...</p>
              <label className="inline-block px-5 py-2.5 bg-amber-500 text-white text-sm font-medium rounded-lg cursor-pointer hover:bg-amber-600 transition-colors">
                Choose File
                <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleDumpUpload} />
              </label>
              {dumpFile && (
                <div className="mt-3 text-xs text-amber-700 font-medium">
                  {dumpFile} ({dumpRows.toLocaleString()} rows)
                </div>
              )}
            </div>
          </div>
        </div>
      ) : (
        /* Dashboard */
        <div className="px-6 py-4 max-w-[1600px] mx-auto">
          {/* Date Filter Bar */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 mb-4 flex items-center gap-6 flex-wrap">
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-teal-700 bg-teal-50 px-2.5 py-1 rounded-md">CURRENT PERIOD</span>
              <input type="date" className="date-input" value={currentPeriod.start}
                onChange={(e) => setCurrentPeriod((p) => ({ ...p, start: e.target.value }))} />
              <span className="text-slate-400 text-sm">to</span>
              <input type="date" className="date-input" value={currentPeriod.end}
                onChange={(e) => setCurrentPeriod((p) => ({ ...p, end: e.target.value }))} />
            </div>
            <div className="w-px h-8 bg-slate-200" />
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-slate-600 bg-slate-100 px-2.5 py-1 rounded-md">COMPARISON PERIOD</span>
              <input type="date" className="date-input" value={compPeriod.start}
                onChange={(e) => setCompPeriod((p) => ({ ...p, start: e.target.value }))} />
              <span className="text-slate-400 text-sm">to</span>
              <input type="date" className="date-input" value={compPeriod.end}
                onChange={(e) => setCompPeriod((p) => ({ ...p, end: e.target.value }))} />
            </div>
            <div className="ml-auto text-xs text-slate-400">
              {currentPeriod.start && (
                <span>
                  {new Date(currentPeriod.start).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' })} - {new Date(currentPeriod.end).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' })}
                  {' vs '}
                  {new Date(compPeriod.start).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' })} - {new Date(compPeriod.end).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' })}
                </span>
              )}
            </div>
          </div>

          {/* KPI Strip */}
          <div className="grid grid-cols-6 gap-3 mb-4">
            <div className="kpi-card">
              <div className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider mb-1">Current Sales</div>
              <div className="text-xl font-bold text-slate-800">{formatCurrency(totals.currentSales)}</div>
            </div>
            <div className="kpi-card">
              <div className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider mb-1">Comp Sales</div>
              <div className="text-xl font-bold text-slate-600">{formatCurrency(totals.compSales)}</div>
            </div>
            <div className="kpi-card">
              <div className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider mb-1">Growth %</div>
              <div className={`text-xl font-bold ${totals.growthPct >= 0 ? 'text-green-700' : 'text-red-600'}`}>
                {formatPct(totals.growthPct)}
              </div>
            </div>
            <div className="kpi-card">
              <div className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider mb-1">Current Dump %</div>
              <div className={`text-xl font-bold ${totals.currentDumpPct > 10 ? 'text-red-600' : 'text-amber-600'}`}>
                {formatPct(totals.currentDumpPct)}
              </div>
            </div>
            <div className="kpi-card">
              <div className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider mb-1">Best Store</div>
              <div className="text-sm font-bold text-green-700 truncate" title={kpiData.best?.name}>{kpiData.best?.name?.split(' - ')[0] || '-'}</div>
              <div className="text-xs text-green-600">{kpiData.best ? formatPct(kpiData.best.growthPct) : '-'}</div>
            </div>
            <div className="kpi-card">
              <div className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider mb-1">Highest Dump</div>
              <div className="text-sm font-bold text-red-600 truncate" title={kpiData.highDump?.name}>{kpiData.highDump?.name?.split(' - ')[0] || '-'}</div>
              <div className="text-xs text-red-500">{kpiData.highDump ? formatPct(kpiData.highDump.currentDumpPct) : '-'}</div>
            </div>
          </div>

          {/* Charts */}
          <div className="grid grid-cols-5 gap-4 mb-4">
            <div className="chart-container col-span-3">
              <h3 className="text-sm font-semibold text-slate-700 mb-3">Sales Comparison: Current vs Comp Period</h3>
              <div style={{ height: '280px' }}>
                <canvas ref={barChartRef} />
              </div>
            </div>
            <div className="chart-container col-span-2">
              <h3 className="text-sm font-semibold text-slate-700 mb-3">Current Dump Distribution</h3>
              <div style={{ height: '280px' }}>
                <canvas ref={donutChartRef} />
              </div>
            </div>
          </div>

          {/* Breadcrumb */}
          <div className="flex items-center gap-1 mb-2 text-sm">
            {breadcrumbs.map((crumb, i) => (
              <span key={i} className="flex items-center gap-1">
                {i > 0 && <span className="text-slate-300 mx-1">/</span>}
                <span
                  className={`breadcrumb-item ${i === breadcrumbs.length - 1 ? 'font-semibold text-teal-700' : 'text-slate-500'}`}
                  onClick={() => handleBreadcrumbClick(crumb.level)}
                >
                  {crumb.label.length > 30 ? crumb.label.substring(0, 28) + '..' : crumb.label}
                </span>
              </span>
            ))}
          </div>

          {/* Data Table */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="overflow-auto max-h-[520px]">
              <table className="metric-table w-full">
                <thead>
                  <tr>
                    <th className="sort-header text-left" onClick={() => handleSort('name')}>
                      {drill.type === 'area' ? 'Store Location' : drill.type === 'store' ? 'Family / Category' : drill.type === 'family' ? 'Item' : 'Article / PLU'}
                      <SortIcon field="name" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('currentSales')}>
                      Current Sales <SortIcon field="currentSales" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('compSales')}>
                      Comp Sales <SortIcon field="compSales" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('netSalesGrowth')}>
                      Net Growth <SortIcon field="netSalesGrowth" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('growthPct')}>
                      Growth % <SortIcon field="growthPct" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('currentDump')}>
                      Curr Dump <SortIcon field="currentDump" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('currentDumpPct')}>
                      Curr Dump % <SortIcon field="currentDumpPct" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('compDump')}>
                      Comp Dump <SortIcon field="compDump" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('compDumpPct')}>
                      Comp Dump % <SortIcon field="compDumpPct" />
                    </th>
                    <th className="sort-header text-right" onClick={() => handleSort('deltaDump')}>
                      Delta Dump <SortIcon field="deltaDump" />
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {sortedGridData.map((row) => (
                    <tr key={row.key} onClick={() => handleRowClick(row)} className={`fade-in ${row.currentDumpPct > 10 ? 'bg-red-50/60' : ''}`}>
                      <td className="font-medium text-slate-800 max-w-[240px] truncate" title={row.label}>
                        {drill.type === 'item' ? (
                          <span>
                            <span className="text-slate-400 text-xs mr-1">{(row as any).itemNo}</span>
                            {row.label}
                          </span>
                        ) : row.label}
                      </td>
                      <td className="text-right font-medium">{formatCurrency(row.currentSales)}</td>
                      <td className="text-right text-slate-500">{formatCurrency(row.compSales)}</td>
                      <td className="text-right"><ValueCell value={row.netSalesGrowth} isCurrency /></td>
                      <td className="text-right"><ValueCell value={row.growthPct} isPct /></td>
                      <td className="text-right">{formatCurrency(row.currentDump)}</td>
                      <td className="text-right">
                        <span className={row.currentDumpPct > 10 ? 'text-red-600 font-bold' : row.currentDumpPct > 5 ? 'text-amber-600 font-medium' : 'text-slate-600'}>
                          {formatPct(row.currentDumpPct)}
                        </span>
                      </td>
                      <td className="text-right text-slate-500">{formatCurrency(row.compDump)}</td>
                      <td className="text-right text-slate-500">{formatPct(row.compDumpPct)}</td>
                      <td className="text-right"><ValueCell value={row.deltaDump} isCurrency invertColor /></td>
                    </tr>
                  ))}
                  {/* Totals Row */}
                  <tr className="total-row">
                    <td className="font-bold text-slate-800">TOTAL</td>
                    <td className="text-right font-bold">{formatCurrency(totals.currentSales)}</td>
                    <td className="text-right font-bold text-slate-600">{formatCurrency(totals.compSales)}</td>
                    <td className="text-right font-bold"><ValueCell value={totals.netSalesGrowth} isCurrency /></td>
                    <td className="text-right font-bold"><ValueCell value={totals.growthPct} isPct /></td>
                    <td className="text-right font-bold">{formatCurrency(totals.currentDump)}</td>
                    <td className="text-right font-bold">
                      <span className={totals.currentDumpPct > 10 ? 'text-red-600' : 'text-amber-600'}>
                        {formatPct(totals.currentDumpPct)}
                      </span>
                    </td>
                    <td className="text-right font-bold text-slate-600">{formatCurrency(totals.compDump)}</td>
                    <td className="text-right font-bold text-slate-600">{formatPct(totals.compDumpPct)}</td>
                    <td className="text-right font-bold"><ValueCell value={totals.deltaDump} isCurrency invertColor /></td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* Footer hint */}
          <div className="mt-3 text-xs text-slate-400 text-center">
            Click any row to drill down. Click breadcrumbs to navigate back. Click column headers to sort.
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
