import type { SalesRecord, DumpRecord, DateRange, StoreMetrics, DrillLevel } from './types';

const EXCEL_EPOCH = new Date(1899, 11, 30);

export function excelSerialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH.getTime() + serial * 86400000);
}

export function excelSerialToDateString(serial: number): string {
  const d = excelSerialToDate(serial);
  return d.toISOString().split('T')[0];
}

export function dateToExcelSerial(dateStr: string): number {
  const d = new Date(dateStr + 'T00:00:00');
  return Math.round((d.getTime() - EXCEL_EPOCH.getTime()) / 86400000);
}

export function filterSalesByDate(
  sales: SalesRecord[],
  range: DateRange
): SalesRecord[] {
  const startSerial = dateToExcelSerial(range.start);
  const endSerial = dateToExcelSerial(range.end);
  return sales.filter(
    (r) => r.b_date >= startSerial && r.b_date <= endSerial
  );
}

export function filterDumpByDate(
  dump: DumpRecord[],
  range: DateRange
): DumpRecord[] {
  const startSerial = dateToExcelSerial(range.start);
  const endSerial = dateToExcelSerial(range.end);
  return dump.filter(
    (r) => r.dump_date >= startSerial && r.dump_date <= endSerial
  );
}

export function aggregateStoreMetrics(
  currentSales: SalesRecord[],
  compSales: SalesRecord[],
  currentDump: DumpRecord[],
  compDump: DumpRecord[]
): StoreMetrics[] {
  const stores = new Set<string>();
  currentSales.forEach((r) => stores.add(r.loc_name));
  compSales.forEach((r) => stores.add(r.loc_name));
  currentDump.forEach((r) => stores.add(r.loc_name));
  compDump.forEach((r) => stores.add(r.loc_name));

  const sumBy = (records: (SalesRecord | DumpRecord)[], field: 'loc_name') => {
    const map: Record<string, number> = {};
    records.forEach((r) => {
      const key = r[field];
      const val = 'sales_value' in r ? (r as SalesRecord).sales_value : (r as DumpRecord).dump_value;
      map[key] = (map[key] || 0) + val;
    });
    return map;
  };

  const csMap = sumBy(currentSales, 'loc_name');
  const psMap = sumBy(compSales, 'loc_name');
  const cdMap = sumBy(currentDump, 'loc_name');
  const pdMap = sumBy(compDump, 'loc_name');

  const sortedStores = [...stores].sort();

  return sortedStores.map((name) => {
    const currentSalesVal = csMap[name] || 0;
    const compSalesVal = psMap[name] || 0;
    const currentDumpVal = cdMap[name] || 0;
    const compDumpVal = pdMap[name] || 0;
    const netGrowth = currentSalesVal - compSalesVal;
    const growthPct = compSalesVal > 0 ? (netGrowth / compSalesVal) * 100 : 0;
    const currentDumpPct = currentSalesVal > 0 ? (currentDumpVal / currentSalesVal) * 100 : 0;
    const compDumpPct = compSalesVal > 0 ? (compDumpVal / compSalesVal) * 100 : 0;

    return {
      name,
      currentSales: currentSalesVal,
      compSales: compSalesVal,
      netSalesGrowth: netGrowth,
      growthPct,
      currentDump: currentDumpVal,
      currentDumpPct,
      compDump: compDumpVal,
      compDumpPct,
      deltaDump: currentDumpVal - compDumpVal,
    };
  });
}

export interface FamilyMetrics {
  family: string;
  currentSales: number;
  compSales: number;
  netSalesGrowth: number;
  growthPct: number;
  currentDump: number;
  currentDumpPct: number;
  compDump: number;
  compDumpPct: number;
  deltaDump: number;
}

export function aggregateFamilyMetrics(
  currentSales: SalesRecord[],
  compSales: SalesRecord[],
  currentDump: DumpRecord[],
  compDump: DumpRecord[]
): FamilyMetrics[] {
  const families = new Set<string>();
  currentSales.forEach((r) => families.add(r.family_name));
  compSales.forEach((r) => families.add(r.family_name));
  currentDump.forEach((r) => families.add(r.family_name));
  compDump.forEach((r) => families.add(r.family_name));

  const sumSalesByFamily = (records: SalesRecord[]) => {
    const map: Record<string, number> = {};
    records.forEach((r) => {
      map[r.family_name] = (map[r.family_name] || 0) + r.sales_value;
    });
    return map;
  };

  const sumDumpByFamily = (records: DumpRecord[]) => {
    const map: Record<string, number> = {};
    records.forEach((r) => {
      map[r.family_name] = (map[r.family_name] || 0) + r.dump_value;
    });
    return map;
  };

  const csMap = sumSalesByFamily(currentSales);
  const psMap = sumSalesByFamily(compSales);
  const cdMap = sumDumpByFamily(currentDump);
  const pdMap = sumDumpByFamily(compDump);

  return [...families].sort().map((family) => {
    const cs = csMap[family] || 0;
    const ps = psMap[family] || 0;
    const cd = cdMap[family] || 0;
    const pd = pdMap[family] || 0;
    return {
      family,
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
  });
}

export interface ItemMetrics {
  itemDesc: string;
  itemNo: number;
  className: string;
  currentSales: number;
  compSales: number;
  netSalesGrowth: number;
  growthPct: number;
  currentDump: number;
  currentDumpPct: number;
  compDump: number;
  compDumpPct: number;
  deltaDump: number;
}

export function aggregateItemMetrics(
  currentSales: SalesRecord[],
  compSales: SalesRecord[],
  currentDump: DumpRecord[],
  compDump: DumpRecord[]
): ItemMetrics[] {
  const items = new Map<string, { desc: string; no: number; cls: string }>();
  currentSales.forEach((r) => {
    const key = `${r.item_no}`;
    if (!items.has(key)) items.set(key, { desc: r.item_desc, no: r.item_no, cls: r.class_name });
  });
  compSales.forEach((r) => {
    const key = `${r.item_no}`;
    if (!items.has(key)) items.set(key, { desc: r.item_desc, no: r.item_no, cls: r.class_name });
  });

  const sumSalesByItem = (records: SalesRecord[]) => {
    const map: Record<string, number> = {};
    records.forEach((r) => {
      const key = `${r.item_no}`;
      map[key] = (map[key] || 0) + r.sales_value;
    });
    return map;
  };

  const sumDumpByItem = (records: DumpRecord[]) => {
    const map: Record<string, number> = {};
    records.forEach((r) => {
      const key = `${r.item}`;
      map[key] = (map[key] || 0) + r.dump_value;
    });
    return map;
  };

  const csMap = sumSalesByItem(currentSales);
  const psMap = sumSalesByItem(compSales);
  const cdMap = sumDumpByItem(currentDump);
  const pdMap = sumDumpByItem(compDump);

  const allKeys = new Set<string>();
  currentSales.forEach((r) => allKeys.add(`${r.item_no}`));
  compSales.forEach((r) => allKeys.add(`${r.item_no}`));
  currentDump.forEach((r) => allKeys.add(`${r.item}`));
  compDump.forEach((r) => allKeys.add(`${r.item}`));

  return [...allKeys].sort().map((key) => {
    const info = items.get(key) || { desc: key, no: Number(key), cls: '' };
    const cs = csMap[key] || 0;
    const ps = psMap[key] || 0;
    const cd = cdMap[key] || 0;
    const pd = pdMap[key] || 0;
    return {
      itemDesc: info.desc,
      itemNo: info.no,
      className: info.cls,
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
  });
}

export interface ArticleMetrics {
  itemNo: number;
  itemDesc: string;
  className: string;
  currentSales: number;
  compSales: number;
  netSalesGrowth: number;
  growthPct: number;
  currentDump: number;
  currentDumpPct: number;
  compDump: number;
  compDumpPct: number;
  deltaDump: number;
  currentSaleQty: number;
  compSaleQty: number;
  currentDumpQty: number;
  compDumpQty: number;
  currentBillCount: number;
  compBillCount: number;
}

export function aggregateArticleMetrics(
  currentSales: SalesRecord[],
  compSales: SalesRecord[],
  currentDump: DumpRecord[],
  compDump: DumpRecord[]
): ArticleMetrics[] {
  const items = new Map<string, { desc: string; no: number; cls: string }>();
  currentSales.forEach((r) => {
    const key = `${r.item_no}`;
    if (!items.has(key)) items.set(key, { desc: r.item_desc, no: r.item_no, cls: r.class_name });
  });
  compSales.forEach((r) => {
    const key = `${r.item_no}`;
    if (!items.has(key)) items.set(key, { desc: r.item_desc, no: r.item_no, cls: r.class_name });
  });

  const sumSalesByItem = (records: SalesRecord[]) => {
    const valMap: Record<string, number> = {};
    const qtyMap: Record<string, number> = {};
    const billMap: Record<string, number> = {};
    records.forEach((r) => {
      const key = `${r.item_no}`;
      valMap[key] = (valMap[key] || 0) + r.sales_value;
      qtyMap[key] = (qtyMap[key] || 0) + r.sale_qty;
      billMap[key] = (billMap[key] || 0) + r.bill_count;
    });
    return { valMap, qtyMap, billMap };
  };

  const sumDumpByItem = (records: DumpRecord[]) => {
    const valMap: Record<string, number> = {};
    const qtyMap: Record<string, number> = {};
    records.forEach((r) => {
      const key = `${r.item}`;
      valMap[key] = (valMap[key] || 0) + r.dump_value;
      qtyMap[key] = (qtyMap[key] || 0) + r.dump_qty;
    });
    return { valMap, qtyMap };
  };

  const cs = sumSalesByItem(currentSales);
  const ps = sumSalesByItem(compSales);
  const cd = sumDumpByItem(currentDump);
  const pd = sumDumpByItem(compDump);

  const allKeys = new Set<string>();
  currentSales.forEach((r) => allKeys.add(`${r.item_no}`));
  compSales.forEach((r) => allKeys.add(`${r.item_no}`));
  currentDump.forEach((r) => allKeys.add(`${r.item}`));
  compDump.forEach((r) => allKeys.add(`${r.item}`));

  return [...allKeys].sort().map((key) => {
    const info = items.get(key) || { desc: key, no: Number(key), cls: '' };
    const csv = cs.valMap[key] || 0;
    const psv = ps.valMap[key] || 0;
    const cdv = cd.valMap[key] || 0;
    const pdv = pd.valMap[key] || 0;
    return {
      itemNo: info.no,
      itemDesc: info.desc,
      className: info.cls,
      currentSales: csv,
      compSales: psv,
      netSalesGrowth: csv - psv,
      growthPct: psv > 0 ? ((csv - psv) / psv) * 100 : 0,
      currentDump: cdv,
      currentDumpPct: csv > 0 ? (cdv / csv) * 100 : 0,
      compDump: pdv,
      compDumpPct: psv > 0 ? (pdv / psv) * 100 : 0,
      deltaDump: cdv - pdv,
      currentSaleQty: cs.qtyMap[key] || 0,
      compSaleQty: ps.qtyMap[key] || 0,
      currentDumpQty: cd.qtyMap[key] || 0,
      compDumpQty: pd.qtyMap[key] || 0,
      currentBillCount: cs.billMap[key] || 0,
      compBillCount: ps.billMap[key] || 0,
    };
  });
}

export function filterSalesByDrillLevel(
  sales: SalesRecord[],
  drill: DrillLevel
): SalesRecord[] {
  let filtered = sales;
  if (drill.store) {
    filtered = filtered.filter((r) => r.loc_name === drill.store);
  }
  if (drill.family) {
    filtered = filtered.filter((r) => r.family_name === drill.family);
  }
  if (drill.item) {
    filtered = filtered.filter((r) => r.item_desc === drill.item);
  }
  return filtered;
}

export function filterDumpByDrillLevel(
  dump: DumpRecord[],
  drill: DrillLevel
): DumpRecord[] {
  let filtered = dump;
  if (drill.store) {
    filtered = filtered.filter((r) => r.loc_name === drill.store);
  }
  if (drill.family) {
    filtered = filtered.filter((r) => r.family_name === drill.family);
  }
  if (drill.item) {
    filtered = filtered.filter((r) => r.item_desc === drill.item);
  }
  return filtered;
}

export function formatCurrency(value: number): string {
  const abs = Math.abs(Math.round(value));
  const formatted = abs.toLocaleString('en-IN');
  return value < 0 ? `-${formatted}` : formatted;
}

export function formatPct(value: number): string {
  return value.toFixed(1) + '%';
}

export function getMinDate(records: { b_date?: number; dump_date?: number }[]): string {
  const dates = records.map((r) => r.b_date ?? r.dump_date ?? Infinity);
  const min = Math.min(...dates);
  if (min === Infinity) return '';
  return excelSerialToDateString(min);
}

export function getMaxDate(records: { b_date?: number; dump_date?: number }[]): string {
  const dates = records.map((r) => r.b_date ?? r.dump_date ?? -Infinity);
  const max = Math.max(...dates);
  if (max === -Infinity) return '';
  return excelSerialToDateString(max);
}
