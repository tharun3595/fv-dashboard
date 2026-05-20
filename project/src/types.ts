export interface SalesRecord {
  b_date: number;
  location_code: number;
  loc_name: string;
  item_no: number;
  item_desc: string;
  division_name: string;
  dept_name: string;
  family_name: string;
  class_name: string;
  sale_qty: number;
  sales_value: number;
  bill_count: number;
}

export interface DumpRecord {
  dump_date: number;
  location: number;
  loc_name: string;
  item: number;
  item_desc: string;
  division_name: string;
  dept_name: string;
  family_name: string;
  class_name: string;
  dump_qty: number;
  av_cost: number;
  dump_value: number;
}

export interface DateRange {
  start: string;
  end: string;
}

export interface StoreMetrics {
  name: string;
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

export interface DrillLevel {
  type: 'area' | 'store' | 'family' | 'item';
  store?: string;
  family?: string;
  item?: string;
}

export interface BreadcrumbItem {
  label: string;
  level: DrillLevel;
}
