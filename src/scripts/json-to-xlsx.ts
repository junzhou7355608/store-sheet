import { mkdirSync, readFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';

interface SheetData {
  name: string;
  columns: string[];
  rows: Record<string, string | number>[];
}

interface DataFile {
  sheets: SheetData[];
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const jsonPath = join(__dirname, '../data/店铺数据统计.json');
const outDir = join(__dirname, '../render');
const outPath = join(outDir, '店铺数据统计.xlsx');

const data: DataFile = JSON.parse(readFileSync(jsonPath, 'utf-8'));
const workbook = XLSX.utils.book_new();

for (const sheet of data.sheets) {
  const header = sheet.columns;
  const rows = sheet.rows.map((row) => header.map((col) => row[col] ?? ''));
  const sheetData = [header, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
}

mkdirSync(outDir, { recursive: true });
XLSX.writeFile(workbook, outPath);
console.log('已生成:', outPath);
