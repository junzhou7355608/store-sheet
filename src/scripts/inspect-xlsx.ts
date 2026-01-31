/**
 * 排查脚本：检查 xlsx 文件结构，打印每个 sheet 的 range、列名、首行数据、公式
 * 用法: npx tsx src/scripts/inspect-xlsx.ts <文件路径>
 */
import { readFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';

const __dirname = dirname(fileURLToPath(import.meta.url));
const defaultPath = join(__dirname, '../render-data/店铺数据统计.xlsx');
const xlsxPath = process.argv[2] || defaultPath;

console.log('检查文件:', xlsxPath);
const workbook = XLSX.readFile(xlsxPath, { cellDates: true });

for (const sheetName of workbook.SheetNames) {
  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1');

  console.log('\n' + '='.repeat(60));
  console.log(`Sheet: ${sheetName}`);
  console.log(`Range: row ${range.s.r}-${range.e.r}, col ${range.s.c}-${range.e.c}`);
  console.log(`(Excel: 行 ${range.s.r + 1}-${range.e.r + 1}, 列 ${range.s.c + 1}-${range.e.c + 1})`);

  const columns: string[] = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const ref = XLSX.utils.encode_cell({ r: range.s.r, c });
    const cell = ws[ref];
    const v = cell?.v ?? cell?.w ?? '';
    columns.push(String(v));
  }
  console.log('列名:', columns);

  console.log('\n首行数据 (row', range.s.r + 1, '= 表头, row', range.s.r + 2, '= 首条数据):');
  for (let r = range.s.r + 1; r <= Math.min(range.s.r + 2, range.e.r); r++) {
    console.log(`  [row ${r + 1}]:`);
    for (let c = range.s.c; c <= range.e.c; c++) {
      const ref = XLSX.utils.encode_cell({ r, c });
      const cell = ws[ref];
      const colName = columns[c - range.s.c];
      const val = cell?.v ?? cell?.w ?? '(空)';
      const formula = cell?.f ? ` [公式: ${cell.f}]` : '';
      console.log(`    ${ref} ${colName}: ${JSON.stringify(val)}${formula}`);
    }
  }

  const formulaCols = new Set<string>();
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.f && columns[c - range.s.c]) {
        formulaCols.add(columns[c - range.s.c]);
      }
    }
  }
  if (formulaCols.size > 0) {
    console.log('\n含公式的列:', [...formulaCols]);
  }
}
