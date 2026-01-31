import { mkdirSync, writeFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import prettier from 'prettier';
import XLSX from 'xlsx';

interface SheetData {
  name: string;
  columns: string[];
  formulas?: Record<string, string>;
  rows: Record<string, string | number>[];
}

interface DataFile {
  sheets: SheetData[];
}

/** 从单元格取值：优先显示值，其次原始值，公式格取计算后的值 */
function getCellValue(cell: XLSX.CellObject | undefined): string | number {
  if (!cell) return '';
  if (cell.t === 'n' && typeof cell.v === 'number') return cell.v;
  if (cell.t === 's' && typeof cell.v === 'string') return cell.v;
  if (cell.t === 'd' && cell.v instanceof Date) return cell.v.toISOString().slice(0, 10);
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  if (typeof cell.v === 'string' || typeof cell.v === 'number') return cell.v;
  if (cell.w) return cell.w; // 格式化后的字符串
  return '';
}

/** Excel 列字母转 0-based 列索引（A=0, B=1, ..., Z=25, AA=26） */
function colLetterToIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

/** 判断 offset 处的引用是否属于跨表引用（表名!ref 或 表名!ref1:ref2） */
function isCrossSheetRef(formula: string, offset: number): boolean {
  for (let i = offset - 1; i >= 0; i--) {
    const c = formula[i];
    if (c === '!') return true; // 遇到 ! 说明是 表名!... 的一部分
    if ('=+-*/(,; '.includes(c)) return false; // 遇到公式操作符，说明是同表引用
  }
  return false;
}

/** 将公式中的同表单元格引用（如 C2、D2）转为列名（销量、售价），便于 json-to-xlsx 为每行生成正确公式 */
function formulaToColumnNames(
  formula: string,
  columns: string[],
  dataRow: number,
): string {
  const refRegex = /\$?[A-Z]+\$?\d+/g;
  return formula.replace(refRegex, (match, offset: number) => {
    if (isCrossSheetRef(formula, offset)) return match; // 跨表引用完整保留
    const rowPart = match.match(/\d+$/)?.[0] ?? '';
    const rowNum = parseInt(rowPart, 10);
    if (rowNum !== dataRow) return match;
    const colPart = match.replace(/\$?\d+$/, '').replace(/\$/g, '');
    const colIdx = colLetterToIndex(colPart);
    if (colIdx >= 0 && colIdx < columns.length) return columns[colIdx];
    return match;
  });
}

/** 数字格式为百分比的，或列名以 % 结尾的，格式为 "xx.x%" 字符串 */
function formatCellValue(
  cell: XLSX.CellObject | undefined,
  value: string | number,
  colName: string,
): string | number {
  if (value === '' || value === undefined) return '';
  const isPercent =
    (cell?.z && typeof cell.z === 'string' && cell.z.includes('%')) ||
    colName.endsWith('%');
  if (isPercent && typeof value === 'number') {
    // Excel 中百分比常存为小数 (如 0.332 -> 33.2%)；若绝对值 <= 2 视为小数，否则视为已为百分数
    const pct = Math.abs(value) <= 2 ? value * 100 : value;
    return `${pct.toFixed(1)}%`;
  }
  if (isPercent && typeof value === 'string' && !value.endsWith('%')) {
    const n = parseFloat(value);
    if (!Number.isNaN(n)) {
      const pct = Math.abs(n) <= 2 ? n * 100 : n;
      return `${pct.toFixed(1)}%`;
    }
  }
  return value;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/** 默认：render-data/店铺数据统计.xlsx -> data/店铺数据统计.json */
const defaultXlsxName = '店铺数据统计.xlsx';
const xlsxName = process.argv[2] || defaultXlsxName;
const baseName = xlsxName.replace(/\.xlsx?$/i, '');

const renderDataDir = join(__dirname, '../render-data');
const dataDir = join(__dirname, '../data');
const xlsxPath = join(renderDataDir, xlsxName);
const jsonPath = join(dataDir, `${baseName}.json`);

const workbook = XLSX.readFile(xlsxPath, { cellDates: true });
const sheets: SheetData[] = [];

for (const sheetName of workbook.SheetNames) {
  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1');

  const columns: string[] = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const ref = XLSX.utils.encode_cell({ r: range.s.r, c });
    const cell = ws[ref];
    const raw = getCellValue(cell);
    columns.push(String(raw ?? ''));
  }

  const formulas: Record<string, string> = {};
  const rows: Record<string, string | number>[] = [];

  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const row: Record<string, string | number> = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const ref = XLSX.utils.encode_cell({ r, c });
      const cell = ws[ref];
      const colName = columns[c - range.s.c];
      if (!colName) continue;

      if (cell?.f && !(colName in formulas)) {
        const excelRow = r + 1; // 表内 0-based 行 r → Excel 行号（首行为 1，数据首行通常为 2）
        formulas[colName] = formulaToColumnNames(cell.f, columns, excelRow);
      }
      const value = getCellValue(cell);
      row[colName] = formatCellValue(cell, value, colName);
    }
    rows.push(row);
  }

  sheets.push({
    name: sheetName,
    columns,
    ...(Object.keys(formulas).length > 0 && { formulas }),
    rows,
  });
}

const data: DataFile & { $schema?: string } = {
  $schema: '../schema/schema.json',
  sheets,
};

mkdirSync(dataDir, { recursive: true });
const jsonContent = JSON.stringify(data, null, 2);
const formatted = await prettier.format(jsonContent, {
  parser: 'json',
  filepath: jsonPath,
});
writeFileSync(jsonPath, formatted, 'utf-8');

console.log('已生成:', jsonPath);
