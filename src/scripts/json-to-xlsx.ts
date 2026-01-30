import { mkdirSync, readFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
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

/** 0-based 列索引转 Excel 列字母 */
function colToLetter(i: number): string {
  let s = '';
  while (i >= 0) {
    s = String.fromCharCode(65 + (i % 26)) + s;
    i = Math.floor(i / 26) - 1;
  }
  return s;
}

/** 各表信息：表名 -> 列名数组、行数，用于生成跨表区域 */
type SheetInfo = { columns: string[]; rowCount: number };

/** 将公式里的列名替换为单元格引用；表名!列名 替换为对应表的数据区域 */
function formulaWithRefs(
  formulaStr: string,
  currentColumns: string[],
  excelRow: number,
  allSheets: Map<string, SheetInfo>,
): string {
  const nameToRef = new Map<string, string>();
  for (let i = 0; i < currentColumns.length; i++) {
    nameToRef.set(currentColumns[i], colToLetter(i) + excelRow);
  }
  const sortedNames = [...currentColumns].sort((a, b) => b.length - a.length);
  let out = formulaStr;
  for (const name of sortedNames) {
    const ref = nameToRef.get(name)!;
    out = out.split(name).join(ref);
  }
  const sheetNames = [...allSheets.keys()].sort((a, b) => b.length - a.length);
  for (const sheetName of sheetNames) {
    const info = allSheets.get(sheetName)!;
    const cols = [...info.columns].sort((a, b) => b.length - a.length);
    for (const colName of cols) {
      const token = sheetName + '!' + colName;
      const colIdx = info.columns.indexOf(colName);
      const letter = colToLetter(colIdx);
      const r2 = 2;
      const rEnd = 2 + info.rowCount - 1;
      const range = `${sheetName}!$${letter}$${r2}:$${letter}$${rEnd}`;
      out = out.split(token).join(range);
    }
  }
  return out;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const jsonPath = join(__dirname, '../data/店铺数据统计.json');
const outDir = join(__dirname, '../render-template');
const outPath = join(outDir, '店铺数据统计-模板.xlsx');

const data: DataFile = JSON.parse(readFileSync(jsonPath, 'utf-8'));
const workbook = XLSX.utils.book_new();

const sheetInfo = new Map<string, SheetInfo>();
for (const sheet of data.sheets) {
  sheetInfo.set(sheet.name, {
    columns: sheet.columns,
    rowCount: sheet.rows.length,
  });
}

for (const sheet of data.sheets) {
  const header = sheet.columns;
  const rows = sheet.rows.map((row) => header.map((col) => row[col] ?? ''));
  const sheetData = [header, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(sheetData);

  const formulas = sheet.formulas;
  if (formulas) {
    for (let r = 0; r < sheet.rows.length; r++) {
      const excelRow = r + 2;
      for (const [colName, formulaStr] of Object.entries(formulas)) {
        const colIdx = header.indexOf(colName);
        if (colIdx === -1) continue;
        const excelFormula = formulaWithRefs(
          formulaStr,
          header,
          excelRow,
          sheetInfo,
        );
        const ref = colToLetter(colIdx) + excelRow;
        const isPercent = colName.endsWith('%');
        ws[ref] = {
          t: 'n',
          f: '=' + excelFormula,
          z: isPercent ? '0.00%' : '0.00',
        };
      }
    }
  }

  XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
}

mkdirSync(outDir, { recursive: true });
XLSX.writeFile(workbook, outPath);
console.log('已生成:', outPath);
