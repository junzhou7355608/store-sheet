import { mkdirSync, readdirSync, readFileSync } from 'fs';
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

/** 日期/月份列名，统一以字符串写入 Excel，与 JSON 一致 */
const DATE_COLUMNS = ['日期', '月份'];

/** 日期/月份列：将值规范为 yyyy-mm-dd 或 yyyy-mm 字符串写入 Excel，并设为文本格式 @，避免回车后被 Excel 转成 2026/1/28 等系统日期格式 */
function makeDateCell(
  val: unknown,
  colName: string,
): { t: 's'; v: string; z: string } | null {
  if (!DATE_COLUMNS.includes(colName) || val === '' || val === undefined)
    return null;
  const s = String(val).trim();
  if (!s) return null;
  const dMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (dMatch) {
    const [, y, m, d] = dMatch;
    return { t: 's' as const, v: `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`, z: '@' };
  }
  const mMatch = s.match(/^(\d{4})-(\d{1,2})$/);
  if (mMatch) {
    const [, y, m] = mMatch;
    return { t: 's' as const, v: `${y}-${m.padStart(2, '0')}`, z: '@' };
  }
  return null;
}

/** 将公式里的列名替换为单元格引用；表名!列名 替换为对应表的数据区域。先替换跨表引用，再替换同表列名，避免「日期」把「采购明细!日期」误替换成「采购明细!A2」。 */
function formulaWithRefs(
  formulaStr: string,
  currentColumns: string[],
  excelRow: number,
  allSheets: Map<string, SheetInfo>,
): string {
  let out = formulaStr;
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
  const nameToRef = new Map<string, string>();
  for (let i = 0; i < currentColumns.length; i++) {
    nameToRef.set(currentColumns[i], colToLetter(i) + excelRow);
  }
  const sortedNames = [...currentColumns].sort((a, b) => b.length - a.length);
  for (const name of sortedNames) {
    const ref = nameToRef.get(name)!;
    out = out.split(name).join(ref);
  }
  return out;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const dataDir = join(__dirname, '../data');
const outDir = join(__dirname, '../render-template');

/** 从 data 目录读取所有 .json 的基础名（排除 schema.json 和 copy 备份） */
function getJsonBaseNames(): string[] {
  return readdirSync(dataDir)
    .filter(
      (f) =>
        f.endsWith('.json') &&
        f !== 'schema.json' &&
        !f.replace(/\.json$/i, '').endsWith(' copy'),
    )
    .map((f) => f.replace(/\.json$/i, ''));
}

function buildXlsx(baseName: string): void {
  const jsonPath = join(dataDir, `${baseName}.json`);
  const outPath = join(outDir, `${baseName}-模板.xlsx`);

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
    const formulas = sheet.formulas;

    const rows = sheet.rows.map((row, r) =>
      header.map((col, c) => {
        const isFormulaCol = formulas && col in formulas;
        if (isFormulaCol) {
          const excelRow = r + 2;
          const excelFormula = formulaWithRefs(
            formulas![col],
            header,
            excelRow,
            sheetInfo,
          );
          const isPercent = col.endsWith('%');
          const cached = row[col];
          const v = typeof cached === 'number' ? cached : 0; // 缓存值供 Excel 显示
          return { t: 'n' as const, v, f: excelFormula, z: isPercent ? '0.00%' : '0.00' };
        }
        const raw = row[col] ?? '';
        const dateCell = makeDateCell(raw, col);
        if (dateCell) return dateCell;
        return raw;
      }),
    );

    const sheetData = [header, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
  }

  mkdirSync(outDir, { recursive: true });
  XLSX.writeFile(workbook, outPath);
  console.log('已生成:', outPath);
}

const baseNames = getJsonBaseNames();
for (const baseName of baseNames) {
  buildXlsx(baseName);
}
