/**
 * 完整流程排查脚本
 * 步骤1: data/*.json → render-template/*.xlsx (json-to-xlsx)
 * 步骤2: 检查 render-template 和 render-data
 * 步骤3: render-data/*.xlsx → data/*.json (xlsx-to-json)
 *
 * 用法: npx tsx src/scripts/debug-pipeline.ts
 */
import { mkdirSync, readdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';

const __dirname = dirname(fileURLToPath(import.meta.url));
const dataDir = join(__dirname, '../data');
const templateDir = join(__dirname, '../render-template');
const renderDataDir = join(__dirname, '../render-data');

function inspectXlsx(path: string, label: string) {
  console.log('\n' + '='.repeat(70));
  console.log(label, path);
  console.log('='.repeat(70));
  try {
    const wb = XLSX.readFile(path, { cellDates: true });
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1');
      const cols: string[] = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const ref = XLSX.utils.encode_cell({ r: range.s.r, c });
        cols.push(String(ws[ref]?.v ?? ws[ref]?.w ?? ''));
      }
      console.log(`\n[${name}] 行${range.s.r}-${range.e.r} 列${range.s.c}-${range.e.c}`);
      console.log('  列名:', cols.join(' | '));
      if (range.e.r >= range.s.r + 1) {
        const r = range.s.r + 1;
        console.log(`  首行数据(row${r + 1}):`);
        for (let c = range.s.c; c <= range.e.c; c++) {
          const ref = XLSX.utils.encode_cell({ r, c });
          const cell = ws[ref];
          const v = cell?.v ?? cell?.w ?? '(空)';
          const f = cell?.f ? ` [f=${cell.f}]` : '';
          console.log(`    ${ref} ${cols[c - range.s.c]}: ${JSON.stringify(v)}${f}`);
        }
      }
    }
  } catch (e) {
    console.log('  无法读取:', (e as Error).message);
  }
}

// 只处理 店铺数据统计，排除 copy
const baseName = '店铺数据统计';
const jsonPath = join(dataDir, `${baseName}.json`);
const templatePath = join(templateDir, `${baseName}-模板.xlsx`);
const renderDataPath = join(renderDataDir, `${baseName}.xlsx`);

console.log('\n========== 步骤 0: 当前 data JSON 状态 ==========');
const inputJson = JSON.parse(readFileSync(jsonPath, 'utf-8'));
console.log('sheets:', inputJson.sheets?.map((s: { name: string }) => s.name));
const xiaoshou = inputJson.sheets?.find((s: { name: string }) => s.name === '销售明细');
if (xiaoshou) {
  console.log('销售明细 columns:', xiaoshou.columns);
  console.log('销售明细 rows[0] keys:', Object.keys(xiaoshou.rows?.[0] ?? {}));
  console.log('销售明细 rows[0] 有材料成本?', '材料成本' in (xiaoshou.rows?.[0] ?? {}));
}

console.log('\n========== 步骤 1: json-to-xlsx 生成 template ==========');
// 直接调用 json-to-xlsx 逻辑
const data: { sheets: unknown[] } = JSON.parse(readFileSync(jsonPath, 'utf-8'));
// 简化版 build - 实际用 npx tsx 跑 json-to-xlsx
console.log('执行: npx tsx src/scripts/json-to-xlsx.ts');
console.log('(请手动运行上述命令，或继续看步骤2)');

console.log('\n========== 步骤 2: 检查 render-template 和 render-data ==========');
inspectXlsx(templatePath, 'render-template:');
inspectXlsx(renderDataPath, 'render-data:');
