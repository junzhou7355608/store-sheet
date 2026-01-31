# 店铺数据统计

用于店铺（如烘焙店、零售店）的销售、采购与利润数据管理。数据以 JSON 格式存储，结构对应 Excel 工作簿（多工作表），便于导入导出和程序处理。

## 项目结构

```
store-sheet/
├── README.md
├── package.json
├── tsconfig.json
├── .prettierrc
├── .vscode/
│   └── settings.json
└── src/
    ├── data/                      # 数据源（JSON），为「单一事实来源」
    │   └── 店铺数据统计.json
    ├── backup/                    # gen:data 执行前对 data 的备份
    │   └── 店铺数据统计.json
    ├── schema/                    # 数据格式的 JSON Schema 定义
    │   └── schema.json
    ├── render-template/           # 模板（由 JSON 生成，含结构+公式）
    │   └── 店铺数据统计-模板.xlsx
    ├── render-data/               # 真实数据工作区（在此编辑 Excel）
    │   └── 店铺数据统计.xlsx
    └── scripts/
        ├── json-to-xlsx.ts        # JSON → 模板 Excel
        ├── xlsx-to-json.ts        # 工作区 Excel → JSON 同步
        ├── inspect-xlsx.ts        # 排查：检查 xlsx 文件结构
        └── debug-pipeline.ts      # 排查：完整流程诊断
```

## 数据格式说明

### 整体结构

数据文件对应一个「工作簿」概念，等同于一个 `.xlsx` 文件：

- **工作簿**：一个 JSON 对象，包含 `sheets` 数组。
- **工作表（Sheet）**：对应 Excel 底部的一个 Tab（如「销售明细」「采购明细」）。
- **列（columns）**：该表的列名，顺序即列顺序。
- **行（rows）**：数据行，每行是一个对象，键为列名，值为数字或字符串。

### Schema 校验

数据文件通过 `"$schema": "../schema/schema.json"` 引用 Schema，可用支持 JSON Schema 的编辑器或工具做格式校验与自动补全。

### 公式列（formulas）

每个工作表的可选字段 `formulas` 用于定义「由公式计算」的列：

- **键**：列名（须在 `columns` 中）
- **值**：公式表达式。
  - **同表**：直接写列名（如 `销量*售价`）。
  - **跨表**：写 `表名!列名`（如 `销售明细!销售额`、`采购明细!金额`），**不要**写死 Excel 区域（如 `$A$2:$A$4`）。
- 支持 `+ - * /`、括号及 Excel 函数（如 `SUMPRODUCT`、`TEXT`）。该列在 `rows` 中可不填或仅作预览，导出时会按公式生成 Excel 公式。

**公式约定（全部行）**：

- 在 JSON 中，跨表引用一律用 **`表名!列名`**，表示「该表该列的全部数据行」。
- **gen:template** 生成 Excel 时，会把 `表名!列名` 展开为固定区域（如 `销售明细!$A$2:$A$10000`），在 Excel 里加行（不超过 10000 行）后公式仍会覆盖新行。
- **gen:data** 从 Excel 读回时，会把公式里的 `表名!$A$2:$A$10000` 等区域**转回** `表名!列名`，保证 JSON 中始终是「全部行」约定，不会写死行号。

### 工作表类型说明

| 工作表   | 说明                                                                                               |
| -------- | -------------------------------------------------------------------------------------------------- |
| 销售明细 | 按日期的销售记录：日期、商品名称、销量、售价、销售额（公式：销量×售价）。                           |
| 采购明细 | 采购记录：日期、品名、数量、单位、单价、金额（公式：数量×单价）。                                   |
| 月度利润 | 按月汇总：销售收入、销售成本（由销售/采购明细按月份汇总）、毛利、各项费用、总费用、净利润、利润率。 |

## 使用方式

### 推荐工作流

1. **生成模板**：`pnpm gen:template` — 从 `data/*.json` 生成 `render-template/*-模板.xlsx`（含结构+公式）。
2. **复制到工作区**：`pnpm run copy:template` — 将模板复制到 `render-data/`（结构或公式有变更时需重新执行）。
3. **编辑数据**：在 `render-data/*.xlsx` 中填入或修改真实数据。
4. **同步回 JSON**：`pnpm gen:data` — 将 `render-data/` 的 Excel 写回 `data/*.json`。执行前会把当前 `data/*.json` 备份到 `backup/`。

**一键同步**：`pnpm sync` 依次执行：`gen:data` → `gen:template` → `copy:template` → `gen:data`，适合「从 Excel 拉回 → 更新模板 → 再写回 JSON」的完整回合。

### 其他

- **直接编辑 JSON**：也可直接编辑 `src/data/*.json`，再运行 `gen:template` 生成模板。
- **校验**：在 VS Code 等编辑器中打开 JSON 时，`$schema` 会提供结构提示和错误高亮。
- **扩展**：可在 `sheets` 中增加新工作表、列、行，符合 `schema.json` 即可。

## 示例片段

单个工作表的形状如下（含公式列）：

```json
{
  "name": "销售明细",
  "columns": ["日期", "商品名称", "销量", "售价", "销售额"],
  "formulas": {
    "销售额": "销量*售价"
  },
  "rows": [
    {
      "日期": "2026-01-29",
      "商品名称": "草莓蛋糕",
      "销量": 12,
      "售价": 40,
      "销售额": 480
    }
  ]
}
```

跨表公式示例（月度利润表引用销售/采购明细）：

```json
"formulas": {
  "销售收入": "SUMPRODUCT((TEXT(销售明细!日期,\"YYYY-MM\")=月份)*销售明细!销售额)",
  "销售成本": "SUMPRODUCT((TEXT(采购明细!日期,\"YYYY-MM\")=月份)*采购明细!金额)"
}
```

## 脚本说明

| 命令 | 说明 |
|------|------|
| `pnpm gen:template` | 从 `data/*.json` 生成 `render-template/*-模板.xlsx` |
| `pnpm gen:data [xlsx]` | 从 `render-data/*.xlsx` 写回 `data/*.json`（执行前备份到 `backup/`） |
| `pnpm run copy:template` | 将 `render-template/*-模板.xlsx` 复制到 `render-data/` |
| `pnpm sync` | 依次执行：gen:data → gen:template → copy:template → gen:data |

### gen:template — JSON → 模板 Excel

- **输入**：`src/data/*.json`（排除 `schema.json` 和 `* copy.json`）
- **输出**：`src/render-template/*-模板.xlsx`
- **行为**：按 `sheets` 顺序生成多个工作表；`formulas` 中同表列名替换为当前行单元格（如 `销量`→`C2`），`表名!列名` 替换为固定区域（如 `销售明细!$A$2:$A$10000`）；`日期`、`月份` 列设为文本格式，避免 Excel 自动转成系统日期。

### gen:data — Excel → JSON 同步

- **输入**：`src/render-data/店铺数据统计.xlsx`（默认）或 `pnpm gen:data 文件名.xlsx`
- **输出**：`src/data/*.json`（执行前若已存在则备份到 `src/backup/`）
- **行为**：读取 Excel 各工作表，提取列名、公式（同表引用转为列名，跨表区域如 `表名!$A$2:$A$10000` 转回 `表名!列名`）、行数据；写入的 JSON 含 `$schema` 引用。

### 排查脚本

- **inspect-xlsx**：`pnpm exec tsx src/scripts/inspect-xlsx.ts [xlsx路径]` — 打印 sheet 结构、列名、首行数据、公式。
- **debug-pipeline**：`pnpm exec tsx src/scripts/debug-pipeline.ts` — 分步检查模板与工作区 xlsx。

## 开发说明

- **环境**：Node.js（建议 18+）、pnpm
- **安装依赖**：`pnpm install`
- **代码规范**：项目使用 Prettier（见 `.prettierrc`），提交前可格式化 JSON/TS 等。
