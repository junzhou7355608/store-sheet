# 店铺数据统计

用于店铺（如烘焙店、零售店）的销售、采购与利润数据管理。数据以 JSON 格式存储，结构对应 Excel 工作簿（多工作表），便于导入导出和程序处理。

## 项目结构

```
shop/
├── README.md
├── package.json
├── tsconfig.json
├── .prettierrc
├── .vscode/
│   └── settings.json
└── src/
    ├── data/                      # 数据源（JSON）
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
- **值**：公式表达式。同表列名直接写列名；跨表引用写 `表名!列名`（如 `采购明细!金额`），导出为 Excel 时会替换为该表该列的数据区域
- 支持 `+ - * /`、括号及 Excel 函数（如 `SUMIF`）。该列在 `rows` 中可不填或仅作预览，导出时会按公式生成 Excel 公式。

### 工作表类型说明

| 工作表   | 说明                                                                                               |
| -------- | -------------------------------------------------------------------------------------------------- |
| 销售明细 | 按日期的销售记录：日期、商品、销量、售价、销售额、材料成本、毛利、毛利率等。                       |
| 采购明细 | 采购记录：日期、品名、数量、单位、单价、金额。                                                     |
| 月度利润 | 按月汇总：销售收入、销售成本、毛利、各项费用（租金、水电、人工、推广等）、总费用、净利润、利润率。 |

## 使用方式

### 推荐工作流

1. **生成模板**：`pnpm gen:template`，从 `data/*.json` 生成 `render-template/*-模板.xlsx`（含结构+公式）。
2. **复制模板**：将模板复制到 `render-data/`（若结构有变更则需重新复制）。
3. **编辑数据**：在 `render-data/*.xlsx` 中填入或修改真实数据。
4. **同步 JSON**：`pnpm gen:data`，将 `render-data/` 的 Excel 写回 `data/*.json`，保持数据同步。

### 其他

- **直接编辑 JSON**：也可直接编辑 `src/data/*.json`，再运行 `gen:template` 生成模板。
- **校验**：在 VS Code 等编辑器中打开 JSON 时，`$schema` 会提供结构提示和错误高亮。
- **扩展**：可在 `sheets` 中增加新工作表、列、行，符合 `schema.json` 即可。

## 示例片段

单个工作表的形状如下：

```json
{
  "name": "销售明细",
  "columns": [
    "日期",
    "商品名称",
    "销量",
    "售价",
    "销售额",
    "材料成本",
    "毛利",
    "毛利率%"
  ],
  "rows": [
    {
      "日期": "2025-01-05",
      "商品名称": "草莓蛋糕",
      "销量": 12,
      "售价": 40,
      "销售额": 480,
      "材料成本": 180,
      "毛利": 300,
      "毛利率%": "62.5%"
    }
  ]
}
```

## 脚本说明

### gen:template — JSON → 模板 Excel

- **命令**：`pnpm gen:template`（执行 `tsx src/scripts/json-to-xlsx.ts`）
- **输入**：`src/data/*.json`（排除 `schema.json` 和 `* copy.json`）
- **输出**：`src/render-template/*-模板.xlsx`
- **行为**：按 `sheets` 顺序生成多个工作表；`formulas` 中同表列名、`表名!列名` 会替换为 Excel 单元格/区域引用。

### gen:data — Excel → JSON 同步

- **命令**：`pnpm gen:data [文件名.xlsx]`（执行 `tsx src/scripts/xlsx-to-json.ts`）
- **输入**：`src/render-data/店铺数据统计.xlsx`（默认）或传入的文件名
- **输出**：`src/data/*.json`（覆盖同名数据文件）
- **行为**：读取 Excel 各工作表，提取列名、公式（转为列名形式）、行数据；含 `$schema` 引用。

### 排查脚本

- **inspect-xlsx**：`npx tsx src/scripts/inspect-xlsx.ts [xlsx路径]` — 打印 sheet 结构、列名、首行数据、公式。
- **debug-pipeline**：`npx tsx src/scripts/debug-pipeline.ts` — 分步检查模板与工作区 xlsx。

## 开发说明

- **环境**：Node.js（建议 18+）、pnpm
- **安装依赖**：`pnpm install`
- **脚本**：
  - `pnpm gen:template` — JSON → 模板 Excel
  - `pnpm gen:data [xlsx]` — Excel → JSON 同步
- **代码规范**：项目使用 Prettier（见 `.prettierrc`），提交前可格式化 JSON/TS 等。
