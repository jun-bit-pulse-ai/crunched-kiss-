# T6 — Large-workbook fixture generator (1.2M cells, planted errors, size flag)

**Agent model:** claude-haiku-4-5 (mechanical, fully specified). **Estimate:** 15 min. **Depends on:** T1. **Runs in parallel with:** T2-T5.
**Worktree:** `/Users/junseki/Documents/GitHub/ck-T6` on branch `task/T6`.

## Context

You are working in `/Users/junseki/Documents/GitHub/ck-T6` (an Excel add-in repo). Read `CLAUDE.md` first. `exceljs` (4.x)
is already installed; do NOT run `npm install`. Create a Node ES-module script that writes a large test workbook with the
exceljs streaming writer so memory stays flat. This workbook is what the demo uses to prove the add-in handles big
workbooks, so the sheet names and planted errors below are load-bearing: reproduce them exactly.

## You own (create exactly these)

- `scripts/make-fixture.mjs`
- `fixtures/.gitkeep` (empty file; `fixtures/*.xlsx` is gitignored)

Do NOT touch any other file.

## Specification

CLI: `node scripts/make-fixture.mjs [--rows N] [--out PATH]`. Defaults: `--rows 100000`, `--out fixtures/big-model.xlsx`.
Parse with a tiny loop over `process.argv` (no dependency). Print the resolved settings at start.

Use `new ExcelJS.stream.xlsx.WorkbookWriter({ filename: out, useStyles: false, useSharedStrings: false })`
(`import ExcelJS from 'exceljs'`). Add rows with `sheet.addRow([...]).commit()` and `sheet.commit()` after each sheet;
`await workbook.commit()` at the end. Formulas are cells of the form `{ formula: 'D2*E2' }` (no leading `=`).
Deterministic pseudo-random numbers via mulberry32 seeded with 42.

Sheets, in this order:

1. `Assumptions`: header `['Assumption', 'Value', 'Unit']`, then exactly 20 rows:
   `['Growth rate', 0.05, '%']`, `['Discount rate', 0.1, '%']`, `['Tax rate', 0.25, '%']`, `['Units start', 1000, 'units']`, `['Price', 49.99, 'USD']`,
   `['COGS ratio', 0.6, '%']`, `['Churn', 0.03, '%']`, `['Headcount', 42, 'people']`, `['Salary', 85000, 'USD']`, `['Rent', 12000, 'USD/month']`,
   `['Marketing', 0.12, '% of revenue']`, `['Capex', 250000, 'USD']`, `['Depreciation years', 5, 'years']`, `['Working capital days', 45, 'days']`,
   `['FX EURUSD', 1.08, 'rate']`, `['Inflation', 0.02, '%']`, `['Terminal growth', 0.02, '%']`, `['WACC', 0.09, '%']`, `['Shares', 1000000, 'count']`, `['Start year', 2024, 'year']`.
2. `Data`: header `['Date','Region','Product','Units','Price','Revenue','COGS','GrossMargin','Channel','Rep','Discount','Net']`, then `N` rows
   (row r = 2..N+1): `Date` = JS Date 2020-01-01 + (r-2) days; `Region` from `['NA','EMEA','APAC','LATAM']`; `Product` = `'P' + (1..50)`;
   `Units` integer 1..500; `Price` 5..500 with 2 decimals; `Revenue` = `{ formula: \`D${r}*E${r}\` }`; `COGS` = Units*Price*0.6 as a NUMBER (2 decimals; a
   deliberately hard-coded derived value); `GrossMargin` = `{ formula: \`F${r}-G${r}\` }`; `Channel` from `['Web','Retail','Partner']`; `Rep` = `'Rep' + (1..40)`;
   `Discount` 0..0.3 (3 decimals); `Net` = `{ formula: \`H${r}*(1-K${r})\` }`.
   Planted hard-codes: in rows `plantA = Math.max(3, Math.floor(N / 20)) + 1` and `plantB = Math.max(4, Math.floor(N * 0.777)) + 1` (sheet row numbers),
   put the NUMBER `12345` in `Revenue` instead of the formula. Print both row numbers at the end.
3. `Summary`: `['Metric', 'Value']`, then rows 2..7: `['Total revenue', { formula: \`SUM(Data!F2:F${N+1})\` }]`, `['Avg price', { formula: \`AVERAGE(Data!E2:E${N+1})\` }]`,
   `['Broken ratio', { formula: '1/0' }]`, `['Bad ref', { formula: '#REF!' }]`, `['Growth applied', { formula: 'B2*(1+Assumptions!B2)' }]`,
   `['Partial total (bug)', { formula: \`SUM(Data!F2:F${Math.floor(N / 20) + 1})\` }]`.
   If exceljs throws on the `#REF!` formula, write the STRING `'#REF!'` there and print a warning (the search demo still finds the text).
4. `Scenario_01` .. `Scenario_30` (zero-padded): header `['Year', 'Revenue', 'Cost', 'Profit']`, 10 rows: `Year` 2024..2033,
   `Revenue` = 1000 * (1 + 0.1*(scenarioIndex/30)) ^ (yearIndex) rounded to 2 decimals, `Cost` = Revenue * (0.5 + 0.01*scenarioIndex) rounded,
   `Profit` = `{ formula: \`B${r}-C${r}\` }`.

Finish with `console.log(\`wrote ${out} (${sizeMB} MB) in ${seconds}s; planted hard-codes at Data!F${plantA} and Data!F${plantB}\`)`.

## Verify, commit

```bash
node scripts/make-fixture.mjs                        # must finish in < 60 s
ls -la fixtures/big-model.xlsx                        # > 5 MB
node scripts/make-fixture.mjs --rows 10000 --out fixtures/small-model.xlsx    # < 10 s (fallback for slow Excel)
git add scripts/make-fixture.mjs fixtures/.gitkeep && git commit -m "T6: large workbook fixture generator"
```

## Acceptance

Both commands above succeed with the stated timings and sizes; opening `fixtures/big-model.xlsx` in Excel (human) shows 33 sheets
in the order Assumptions, Data, Summary, Scenario_01..Scenario_30; `Data` has 100,001 rows x 12 columns; `Summary!B4` shows `#DIV/0!`
and `Summary!B5` shows `#REF!`; `Summary!B2` is a number.

## When done, report

Final message: the two command outputs (timings, sizes, planted rows), and whether the `#REF!` formula was accepted or replaced by text.
