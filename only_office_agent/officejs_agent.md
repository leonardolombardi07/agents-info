# Excel AI Agent — Blueprint (Office.js Add-in)

This is the blueprint for an AI chat agent that runs as a **Microsoft Excel add-in** (task pane) and works on the open workbook through the **Office.js Excel JavaScript API**. It covers the agent's behaviour, its complete system prompt, its complete tool definitions, and the conventions needed to rebuild it.

The design is adapted from the AI agent in the ONLYOFFICE spreadsheet editor (AI plugin v3.3.1, © Ascensio System SIA, AGPL-3.0). Its prompts and tool texts are reused, with everything editor-specific rewritten for Excel and Office.js.

---

## 1. What the agent is

- A **chat panel in the Excel task pane**. The user asks in natural language, and the agent reads and edits the open workbook through **24 tools**.
- It is a single agent with **native tool calling**. A few tools make their own small LLM call internally (§7).
- Replies mention cells as **clickable citation links**. Clicking one selects the range in Excel (§6).

## 2. Architecture

```
┌──────────────────────── Excel task pane (add-in web page) ────────────────────────┐
│ Chat UI (streaming, tool-call cards with UI labels, Stop button, citation links)  │
│ Agent loop  ──► LLM (native tool calling) — via your backend, never expose keys    │
│ Tool implementations: plain async JS functions, each one runs                      │
│   Excel.run(async (context) => { ... await context.sync(); ... })                  │
└────────────────────────────────────────────────────────────────────────────────────┘
```

- **Tools execute in the task pane** because that's where Office.js is available. If the agent loop runs on a server, it must send each tool call to the task pane and wait for the result.
- **Gate tools by Excel capability.** Before registering a tool, check `Office.context.requirements.isSetSupported("ExcelApi", "<version>")` (§8 lists versions), and leave out tools whose APIs are missing.
- **The human-readable "UI label"** of each tool is shown on its tool-call card. It is not sent to the model.

## 3. Agent loop rules

1. **Each request** sends the system prompt (§4), the full conversation history (including earlier tool calls and their results), and the tool list (§10).
2. If the model returns a tool call, **execute it, append the result as a tool message, and call the model again**. Stop when the model answers without a tool call.
3. **One tool call at a time.** If the model emits several, run them sequentially, in order.
4. **Errors are returned to the model, not raised.** A failing tool returns an error string, and the model is expected to fix its approach and retry. The prompt allows up to 5 retries for `executeOfficeJs`.
5. **Atomicity.** Office.js has no transactions, rollback, or undo grouping.
   - Specialized tools validate **all** inputs (addresses, sheet names, column names, shapes) **before writing anything**, so they either fail without changes or succeed. For example, `setCellData` promises "on any validation error nothing is written".
   - `executeOfficeJs` can leave partial changes when it fails. Its error message says so (§5.3).
6. **Read tools get a citation reminder** appended to their result (§5.2).
7. The workbook tools run **without asking the user for approval**.
8. A **Stop** button aborts generation and the running tool. The model then receives `{"cancelled":true,"message":"Tool call cancelled by the user."}`.
9. **Cap the loop** at about 25 tool calls per user message.
10. The model only sees each tool's `name`, `description` and `parameters`.

## 4. System prompt (complete)

~~~~~text
You are an AI assistant running as an add-in inside Microsoft Excel.
You have access to tools that can read and modify the open workbook.

IMPORTANT RULES FOR WORKBOOK OPERATIONS:

1. When the user asks about the workbook content or wants to perform an action on it, you MUST use the available tools.
2. First, check if there is a specialized tool that matches the request (e.g., readSheetContext, setCellData, addChart, setSort, etc.).
3. If no specialized tool fits, use the "executeOfficeJs" tool. It can execute any JavaScript code using the Office.js Excel API.
4. The "executeOfficeJs" tool can also READ data from the workbook. Load the properties, sync, and explicitly return the value you want to retrieve. Example: const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1"); r.load("values"); await context.sync(); return r.values[0][0];
5. If an "executeOfficeJs" call returns an error, analyze the error message, fix the code, and retry. You may retry up to 5 times for a single request before giving up and informing the user.
6. If the user’s question is NOT about the workbook content (e.g., general knowledge, coding help, conversation), respond normally without calling any tools.

Always prefer specialized tools over executeOfficeJs when available. Use executeOfficeJs as a universal fallback for any workbook operation.

CITATIONS: Whenever your reply names a cell, range, column, row, or sheet, write it as a clickable markdown citation link, NOT plain text - in every reply, including lists, summaries, and audits, and even when answering from earlier context with no new tool call. Angle brackets are REQUIRED (sheet names can contain spaces):
- Cell / range: [B5](<#citation:cell:Sheet1!B5>), [A1:D10](<#citation:cell:Sheet1!A1:D10>)
- Column / row: [C:C](<#citation:cell:Sheet1!C:C>), [5:5](<#citation:cell:Sheet1!5:5>)
- Whole sheet: [Sales Data](<#citation:cell:Sales Data>); active sheet omit the name: [B5](<#citation:cell:B5>)
Write findings and list items this way - "Assumptions block: [A3:C9](<#citation:cell:Analytics!A3:C9>)" - and NEVER as bare text like "Assumptions block: A3:C9" or "Analytics!A3:C9".
RULE: cite only addresses from tool results or ones you wrote.
LOCALE: Write formulas and values in en-US syntax only, whatever the user's language or Excel's display language: comma between arguments, dot decimals, English function names, TRUE/FALSE (=WENN(B2>0,5;1;0) -> =IF(B2>0.5,1,0)). Formulas are entered through range.formulas, which always takes en-US syntax.
~~~~~

## 5. Tool result conventions

### 5.1 Strings returned to the model

| Case | Tool message content |
|---|---|
| Success, nothing returned | `Function executed successfully` |
| Success with a value | `Function executed successfully\n` + the value (strings as-is, objects as JSON) |
| Tool failed | `{"error":"<message>"}` |
| Unknown tool name | `{"error":"Tool \"<name>\" is not available."}` |
| Cancelled by the user | `{"cancelled":true,"message":"Tool call cancelled by the user."}` |

Tool error messages are written **for the model**: they say what was wrong and what to do instead, e.g. `Parameter "range" must be a string like "A1:D100". Got: …`.

### 5.2 Citation reminder

Append this to the result of `readSheetContext`, `getRangeData`, `getCellDetails`, `searchData` and `getSheetObjects`:

~~~~~text
REMINDER: in THIS reply write every cell, range, or sheet you mention as a citation link exactly like [A3:C9](<#citation:cell:Analytics!A3:C9>) (whole sheet, spaces need no quotes: [Sales Data](<#citation:cell:Sales Data>)) - not as plain text or any other link format, and not in a follow-up.
~~~~~

### 5.3 `executeOfficeJs`: how the code runs and how errors are enriched

**Running the code.** Wrap the model's code as the body of an async function and run it in a batch:

```js
const fn = new AsyncFunction("context", "Excel", code);   // AsyncFunction = (async () => {}).constructor
// a SyntaxError here → return it to the model as the error
const result = await Excel.run(async (context) => {
  const r = await fn(context, Excel);
  await context.sync();
  return r;
});
```

**Security:** this executes model-written code with full access to the workbook and the task pane. Restrict it to your organization's needs. Options include blocking network APIs inside the code (`fetch`, `XMLHttpRequest`), running it in a sandboxed iframe or worker that only exposes the Excel proxy, or asking for user confirmation for destructive operations.

**Error enrichment.** The error returned is `<error.code>: <error.message>`, followed by an optional **fix hint**, a **partial-changes note**, and a **docs hint**.

Fix hints, matched on the error code or message (first match wins):

| Error code / message contains | Appended text |
|---|---|
| `PropertyNotLoaded` | `Fix: load before reading — obj.load("prop1,prop2"); await context.sync(); then read obj.prop1. Every property you read (values, address, name, isNullObject, items...) must be loaded first.` |
| `InvalidArgument` with "doesn't match the size" / "dimensions" | `Fix: values/formulas/numberFormat must be a 2-D array with exactly the range's shape: "A1" → [["x"]], "A1:C1" → [["a","b","c"]], "A1:A3" → [["a"],["b"],["c"]]. Build matrices with Array.from({length: rows}, ...).` |
| `ItemNotFound` | `Fix: the sheet/table/name/chart/item does not exist. Use exact names from readSheetContext or getSheetObjects, or getItemOrNullObject(name) + await context.sync() + check isNullObject.` |
| `ItemAlreadyExists` | `Fix: the name is already used. Create-or-get with getItemOrNullObject, or choose a unique name (add a numeric suffix).` |
| `InvalidReference` | `Fix: invalid or out-of-sheet address. Use sheet.getRange(colLetter + row) with letters from readSheetContext (nextCol, cols[].letter); check for NaN/undefined in built addresses; getCell/getRangeByIndexes are 0-based (A1 = 0,0).` |
| `AccessDenied` / "protected" | `Fix: the sheet or range is protected. Do not unprotect it on your own — tell the user what you need to change.` |
| `InsertDeleteConflict` | `Fix: the insert/delete collides with a table, merged cells, a pivot or an array formula. Delete/insert whole rows or columns outside those objects, or modify the table through its API.` |
| `RequestPayloadSizeLimitExceeded` / `ResponsePayloadSizeLimitExceeded` | `Fix: too much data in one batch. Read or write in chunks of ~5,000 rows with await context.sync() between chunks; return only aggregates or matching rows.` |
| `ApiNotFound` / "not supported" / `NotImplemented` | `Fix: this API is not available in the user's Excel version. Use an older equivalent (see getSpreadsheetApiReference) or a different approach.` |
| `shiftCol is not defined` | `Fix: shiftCol is a helper you must define in your code. Copy this verbatim before using it: function shiftCol(col,n){var v=0;for(var i=0;i<col.length;i++)v=v*26+(col.charCodeAt(i)-64);v+=n;var r='';while(v>0){r=String.fromCharCode((v-1)%26+65)+r;v=Math.floor((v-1)/26);}return r;}` |
| `ExcelScript` / `main(` / VBA keywords (`Dim `, `Sub `, `.Value`) | `Fix: this is Office.js, not Office Scripts or VBA. Write only the body of async (context, Excel) => {...}, e.g. const ws = context.workbook.worksheets.getActiveWorksheet(); ws.getRange("A1").values = [["x"]];` |
| `is not a function` / `Cannot read properties of undefined` | `Hint: that member does not exist in Office.js (names are camelCase properties, e.g. range.format.font.bold, not SetBold()). Call getSpreadsheetApiReference({className:'Excel.Range'}) (or the relevant class) to see the real members.` |

Always appended after the hint:

~~~~~text
Note: Office.js is not transactional — commands that ran before the error may already be applied to the workbook. Check the current state (readSheetContext / getRangeData) before retrying, and make setup steps idempotent (getItemOrNullObject, overwrite instead of append).

Docs: call getSpreadsheetApiReference({topics:['gotchas']}) plus the relevant area topic (for example io, defined_names, validation, comments, pivots, cond_format), or use className:'Excel.<Class>' to inspect real members.
~~~~~

## 6. Citations

- Link format the model writes: `[visible text](<#citation:cell:<reference>>)`. The angle brackets are required because sheet names can contain spaces.
- Reference forms: `Sheet1!B5`, `Analytics!A3:C9`, `Sheet1!C:C`, `Sheet1!5:5`, `Sales Data` (whole sheet), and `B5` (active sheet).
- The chat UI intercepts clicks on links starting with `#citation:cell:` and runs:

```js
await Excel.run(async (context) => {
  const ref = decodeURIComponent(target);            // e.g. "Sales Data!B5" or "Sales Data"
  const bang = ref.lastIndexOf("!");
  const sheets = context.workbook.worksheets;
  const ws = bang === -1 ? sheets.getItemOrNullObject(ref.replace(/^'|'$/g, ""))
                         : sheets.getItemOrNullObject(ref.slice(0, bang).replace(/^'|'$/g, ""));
  await context.sync();
  if (bang === -1 && !ws.isNullObject) { ws.activate(); return; }       // whole-sheet citation
  const sheet = ws.isNullObject ? sheets.getActiveWorksheet() : ws;     // "B5" → active sheet
  sheet.activate();
  sheet.getRange(bang === -1 ? ref : ref.slice(bang + 1)).select();
  await context.sync();
});
```

## 7. Tool overview

| Tool | UI label | Type | Returns |
|---|---|---|---|
| `addCellValueCondition` | Highlight Cells by Condition | write | Confirmation, e.g. "Conditional formatting applied to E2:E81: greater than 100." |
| `addChart` | Create Chart | write | Summary of the chart created (type, range, position) |
| `addColorScale` | Apply Color Scale | write | "Applied <type> scale to <range>." |
| `addDataBars` | Add Data Bars | write | "Added data bars to <range>." |
| `addIconSet` | Add Icon Indicators | write | "Added icon set \"<name>\" to <range>." |
| `addImage` | Insert Image | write · LLM (image) | "Image inserted at <cell> (W×H mm)." |
| `addTop10Condition` | Highlight Top or Bottom Values | write | "Highlighted top/bottom N[%] in <range>." |
| `describeImage` | Describe Image | read · LLM (vision) | `{ description, altText }` |
| `explainFormula` | Explain Formula | read · LLM | Explanation text |
| `fillMissingData` | Fill Missing Data | write · LLM | Summary of the cells filled |
| `fixFormula` | Fix Formula Errors | write · LLM | List of fixed cells (old → new formula) |
| `getSpreadsheetApiReference` | Office.js API Reference | read (docs) | Topic texts (§9) and/or the member list of an Excel.* class |
| `getCellDetails` | Get Cell Details | read | `{hasMore, ranges:[{range, cells, styles}]}` (see its description) |
| `getRangeData` | Get Range Data | read | `{headers, rows, returnedRange, nextRange}` (see its description) |
| `getSheetObjects` | Get Sheet Objects | read | Per-sheet charts, pivots, tables, CF rules, shapes, names, comments, hyperlinks, freeze panes |
| `insertPivotTable` | Create Pivot Table | write | Name and location of the new pivot sheet |
| `readSheetContext` | Read Spreadsheet Context | read | Workbook/sheet map: regions, header row, row count, column letters/types, samples (see its description) |
| `searchData` | Search Data | read | `[{sheet, cells:[{addr, value, formula?}]}]` |
| `setAltText` | Set Alternative Text | write | Confirmation |
| `setAutoFilter` | Apply Data Filter | write | Confirmation sentence |
| `setCellData` | Set Cell Data | write | `{writtenCount, formulaChecks, formulaChecksTruncated, duplicates, ignoredFields, formulasRejected}` |
| `setMultiSort` | Sort by Multiple Columns | write | "Sorted <range> by <keys>." |
| `setSort` | Sort Data | write | "Sorted <range> by <key> <order>." |
| `executeOfficeJs` | Run Office.js Code | read/write | Whatever the code returns (or nothing) |

**Tools with an internal LLM call:** the tool sends its own one-off request to the model, outside the chat history, and uses the answer.

**`explainFormula`**, a chat request. The answer text is returned as the tool result:
~~~~~text
Explain the following Excel formula in detail:

Formula: <formula>
Cell: <address>

IMPORTANT RULES:
1. Provide a clear, detailed explanation of what the formula does.
2. Break down each part of the formula if it's complex.
3. Explain the functions used and their parameters.
4. Describe the expected result or output.
5. Use simple, understandable language.
6. If there are nested functions, explain the order of operations.
7. Mention any potential issues or common mistakes.
8. Keep the explanation concise but comprehensive.
9. Be brief and avoid unnecessary verbose explanations.
10. Get straight to the point without filler text.
11. Focus only on essential information.
12. Keep response length under 1024 characters (recommended), maximum 32767 characters.
13. Prioritize the most important information if length constraint requires cuts.

Please provide a detailed but concise explanation of this formula.
~~~~~

**`fixFormula`**, a chat request. It collects every formula in the range (or in all used ranges of the workbook), one per line as `  <Sheet>!<A1>: <formula>`, and appends ` → <#ERROR> [FIX THIS]` to the cells whose value is an error. It parses the JSON answer and writes the fixed formulas back:
~~~~~text
You are an Excel formula expert. Some formulas in the list below produce errors (marked with [FIX THIS]).
The other formulas are provided as context to help you understand the data structure and patterns.

Fix guidelines:
- #DIV/0!: wrap division in IF(denominator<>0, ..., 0) — do NOT use IFERROR to hide the error
- #REF!: correct the broken cell reference based on surrounding formula patterns
- #NAME?: fix the function name typo (e.g. VLOOKPU → VLOOKUP)
- #VALUE!: fix argument types (e.g. wrap text in VALUE() or use correct data type)
- #N/A from VLOOKUP/MATCH: wrap in IFERROR(..., "") only if the missing value is expected
- #NULL!: fix the range operator (missing colon or incorrect intersection)
- #NUM!: fix the numeric argument (e.g. negative SQRT input, date overflow)

Return ONLY a JSON array of objects for cells you are fixing:
[{"sheet": "Sheet1", "address": "A5", "formula": "=fixed_formula"}, ...]
- Only include cells marked [FIX THIS] where the formula actually changes
- If you cannot determine a safe fix, omit that cell
- No markdown, no code blocks — ONLY the raw JSON array

Formulas:
<formula lines>
~~~~~

**`fillMissingData`**, a chat request. It sends the range as CSV, parses the JSON array, and writes the values into the empty cells. `FORWARD_FILL` means copy the previous non-empty value in the column:
~~~~~text
You are a data analyst.
Input is CSV (comma-separated, ','). Empty cells represent missing values to be filled.
Rules:
1. NUMERIC columns: Fill missing values with MEDIAN of non-empty values.
2. CATEGORICAL columns: Fill missing values with MOST FREQUENT value.
3. TIME_SERIES columns: Fill missing values with FORWARD FILL (previous non-empty value).

Output format: JSON array with exact row/column coordinates (1-based indexing):
[
  {"row": 2, "column": 1, "new_value": 25.5},
  {"row": 3, "column": 2, "new_value": "Category A"},
  {"row": 4, "column": 3, "new_value": "FORWARD_FILL"}
]
- Use "FORWARD_FILL" as new_value for time series columns
- Row and column numbers are 1-based (first row = 1, first column = 1)
- Only include cells that need to be filled
- The answer MUST be valid JSON array
- No extra text, spaces, or newlines outside JSON

CSV:
<csv>
~~~~~

**`describeImage`**, a vision request. It gets the object's image as a base64 PNG (`shape.getAsImage(Excel.PictureFormat.png)` for shapes and pictures, `chart.getImage()` for charts), sends it with the user's question (`prompt`) to a vision-capable model, and returns `{ description, altText }`. `altText` is `{title, description}` from `shape.altTextTitle` / `altTextDescription`, or `null` for a chart.

**`addImage`**, an image-generation request. The prompt is `<style or "realistic"> style, <description>, image size <W>x<H> pixels` (mm converted at 96 dpi; default 100×100 mm), plus `, wide panoramic format` if the aspect ratio is > 1.8, `, tall vertical format` if < 0.6, or `, square format` if between 0.9 and 1.1. The generated image (as base64 PNG/JPEG) is inserted with `sheet.shapes.addImage(base64)` at the selected cell (position from the cell's `left`/`top`, size in points: mm × 72 / 25.4).

## 8. Implementation map: Office.js per tool

The model never sees this section. It is guidance for implementing each tool.

| Tool | Main Office.js APIs | Min. ExcelApi |
|---|---|---|
| `readSheetContext` | `worksheets` (`load items/name`), `getActiveWorksheet`, `getSelectedRange`, `sheet.getRangeByIndexes(0,0,300,150)` → `load values, formulas, valueTypes`, then detect regions in JS (blocks separated by empty rows/cols); `workbook.name` for fileName | 1.1 (1.7 for `getActiveCell`) |
| `getRangeData` | `sheet.getRange(range).getUsedRangeOrNullObject(true)` to trim, `load values` (empty `""` → `null`), page by `maxRows` → `nextRange` | 1.4 |
| `getCellDetails` | `load values, formulas, numberFormat`; styles via `range.getCellProperties({format:{font:{bold,italic,size,name,color},fill:{color},wrapText:true,textOrientation:true}})` → group identical styles into ranges | 1.9 |
| `searchData` | Load `values`/`formulas` of the used ranges and match in JS (wildcards → RegExp, regex option, matchCase, wholeCell, formula text); page with `offset` | 1.1 |
| `getSheetObjects` | `sheet.charts` (`id,name,chartType,left,top,width,height`, `series`), `sheet.pivotTables`, `sheet.tables` (`name,style,showTotals`, `getRange().address`, `columns`), `range.conditionalFormats` on the used range, `sheet.shapes` (`id,name,type,altTextTitle,altTextDescription,left,top,width,height`), `workbook.names`/`sheet.names`, `workbook.comments` (`getLocation()`), hyperlinks from `getCellProperties({hyperlink:true})`, `sheet.freezePanes.getLocationOrNullObject()`, `workbook.getActiveShapeOrNullObject()` / `getActiveChartOrNullObject()` → `selected: true`, `workbook.linkedWorkbooks` for externalLinks (null if unsupported) | 1.9 (1.10 comments, 1.19 active shape) |
| `setCellData` | Validate everything first; then `range.numberFormat`, `range.values` / `range.formulas`, CSS → `range.format.*` (font, fill, borders, alignment, wrapText), `copyTo` → `dest.copyFrom(seed, Excel.RangeCopyType.all)` (relative refs shift); `note` → `workbook.comments.add(address, text)` (or `workbook.notes` 1.18); `format.autofitColumns()/autofitRows()`; `validateFormulas` → `sync`, `load values` of sampled cells; a cell whose formula throws goes to `formulasRejected` | 1.9 (1.10 comments) |
| `executeOfficeJs` | Runs model code inside `Excel.run` (§5.3) | 1.1 |
| `getSpreadsheetApiReference` | `topics` → return the texts in §9; `className` → list `Object.getOwnPropertyNames(Excel.<Class>.prototype)` (skip `_`-prefixed, constructor, load/set/toJSON/track/untrack) | — |
| `addChart` | Build the source range from `categoryColumn` + `seriesColumns` (union via helper columns or `chart.series.add(name)` + `series.setValues(range)` / `setXAxisValues(range)`); `sheet.charts.add(Excel.ChartType.<mapped>, range, Excel.ChartSeriesBy.auto)`; `chart.title.text`; place below the data with `chart.setPosition(topLeftCell)`; `outputSheet` → add on that sheet | 1.1 (1.7 series ops) |
| `insertPivotTable` | New sheet; `sheet.pivotTables.add(name, sourceRange, destCell)`; `rowHierarchies.add`, `columnHierarchies.add`, `dataHierarchies.add(...).summarizeBy = Excel.AggregationFunction.<agg>` | 1.8 |
| `setSort` / `setMultiSort` | `range.sort.apply([{key, ascending}...], false, hasHeaders)` with `key` = 0-based column offset; map `xlYes/xlNo` → hasHeaders | 1.2 |
| `setAutoFilter` | `sheet.autoFilter.apply(range, columnIndex, criteria)` with criteria `{filterOn: values|topItems|topPercent|bottomItems|bottomPercent|cellColor|fontColor|dynamic|custom, criterion1, criterion2, operator: and|or, values}`; resolve `fieldName` by fuzzy header match | 1.9 |
| `addCellValueCondition` | `range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue)`; `rule = {formula1, formula2, operator}` mapping `xlGreater→GreaterThan, xlLess→LessThan, xlEqual→EqualTo, xlNotEqual→NotEqualTo, xlGreaterEqual→GreaterThanOrEqual, xlLessEqual→LessThanOrEqual, xlBetween→Between, xlNotBetween→NotBetween`; `format.fill.color`, `format.font.color` | 1.6 |
| `addColorScale` | `ConditionalFormatType.colorScale`; `colorScale.criteria = {minimum, midpoint?, maximum}` | 1.6 |
| `addDataBars` | `ConditionalFormatType.dataBar`; `dataBar.positiveFormat.fillColor`, `showDataBarOnly = !showValue`, `barDirection` (`Context`/`LeftToRight`/`RightToLeft`) | 1.6 |
| `addIconSet` | `ConditionalFormatType.iconSet`; `iconSet.style = Excel.IconSet.<mapped>`, `showIconOnly`, `reverseIconOrder` | 1.6 |
| `addTop10Condition` | `ConditionalFormatType.topBottom`; `topBottom.rule = {rank, type: TopItems|BottomItems|TopPercent|BottomPercent}`; `format.fill.color` | 1.6 |
| `explainFormula` / `fixFormula` / `fillMissingData` | Read with `load formulas/values/valueTypes`; the LLM call as in §7; write back with `range.formulas` / `range.values` | 1.1 |
| `describeImage` | `shapes.getItem(id).getAsImage(Excel.PictureFormat.png)` or `chart.getImage()`; vision model call; `sheet.activate()` (+ `chart.activate()`) | 1.9 |
| `setAltText` | `sheet.shapes.getItem(objectId)`: `altTextTitle`, `altTextDescription` (search all sheets for the id) | 1.9 |
| `addImage` | Image model → base64; `sheet.shapes.addImage(base64)`; `left/top/width/height` in points | 1.9 |

## 9. Reference topics for `getSpreadsheetApiReference`

The texts the tool returns for each topic. Called with no arguments, it returns the topic list from its `topics` parameter description. Expand these as you find new failure patterns.

~~~~~text
=== entry ===
Code runs as the body of async (context, Excel) => {...} inside Excel.run. Root: context.workbook.
- wb.worksheets.getActiveWorksheet() → Worksheet
- wb.worksheets.getItem("Name") → Worksheet (THROWS ItemNotFound if missing)
- wb.worksheets.getItemOrNullObject("Name") → after sync check .isNullObject
- wb.worksheets.add("Name") → Worksheet (THROWS if the name exists); ws.activate(); ws.delete(); ws.name = "New"
- wb.worksheets.load("items/name") → list sheets; wb.getSelectedRange(); wb.getActiveCell()
Create-or-get:
  let ws = context.workbook.worksheets.getItemOrNullObject("Report"); await context.sync();
  if (ws.isNullObject) ws = context.workbook.worksheets.add("Report");

=== ranges ===
- ws.getRange("A1:D10"), ws.getRange("C:C"), ws.getRange("5:5")
- ws.getCell(row, col), ws.getRangeByIndexes(row, col, rowCount, colCount) — 0-based from A1
- r.getCell(r, c) (relative, 0-based), r.getOffsetRange(dRows, dCols), r.getResizedRange(dRows, dCols)
- r.getEntireRow(), r.getEntireColumn(), r.getSurroundingRegion() (= Ctrl+A region), r.getLastRow(), r.getLastColumn()
- ws.getUsedRange(true) (values only) / ws.getUsedRangeOrNullObject(true)
- r.address after load is "Sheet1!A1:D10"; r.rowCount, r.columnCount, r.rowIndex, r.columnIndex (0-based)
- Column letter helper (copy verbatim): function shiftCol(col,n){var v=0;for(var i=0;i<col.length;i++)v=v*26+(col.charCodeAt(i)-64);v+=n;var r='';while(v>0){r=String.fromCharCode((v-1)%26+65)+r;v=Math.floor((v-1)/26);}return r;}

=== io ===
- READ: r.load("values, formulas, text, numberFormat, valueTypes"); await context.sync(); then r.values[i][j]
- WRITE: r.values = [[...]]; r.formulas = [["=SUM(B2:B10)"]]; r.formulasR1C1 = [["=RC[-1]*2"]]
- Shapes must match exactly (rows × cols); empty cells read as "" ; errors read as "#DIV/0!" etc; valueTypes: Empty, String, Double, Boolean, Error
- Fill a formula down: r.formulas = Array.from({length: n}, (_, i) => ["=C" + (first + i) + "*2"])  or  r.formulasR1C1 = Array.from({length: n}, () => ["=RC[-1]*2"])
- Bulk transform: load once → map in JS → write once. Never sync inside loops.
- Big data: chunk reads/writes (~5,000 rows) with a sync per chunk; return aggregates, not whole sheets.
- Always use formulas (en-US), never formulasLocal.

=== rows_cols ===
- Insert: ws.getRange("5:7").insert(Excel.InsertShiftDirection.down); ws.getRange("C:C").insert(Excel.InsertShiftDirection.right)
- Delete: ws.getRange("5:5").delete(Excel.DeleteShiftDirection.up); ws.getRange("C:D").delete(Excel.DeleteShiftDirection.left)
- Delete several rows: loop from the BOTTOM up (indexes shift), all in one batch, one sync.
- Hide: r.rowHidden = true; r.columnHidden = true
- Size (points): r.format.columnWidth = 120; r.format.rowHeight = 20; r.format.autofitColumns(); r.format.autofitRows()

=== format ===
- Font: r.format.font.bold/italic/strikethrough = true; .underline = "Single"|"None"; .size = 11; .name = "Calibri"; .color = "#1F4E79"
- Fill: r.format.fill.color = "#FFF2CC"; r.format.fill.clear()
- Number format (2-D!): r.numberFormat = Array.from({length: rows}, () => Array(cols).fill("#,##0.00"))
- Alignment: r.format.horizontalAlignment = "Left"|"Center"|"Right"|"Justify"; r.format.verticalAlignment = "Top"|"Center"|"Bottom"; r.format.wrapText = true; r.format.indentLevel = 1
- Borders: const b = r.format.borders.getItem(Excel.BorderIndex.edgeBottom); b.style = "Continuous"; b.weight = "Thin"; b.color = "#000000"
  Indexes: edgeTop, edgeBottom, edgeLeft, edgeRight, insideVertical, insideHorizontal, diagonalDown, diagonalUp. Weights: Hairline, Thin, Medium, Thick.
- Merge: r.merge(false) (true = merge each row separately); r.unmerge()
- Styles: r.style = "Good" / "Bad" / "Neutral" / "Heading 1" (built-in style names)

=== defined_names ===
- wb.names.add("Prices", ws.getRange("B2:B100"), "comment?") ; wb.names.add("Rate", "=0.07")
- wb.names.getItemOrNullObject("Prices") → .getRange() for range names; .formula; .delete()
- Sheet scope: ws.names.add(...); list: wb.names.load("items/name,formula,type")

=== validation ===
- List: r.dataValidation.rule = { list: { inCellDropDown: true, source: "Open,Closed,Pending" } }  (or source: "=$H$2:$H$10" / "=NamedRange")
- Numbers: { wholeNumber: { formula1: 0, formula2: 100, operator: Excel.DataValidationOperator.between } }; also decimal, date, time, textLength, custom: { formula: "=ISNUMBER(A2)" }
- Alerts: r.dataValidation.errorAlert = { showAlert: true, style: Excel.DataValidationAlertStyle.stop, title: "Invalid", message: "..." }
- Prompt: r.dataValidation.prompt = { showPrompt: true, title: "...", message: "..." }
- Remove: r.dataValidation.clear()  (clear before setting a different rule type)

=== comments ===
- Threaded comments: wb.comments.add("Sheet1!B5", "text"); wb.comments.getItemByCell("Sheet1!B5") (THROWS if none); c.content; c.resolved = true; c.replies.add("text"); c.delete()
- List: wb.comments.load("items/content,authorName"); location: c.getLocation() → Range (load address)
- Notes (legacy, where supported): wb.notes.add("Sheet1!B5", "text")

=== freeze ===
- ws.freezePanes.freezeRows(1); ws.freezePanes.freezeColumns(2); ws.freezePanes.freezeAt(ws.getRange("B2")); ws.freezePanes.unfreeze()
- Read: const loc = ws.freezePanes.getLocationOrNullObject(); loc.load("address"); await context.sync(); (isNullObject → nothing frozen)

=== find_replace ===
- const f = r.findOrNullObject("text", { completeMatch: false, matchCase: false }); f.load("address"); await context.sync(); if (!f.isNullObject) ...
- ws.findAllOrNullObject("text", {completeMatch, matchCase}) → RangeAreas of all matches
- const n = r.replaceAll("Open", "Pending", { completeMatch: true, matchCase: false }); await context.sync(); return n.value;
- Remove duplicate rows: const res = r.removeDuplicates([0, 1], true); res.load("removed,uniqueRemaining"); await context.sync();  (0-based columns, includesHeader)

=== charts ===
- const ch = ws.charts.add(Excel.ChartType.columnClustered, ws.getRange("A1:C13"), Excel.ChartSeriesBy.columns)
- ch.title.text = "Sales"; ch.legend.position = "Bottom"; ch.setPosition("E2", "L18"); ch.name = "SalesChart"
- Series: const s = ch.series.add("Profit"); s.setValues(ws.getRange("D2:D13")); s.setXAxisValues(ws.getRange("A2:A13")); s.chartType = Excel.ChartType.line
- Axes: ch.axes.valueAxis.title.text = "USD"; ch.axes.categoryAxis.title.text = "Month"
- Types: columnClustered, columnStacked, columnStacked100, barClustered, line, lineMarkers, pie, doughnut, area, xyscatter, ... (Excel.ChartType)
- Prefer the addChart tool for normal charts.

=== pivots ===
- const pt = ws.pivotTables.add("SalesPivot", srcRangeOrTable, destSheet.getRange("A3"))
- pt.rowHierarchies.add(pt.hierarchies.getItem("Region")); pt.columnHierarchies.add(pt.hierarchies.getItem("Year"))
- const dh = pt.dataHierarchies.add(pt.hierarchies.getItem("Sales")); dh.summarizeBy = Excel.AggregationFunction.sum (count, average, max, min, ...)
- Filters (page fields): pt.filterHierarchies.add(pt.hierarchies.getItem("Country"))
- pt.refresh(); pt.delete(); layout: pt.layout.layoutType = "Tabular"
- Inspect existing pivots with getSheetObjects first.

=== tables ===
- const t = ws.tables.add("A1:D20", true); t.name = "Sales"; t.style = "TableStyleMedium2"
- t.getHeaderRowRange(), t.getDataBodyRange(), t.columns.getItem("Amount").getDataBodyRange()
- Add rows: t.rows.add(null, [["a", 1], ["b", 2]]); add column: t.columns.add(null, [["Header"], ["v1"], ...])
- Totals: t.showTotals = true; t.columns.getItem("Amount").getTotalRowRange().formulas = [["=SUBTOTAL(109,[Amount])"]]
- A table cannot overlap another table, a pivot, or merged cells. t.convertToRange() removes the table object.

=== cond_format ===
- Value rule: const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.cellValue); cf.cellValue.rule = { formula1: "=100", operator: "GreaterThan" }; cf.cellValue.format.fill.color = "#FFC7CE"; cf.cellValue.format.font.color = "#9C0006"
- Formula rule: add(Excel.ConditionalFormatType.custom); cf.custom.rule.formula = "=$C2>$D2"; cf.custom.format.fill.color = "#C6EFCE"
- Duplicates / above average: add(Excel.ConditionalFormatType.presetCriteria); cf.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.duplicateValues }; cf.preset.format.fill.color = ...
- Text contains: add(Excel.ConditionalFormatType.containsText); cf.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "late" }
- Other types: colorScale, dataBar, iconSet, topBottom. Inspect: r.conditionalFormats.load("items/type"); remove all: r.conditionalFormats.clearAll()
- Prefer the specialized CF tools for primary CF tasks.

=== drawings ===
- Image: const img = ws.shapes.addImage(base64WithoutPrefix); img.left = 20; img.top = 40; img.width = 200; img.lockAspectRatio = true (points)
- Shape: const s = ws.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle); s.fill.setSolidColor("#4472C4"); s.textFrame.textRange.text = "Note"
- Text box: ws.shapes.addTextBox("text"); line: ws.shapes.addLine(x1, y1, x2, y2)
- Find: ws.shapes.getItem(idOrName); list: ws.shapes.load("items/id,name,type,left,top,width,height")
- Alt text: s.altTextTitle, s.altTextDescription. Position a shape next to a cell: load the cell's left/top (Range.left/top in points).

=== protection ===
- ws.protection.protect({ allowFormatCells: true, allowAutoFilter: true }, "password?"); ws.protection.unprotect("password?")
- Unlock input cells before protecting: r.format.protection.locked = false
- Read: ws.protection.load("protected"). Writing to a protected sheet fails with AccessDenied — tell the user; never unprotect without their permission.

=== patterns ===
1) Formula column next to a table (letters/rows from readSheetContext):
   const ws = context.workbook.worksheets.getActiveWorksheet(); const h = 3, n = 120, col = "H";
   ws.getRange(col + h).values = [["Margin"]];
   ws.getRange(`${col}${h+1}:${col}${h+n}`).formulas = Array.from({length: n}, (_, i) => [`=F${h+1+i}/E${h+1+i}`]);
   ws.getRange(`${col}${h+1}:${col}${h+n}`).numberFormat = Array.from({length: n}, () => ["0.0%"]);
   ws.getRange(`${col}${h}:${col}${h+n}`).select();
2) Bulk transform: const r = ws.getRange("B2:B501"); r.load("values"); await context.sync(); r.values = r.values.map(([v]) => [String(v).trim().toUpperCase()]);
3) Bottom-up delete of rows matching a test: load values once, collect row numbers, sort descending, ws.getRange(`${row}:${row}`).delete(Excel.DeleteShiftDirection.up) for each, one sync.
4) Lookup column: formulas like =XLOOKUP(A2,Prices!A:A,Prices!B:B,"") (or INDEX/MATCH for older Excel) filled down as in 1).
5) Audit: load formulas + valueTypes of the used range; return only cells with valueTypes "Error" or inconsistent formulas (cap ~100).

=== gotchas ===
✗ const v = r.values (no load)                   → ✓ r.load("values"); await context.sync(); r.values
✗ r.values = "x" / [ "a","b" ]                  → ✓ r.values = [["x"]] / [["a","b"]] (exact 2-D shape)
✗ await context.sync() inside a for loop        → ✓ queue everything, sync once (or per chunk)
✗ worksheets.getItem(maybeMissing)              → ✓ getItemOrNullObject + sync + isNullObject
✗ r.find(text) when it may not exist            → ✓ findOrNullObject
✗ return r (proxy object)                       → ✓ return r.values / plain data
✗ r.formulasLocal / ";" separators              → ✓ r.formulas with en-US syntax and ","
✗ Sub/Dim/Range("A1").Value, ExcelScript, main() → ✓ Office.js body: context.workbook...
✗ SetBold(true) / GetValue()                    → ✓ r.format.font.bold = true / load values
✗ colors as RGB objects or names                → ✓ "#RRGGBB" strings
✗ column width in characters                   → ✓ points (r.format.columnWidth = 90)
✗ getUsedRange().rowCount as data bounds        → ✓ headerRow / dataRowCount from readSheetContext
✗ letter math String.fromCharCode(65+n)         → ✓ nextCol / shiftCol / getOffsetRange
✗ deleting rows top-down                        → ✓ bottom-up
✗ writing into a table's "formula" column       → ✓ add a new column
✗ assuming a failed call changed nothing        → ✓ re-check state; Office.js is not transactional

=== hyperlinks ===
- r.hyperlink = { address: "https://example.com", textToDisplay: "Open site", screenTip: "Opens the site" }
- Internal: r.hyperlink = { documentReference: "'Sales Data'!A1", textToDisplay: "Go to Sales" }
- Read: r.load("hyperlink"); clear: r.clear(Excel.ClearApplyTo.hyperlinks)

=== page_setup ===
- ws.pageLayout.orientation = Excel.PageOrientation.landscape; ws.pageLayout.paperSize = Excel.PaperType.a4
- ws.pageLayout.setPrintArea("A1:H50"); ws.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 }
- Margins (points): ws.pageLayout.topMargin = 36; headers/footers: ws.pageLayout.headersFooters.defaultForAllPages.centerHeader = "Report"
- Repeat header rows: ws.pageLayout.setPrintTitleRows("$1:$1")
~~~~~

## 10. Complete tool definitions

These are exactly what the model receives (name, description, parameters), in order. Enum values such as `xlGreater`, `xlFilterValues`, `xlYes` and the chart type names are the agent's own vocabulary; each tool maps them to Office.js enums (§8).

~~~~~json
[
  {
    "name": "addCellValueCondition",
    "description": "Creates conditional formatting rules based on cell values using comparison operators (greater than, less than, equal to, between, etc.). This is the primary tool for all value-based conditional formatting. If the user specifies an operator and value — use them directly. If the user asks for conditional formatting without a concrete threshold (e.g. 'highlight high values', 'mark important cells') — first call getRangeData to read the range, analyze the values, choose a meaningful threshold yourself, then call this tool. Do NOT use for above/below average rules — use addAboveAverage. Do NOT use for color scale/heat map — use addColorScale",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to apply condition (e.g., 'E2:E81'). If the user explicitly says 'current selection' or 'selected cells' — omit this parameter. Otherwise, ALWAYS determine the range explicitly: if the user named a column or table — call readSheetContext to resolve it; if no range is mentioned at all — call readSheetContext to find the relevant data range on the sheet. Do NOT silently fall back to selection when the user hasn't asked for it."
        },
        "operator": {
          "type": "string",
          "description": "Comparison operator - 'xlGreater', 'xlLess', 'xlEqual', 'xlNotEqual', 'xlGreaterEqual', 'xlLessEqual', 'xlBetween', 'xlNotBetween'.",
          "enum": [
            "xlGreater",
            "xlLess",
            "xlEqual",
            "xlNotEqual",
            "xlGreaterEqual",
            "xlLessEqual",
            "xlBetween",
            "xlNotBetween"
          ]
        },
        "value1": {
          "type": [
            "string",
            "number"
          ],
          "description": "First comparison value or formula (e.g., 100 or '=AVERAGE(A:A)'). Required for all operators."
        },
        "value2": {
          "type": [
            "string",
            "number"
          ],
          "description": "Second boundary value — required only for 'xlBetween' and 'xlNotBetween' operators (e.g., value1=5, value2=15 means 'between 5 and 15')."
        },
        "fillColor": {
          "type": "object",
          "description": "Background fill color as RGB object. Optional — defaults to yellow if omitted. Must be {r, g, b} integers 0–255. Not a hex string.",
          "properties": {
            "r": {
              "type": "number"
            },
            "g": {
              "type": "number"
            },
            "b": {
              "type": "number"
            }
          }
        },
        "fontColor": {
          "type": "object",
          "description": "Font (text) color as RGB object. Optional — use when the user asks to change text color. Must be {r, g, b} integers 0–255. Not a hex string.",
          "properties": {
            "r": {
              "type": "number"
            },
            "g": {
              "type": "number"
            },
            "b": {
              "type": "number"
            }
          }
        }
      },
      "required": [
        "operator",
        "value1"
      ]
    }
  },
  {
    "name": "addChart",
    "description": "Primary tool for inserting charts into a spreadsheet. Use for ANY chart request where the data has column headers — bar, line, pie, stacked, 3D, etc. Builds categories (X-axis) and series by exact column name. Column names in any language are supported (ASCII, Cyrillic, CJK, etc.).\nGood for: visualizing columns chosen by the user or by the assistant, summary charts with specific category and value fields.\nData without a text/date header column (e.g. numeric header row, values laid out in rows): omit categoryColumn and seriesColumns — the whole range is plotted the way Excel's Insert Chart does it (series by rows when the range is not taller than wide, otherwise by columns).\nCall readSheetContext first to get regionAddress (→ range) and column names with their types (text/date → categoryColumn; number → seriesColumns).\nDo NOT pass guessed column names — only exact header names from readSheetContext. Pass categoryColumn and seriesColumns together or omit both.\nOn error: re-read headers via readSheetContext and retry with exact names. Do NOT fall back to executeOfficeJs unless addChart is explicitly unsuitable.",
    "parameters": {
      "type": "object",
      "properties": {
        "sheet": {
          "type": "string",
          "description": "Sheet name containing the source data. If omitted, uses the active sheet."
        },
        "outputSheet": {
          "type": "string",
          "description": "Sheet to place the chart on, when it differs from sheet (e.g. the user asks for the chart on another existing sheet). The chart still references the source data on sheet. If omitted, the chart is placed on sheet below the data."
        },
        "range": {
          "type": "string",
          "description": "Cell range of the source table, e.g. 'A1:D24'. Use regionAddress from readSheetContext. If omitted, uses the current region around the selection."
        },
        "categoryColumn": {
          "type": "string",
          "description": "Column name for the X-axis categories (text or date columnType). Omit together with seriesColumns to plot the whole range as Excel does."
        },
        "seriesColumns": {
          "type": "array",
          "description": "Column names for chart series (number columnType). One or more. Omit together with categoryColumn to plot the whole range as Excel does.",
          "items": {
            "type": "string"
          }
        },
        "chartType": {
          "type": "string",
          "description": "Chart type. Default: ColumnClustered.",
          "enum": [
            "ColumnClustered",
            "ColumnStacked",
            "ColumnStacked100",
            "3DColumnClustered",
            "3DColumnStacked",
            "3DColumnStacked100",
            "3DColumn",
            "BarClustered",
            "BarStacked",
            "BarStacked100",
            "3DBarClustered",
            "3DBarStacked",
            "3DBarStacked100",
            "Line",
            "LineStacked",
            "LineStacked100",
            "3DLine",
            "Pie",
            "3DPie",
            "Doughnut",
            "XYScatter",
            "StockHLC",
            "Area",
            "AreaStacked",
            "AreaStacked100"
          ]
        },
        "title": {
          "type": "string",
          "description": "Chart title text."
        }
      },
      "required": []
    }
  },
  {
    "name": "addColorScale",
    "description": "Applies color scale conditional formatting to visualize data with gradient colors. Creates a heat map effect where values are represented by colors ranging from one color (minimum values) to another color (maximum values). Use 2-color scale for simple comparisons or 3-color scale for more detailed data visualization. Do NOT use for value-based rules (use addCellValueCondition) or above/below average (use addAboveAverage).",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to apply color scale (e.g., 'A1:D10'). REQUIRED unless the user explicitly says 'current selection' or 'selected cells'. If the user identified the column or table by name, call readSheetContext first to get the address, then pass it here."
        },
        "colorScaleType": {
          "type": "number",
          "description": "Color scale type: 2 = two-color gradient (min→max, simpler); 3 = three-color gradient (min→midpoint→max, better for showing both extremes). Default: 3.",
          "enum": [
            2,
            3
          ],
          "default": 3
        }
      },
      "required": []
    }
  },
  {
    "name": "addDataBars",
    "description": "Adds data bar conditional formatting to display values as horizontal bars within cells. The length of each bar represents the value relative to other values in the range. Useful for creating in-cell bar charts and comparing values at a glance without additional charts. Do NOT use for color gradient visualization — use addColorScale. Do NOT use for value-based highlight rules — use addCellValueCondition.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to apply data bars (e.g., 'A1:D10'). If the user explicitly says 'current selection' or 'selected cells' — omit this parameter. Otherwise, ALWAYS determine the range explicitly: if the user named a column or table — call readSheetContext to resolve it; if no range is mentioned at all — call readSheetContext to find the relevant data range on the sheet. Do NOT silently fall back to selection when the user hasn't asked for it."
        },
        "barColor": {
          "type": "object",
          "description": "Color of the data bars as RGB object. Optional — defaults to blue {r:0, g:112, b:192} if omitted. Must be {r, g, b} integers 0–255. Not a hex string.",
          "properties": {
            "r": {
              "type": "number"
            },
            "g": {
              "type": "number"
            },
            "b": {
              "type": "number"
            }
          }
        },
        "showValue": {
          "type": "boolean",
          "description": "Whether to show the cell values along with bars (default: true).",
          "default": true
        },
        "direction": {
          "type": "string",
          "description": "Direction of bars - 'xlLTR' (left to right), 'xlRTL' (right to left), 'xlContext' (context-based, default: 'xlLTR').",
          "enum": [
            "xlContext",
            "xlLTR",
            "xlRTL"
          ],
          "default": "xlLTR"
        }
      },
      "required": []
    }
  },
  {
    "name": "addIconSet",
    "description": "Applies icon set conditional formatting to display icons (arrows, traffic lights, symbols) based on value ranges. Icons provide visual indicators for data trends and performance levels. Each icon represents a different value range, making it easy to identify patterns and outliers in your data. Do NOT use for color gradient visualization — use addColorScale. Do NOT use for bar-based visualization — use addDataBars. Do NOT use for value-based highlight rules — use addCellValueCondition.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to apply icon set (e.g., 'A1:D10'). If the user explicitly says 'current selection' or 'selected cells' — omit this parameter. Otherwise, ALWAYS determine the range explicitly: if the user named a column or table — call readSheetContext to resolve it; if no range is mentioned at all — call readSheetContext to find the relevant data range on the sheet. Do NOT silently fall back to selection when the user hasn't asked for it."
        },
        "iconSetType": {
          "type": "string",
          "description": "Type of icon set. Default: 'xl3Arrows'. Choose based on user intent: arrows (xl3Arrows, xl4Arrows, xl5Arrows, xl3ArrowsGray, xl4ArrowsGray, xl5ArrowsGray) — for trends/direction; traffic lights (xl3TrafficLights1, xl3TrafficLights2, xl4TrafficLights) — for status/performance; symbols (xl3Symbols, xl3Symbols2) — for yes/no, pass/fail; shapes/signs (xl3Flags, xl3Signs, xl3Stars, xl3Triangles) — for ratings/priority; filled shapes (xl5Quarters, xl5Boxes, xl4RedToBlack, xl4CRV, xl5CRV) — for percentage/progress.",
          "enum": [
            "xl3Arrows",
            "xl3ArrowsGray",
            "xl3Flags",
            "xl3TrafficLights1",
            "xl3TrafficLights2",
            "xl3Signs",
            "xl3Symbols",
            "xl3Symbols2",
            "xl4Arrows",
            "xl4ArrowsGray",
            "xl4RedToBlack",
            "xl4CRV",
            "xl4TrafficLights",
            "xl5Arrows",
            "xl5ArrowsGray",
            "xl5CRV",
            "xl5Quarters",
            "xl3Stars",
            "xl3Triangles",
            "xl5Boxes"
          ],
          "default": "xl3Arrows"
        },
        "showIconOnly": {
          "type": "boolean",
          "description": "Whether to show only icons without cell values (default: false).",
          "default": false
        },
        "reverseOrder": {
          "type": "boolean",
          "description": "Whether to reverse the icon order (default: false).",
          "default": false
        }
      },
      "required": []
    }
  },
  {
    "name": "addImage",
    "description": "Generates an image using AI from a text description and inserts it into the spreadsheet. Pass a description of what the image should show; optionally specify width, height (in mm), and style (realistic, cartoon, abstract, etc.).",
    "parameters": {
      "type": "object",
      "properties": {
        "description": {
          "type": "string",
          "description": "Text description of the image to generate"
        },
        "width": {
          "type": "number",
          "description": "Image width in mm (default: 100)",
          "default": 100
        },
        "height": {
          "type": "number",
          "description": "Image height in mm (default: 100)",
          "default": 100
        },
        "style": {
          "type": "string",
          "description": "Image style (realistic, cartoon, abstract, etc.)"
        }
      },
      "required": [
        "description"
      ]
    }
  },
  {
    "name": "addTop10Condition",
    "description": "Highlights the top or bottom ranked values in a range. You can choose to highlight by item count (e.g., top 10 values) or by percentage (e.g., top 20% of values). Perfect for identifying highest performers or values that need attention in your dataset. Do NOT use for above/below average rules — use addAboveAverage. Do NOT use for value-based rules with explicit thresholds — use addCellValueCondition.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to apply condition (e.g., 'A1:D10'). If the user explicitly says 'current selection' or 'selected cells' — omit this parameter. Otherwise, ALWAYS determine the range explicitly: if the user named a column or table — call readSheetContext to resolve it; if no range is mentioned at all — call readSheetContext to find the relevant data range on the sheet. Do NOT silently fall back to selection when the user hasn't asked for it."
        },
        "rank": {
          "type": "number",
          "description": "Number of top/bottom items to highlight (default: 10).",
          "default": 10
        },
        "isBottom": {
          "type": "boolean",
          "description": "True for bottom values, false for top values (default: false).",
          "default": false
        },
        "isPercent": {
          "type": "boolean",
          "description": "True to treat rank as percentage (e.g., rank=20 + isPercent=true = top 20%), false to treat as item count (e.g., rank=10 = top 10 items). Default: false.",
          "default": false
        },
        "fillColor": {
          "type": "object",
          "description": "Background fill color as RGB object. Optional — defaults to light green if omitted. Must be {r, g, b} integers 0–255. Not a hex string.",
          "properties": {
            "r": {
              "type": "number"
            },
            "g": {
              "type": "number"
            },
            "b": {
              "type": "number"
            }
          }
        }
      },
      "required": []
    }
  },
  {
    "name": "describeImage",
    "description": "Analyzes an object on a sheet as a picture (a photo, a chart, a diagram) and answers a question about it (describe what is shown, recognize text on it, explain a chart/diagram, etc.).\nUse this tool whenever the user asks about a picture, an image or a photo, or wants something made from one (a title, a caption, a description). If the user did not attach the picture to the message, it is a picture in the workbook, usually the selected one, even when the user does not say so: never answer that you cannot see it. First call getSheetObjects: each shape (picture, drawing) and chart has its id, and the one selected on the active sheet has selected: true; then pass the id of the picture as objectId to this tool. Do not ask the user to select the image. If nothing is selected and the active sheet has several pictures, ask the user which one they mean, naming each by its name; if it has one, use it. The tool activates the sheet of the object (and the chart itself, when it is a chart), so the user sees which one is analyzed.\nReturns {description, altText}: altText is the alternative text the object already has, {title, description} with null for a missing part; it is null for a chart, whose alternative text cannot be saved.\nTo put a title or a caption next to the object, use executeOfficeJs: get a shape with sheet.shapes.getItem(objectId) (a chart: load sheet.charts items with \"id,name\" and match the id); its left/top/width/height are in points, so place the caption in a cell near it (e.g. the first empty cell below the data or to the right of the object).\nWhen the user asked to describe the image, after the description offer to save it: as the alternative text of the object (a short title and a description; not for a chart) or as a caption in a cell next to it. If the object already has alternative text, show it and ask whether to replace it. Save only after the user agrees: the alternative text with setAltText, the caption with setCellData or executeOfficeJs.",
    "parameters": {
      "type": "object",
      "properties": {
        "objectId": {
          "type": "string",
          "description": "The id of the shape or chart to analyze, from getSheetObjects."
        },
        "prompt": {
          "type": "string",
          "description": "The question about the image to send to the vision model. Put the user's actual request here (e.g. \"Describe what is shown in the image\")."
        }
      },
      "required": [
        "objectId",
        "prompt"
      ]
    }
  },
  {
    "name": "explainFormula",
    "description": "Analyzes and explains Excel formulas in natural language. Uses AI to provide detailed explanations of formula logic, function parameters, nested operations, and expected results. Particularly useful for understanding complex formulas with multiple nested functions or unfamiliar Excel functions.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range containing formula to explain (e.g., 'A1'). If omitted, uses active/selected cell."
        }
      },
      "required": []
    }
  },
  {
    "name": "fillMissingData",
    "description": "Intelligently fills missing or empty cells in a data range using appropriate statistical methods. For numeric columns, fills with median values. For categorical columns, uses the most frequent value (mode). For time series data, applies forward fill (uses the previous non-empty value). Automatically detects column types and highlights filled cells with light blue color for easy identification. Note: for explicit 'fill down' tasks (propagate each non-empty value to blank cells below it until the next non-empty value), use executeOfficeJs instead — this tool's categorical fill uses mode, not fill-down.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to fill missing data (e.g., 'A1:D10'). REQUIRED unless the user explicitly says 'current selection' or 'selected cells'. If the user identified the column or table by name, call readSheetContext first to get the address, then pass it here."
        }
      },
      "required": []
    }
  },
  {
    "name": "fixFormula",
    "description": "Scans cells for formulas that produce errors and fixes them using AI. Detects #DIV/0!, #REF!, #NAME?, #VALUE!, #N/A, #NULL!, #NUM! errors. Sends all formulas as context with error markers so AI can understand patterns and fix broken formulas accurately. When no range is specified, scans ALL sheets in the workbook. When a range is specified, scans only that range on the active sheet.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to scan for formula errors (e.g., 'A1:D10'). If omitted, scans all sheets in the workbook. Provide only when the user explicitly wants to fix a specific range on the active sheet."
        }
      },
      "required": []
    }
  },
  {
    "name": "getSpreadsheetApiReference",
    "description": "Returns Office.js Excel API reference for writing executeOfficeJs code. Call this BEFORE executeOfficeJs when you are unsure about an API area, and AFTER an executeOfficeJs error (especially 'is not a function', 'Cannot read', PropertyNotLoaded, InvalidArgument) to self-correct. Two modes: (1) topics — curated compact reference for one or more areas; (2) className — live introspection of the members available on an Excel.* class in this Excel version. Prefer topics for general questions, add className only when you suspect a specific member name but do not know its exact signature. Calling with no arguments returns the topic guide (key → one-line purpose) so you can pick the right area cheaply.",
    "parameters": {
      "type": "object",
      "properties": {
        "topics": {
          "type": "array",
          "items": {
            "type": "string"
          },
          "description": "One or more topic keys. Pick by purpose:\n  - entry: Excel.run context, workbook, worksheets, create-or-get sheet, selection.\n  - ranges: Addressing, getCell/getRangeByIndexes (0-based), offset/resize, used range, surrounding region, shiftCol.\n  - io: load/sync/read, values/formulas/text, 2-D shapes, bulk read-modify-write, payload chunking.\n  - rows_cols: insert/delete/hide rows and columns, bottom-up deletion, autofit.\n  - format: Font, fill, borders, alignment, number format, column width/row height (points).\n  - defined_names: workbook.names / worksheet.names add, read, delete.\n  - validation: range.dataValidation rules (list, whole number, decimal, date, custom), alerts, prompts.\n  - comments: workbook.comments (threaded) and workbook.notes; read by cell.\n  - freeze: worksheet.freezePanes.\n  - find_replace: findOrNullObject, replaceAll, removeDuplicates.\n  - charts: worksheet.charts.add fallback, series, axes, position. Prefer the addChart tool.\n  - pivots: pivotTables.add, row/column/data/filter hierarchies, refresh, delete.\n  - tables: worksheet.tables add/read/modify, styles, totals, rows/columns.\n  - cond_format: conditionalFormats add/inspect/clear; prefer specialized CF tools for primary tasks.\n  - drawings: shapes.addImage (base64), geometric shapes, position/size in points, alt text.\n  - protection: worksheet.protection, locked cells.\n  - patterns: proven recipes: formula column, bulk transform, bottom-up delete, lookup column, audit.\n  - gotchas: Canonical ✗→✓ mapping of mistakes. Call AFTER an executeOfficeJs error.\n  - hyperlinks: range.hyperlink create/read/clear.\n  - page_setup: worksheet.pageLayout: orientation, paper size, print area, margins, headers/footers.\nCommon aliases are accepted (e.g. 'names' → defined_names, 'conditional' → cond_format, 'shapes'/'images' → drawings, 'chart' → charts, 'replace' → find_replace)."
        },
        "className": {
          "type": "string",
          "description": "Optional Excel.* class name for live introspection (lists its properties and methods in this Excel version). Examples: Excel.Workbook, Excel.Worksheet, Excel.Range, Excel.RangeFormat, Excel.RangeFont, Excel.RangeFill, Excel.RangeBorder, Excel.Table, Excel.Chart, Excel.ChartSeries, Excel.PivotTable, Excel.ConditionalFormat, Excel.DataValidation, Excel.NamedItem, Excel.Comment, Excel.Shape, Excel.WorksheetFreezePanes, Excel.WorksheetProtection, Excel.PageLayout."
        }
      },
      "required": []
    }
  },
  {
    "name": "getCellDetails",
    "description": "READ. Returns detailed cell info grouped by style similarity (compact format). Use for auditing, formula inspection, or style checks. For bulk data reading (values only) use getRangeData instead.\n\nOutput: {hasMore, ranges:[{range, cells:{addr:value|[value,formula]}, styles:{range:{b,i,sz,fn,fc,bg,fmt,wt,or}}}]}. Cells with formulas: [value, formula]. Style keys: b=bold, i=italic, sz=fontSize, fn=fontName, fc=fontColor(#rrggbb), bg=fillColor(#rrggbb), fmt=numberFormat, wt=wrapText, or=orientation(degrees). Only non-default values included. Colors are #rrggbb strings and can be used directly in Office.js (e.g. range.format.fill.color = \"#FFC7CE\").",
    "parameters": {
      "type": "object",
      "properties": {
        "sheet": {
          "type": "string",
          "description": "Sheet name. If omitted, uses the active sheet."
        },
        "ranges": {
          "type": "array",
          "items": {
            "type": "string"
          },
          "description": "A1 notation ranges, e.g. [\"G2:G20\", \"B1:F1\"]."
        },
        "includeStyles": {
          "type": "boolean",
          "description": "Include styles in response (default: true). Set false when you only need values and formulas.",
          "default": true
        },
        "cellLimit": {
          "type": "number",
          "description": "Max total cells across all ranges (default: 200). Ranges are truncated at row boundaries to fit within the limit.",
          "default": 200
        }
      },
      "required": [
        "ranges"
      ]
    }
  },
  {
    "name": "getRangeData",
    "description": "READ. Returns headers (first-row array) and rows (2-D array, empty cells null). Leading and trailing empty rows/columns are trimmed to the data bounding box; returnedRange shows actual extent. If nextRange is not null, call again with it as range and includeHeaders:false. Reading a small sheet (up to ~500 data rows) in full is fine. Do NOT dump larger sheets just to LOCATE values/patterns - use searchData (wildcards/regex), or a filtering executeOfficeJs for logic search cannot express.",
    "parameters": {
      "type": "object",
      "properties": {
        "sheet": {
          "type": "string",
          "description": "Sheet name. If omitted, uses the active sheet."
        },
        "range": {
          "type": "string",
          "description": "A1 notation range to read, e.g. \"A1:D10\" or \"H12:N200\"."
        },
        "maxRows": {
          "type": "number",
          "description": "Maximum rows to return. Default: 500. If nextRange is not null, call again with it."
        },
        "includeHeaders": {
          "type": "boolean",
          "description": "If true (default), the first row of the bounding box is extracted into headers[]; rows[] contains the remaining rows. Set false for continuation ranges where all rows are data and headers[] should be empty."
        }
      },
      "required": [
        "range"
      ]
    }
  },
  {
    "name": "getSheetObjects",
    "description": "READ. Returns charts, pivot tables, Excel tables, conditional formatting rules, drawings (images, shapes), named ranges, comments, hyperlinks (target is a URL or file path), and freeze panes state. Each sheet entry includes freezePanes: { frozenRows, frozenColumns }. externalLinks lists the workbook's links to other workbooks, null if they could not be read. Each chart and shape has id and name (pass the id to describeImage / setAltText as objectId, or get a shape with sheet.shapes.getItem(id) in executeOfficeJs) and its position left/top/width/height in points; the object selected on activeSheet, the active sheet, has selected: true. Use this to audit the workbook, inspect existing objects before editing (chart series, pivot field layout, table columns/style/totals, CF rules), find objects on a specific sheet, or locate commented cells before calling executeOfficeJs. This tool is for object discovery; for API method documentation call getSpreadsheetApiReference.",
    "parameters": {
      "type": "object",
      "properties": {
        "sheet": {
          "type": "string",
          "description": "Sheet name to scan. If omitted, scans the entire workbook."
        }
      },
      "required": []
    }
  },
  {
    "name": "insertPivotTable",
    "description": "Creates a pivot table on a new worksheet. Handles simple and medium pivots: row grouping, cross-tabulation (columns), aggregation (sum/count/average/min/max).\nGood for: known field names, repeatable pivot creation, structured data with clear headers.\nNot for: automatic field discovery, guessing columns from raw data, complex layouts with page fields or multiple value fields — use executeOfficeJs for those.\nCall readSheetContext first to get regionAddress (→ range) and column names with their types (text/date → rows/columns; number → valueColumn).\nDo NOT pass guessed column names — only exact header names from readSheetContext. Do NOT omit valueColumn.",
    "parameters": {
      "type": "object",
      "properties": {
        "sheet": {
          "type": "string",
          "description": "Sheet name containing the source data. If omitted, uses the active sheet."
        },
        "range": {
          "type": "string",
          "description": "Cell range of the source table, e.g. 'A1:I85'. Use regionAddress from readSheetContext. If omitted, uses the current region around the selection."
        },
        "rows": {
          "type": "array",
          "description": "Column names for row grouping (text or date columnType). Required unless columns is provided.",
          "items": {
            "type": "string"
          }
        },
        "columns": {
          "type": "array",
          "description": "Column names for cross-tabulation matrix (text columnType). Required unless rows is provided.",
          "items": {
            "type": "string"
          }
        },
        "valueColumn": {
          "type": "string",
          "description": "Column name to aggregate (number columnType)."
        },
        "aggregation": {
          "type": "string",
          "description": "Aggregation function. Default: Sum.",
          "enum": [
            "Sum",
            "Count",
            "Average",
            "Max",
            "Min",
            "CountNumbers",
            "Product",
            "StdDev",
            "StdDevP",
            "Var",
            "VarP"
          ]
        },
        "outputSheetName": {
          "type": "string",
          "description": "New pivot sheet name (max 31 chars, Excel limit). Auto-generated from field names if omitted."
        }
      },
      "required": [
        "valueColumn"
      ]
    }
  },
  {
    "name": "readSheetContext",
    "description": "Scans the sheet and returns every data region found (contiguous blocks separated by empty rows/columns).\n\nWHEN TO CALL: Before any action on existing data — executeOfficeJs, sort, filter, conditional formatting, chart, or any tool that needs column letters or row bounds. Do NOT guess column letters.\nWHEN TO REUSE: Reuse the result if you already scanned the same sheet in this conversation. Re-scan only if: (a) the user switched sheets, (b) a tool modified the sheet structure (added/deleted columns or rows), or (c) you need more sampleRows than previously fetched.\nACTIVE SHEET: This tool does NOT change the active sheet. The result always includes the real activeSheet from the API. If activeSheet differs from the sheet you scanned, switch explicitly before specialized tools: executeOfficeJs({code:'context.workbook.worksheets.getItem(\"Sales\").activate(); await context.sync();'}).\nSCAN LIMITS: The scanner checks up to 300 rows × 150 columns per sheet. Tables that start within the first 300 rows are fully captured regardless of their total size.\nA sheet lists at most 50 regions (15 in allSheets mode, 120 over the whole workbook): the largest ones and the active one stay, tablesOmitted counts the rest. In allSheets mode a region lists at most 30 cols, colsOmitted counts the rest; a call with sheet lists up to 50 regions and 150 cols.\n\nOUTPUT (single-sheet mode):\n  fileName          — workbook file name\n  activeSheet       — real active sheet name (may differ from the sheet param you passed)\n  sheets[]          — all sheet names in the workbook\n  selection         — current selection address (e.g. \"B4\")\n  activeTableIndex  — index of the table closest to the current selection\n  tables[]          — every detected data region; use tables[activeTableIndex] as primary\n    regionAddress     — bounding box including the header row (e.g. \"A1:G26\")\n    headerRow         — 1-based row number of the header (e.g. 3)\n    dataRowCount      — number of data rows; use for loop bounds: headerRow+1 … headerRow+dataRowCount\n    nextCol           — first free column letter after the table (use this to append new columns)\n    cols[]            — canonical column metadata; cols[i] matches sampleRows[*][i]\n      letter           — column letter (e.g. \"C\")\n      header           — header text, or null if the header cell is empty\n      type             — \"text\"|\"number\"|\"date\"|\"formula\"|\"empty\"\n                         RULE: do NOT write to \"formula\" columns — add a new column instead\n      formulaExample   — (formula columns only, single-sheet mode) one example formula\n    sampleRows[]      — array of up to N data rows; values are native types (null for empty cells)\n    warning           — present when many headers are null; inspect sampleRows before inferring column meaning\n  To find a named column: scan cols[] for a header match and use its letter.\n\nOUTPUT (allSheets mode, allSheets:true):\n  activeSheet          — real active sheet name\n  sheets[]             — all sheet names\n  allSheetsContext[]   — one entry per sheet: { sheetName, tables[], activeTableIndex }\n  Tables keep the same shape, but sampleRows is empty by default and formulaExample is omitted.\n\nWORKBOOK OVERVIEW: For structure/walk-through/audit requests, ALWAYS use allSheets:true. Do NOT use getCellDetails or getRangeData for this purpose.",
    "parameters": {
      "type": "object",
      "properties": {
        "sheet": {
          "type": "string",
          "description": "Sheet name to read from. OMIT unless the user explicitly named a sheet — never guess (e.g. 'Sheet1'). Default: active sheet."
        },
        "sampleRows": {
          "type": "number",
          "description": "Number of data rows (after header) to include as samples. Default 3, max 5, use 0 for headers only."
        },
        "allSheets": {
          "type": "boolean",
          "description": "If true, scans ALL sheets in the workbook and returns allSheetsContext: array of {sheetName, tables[], activeTableIndex} per sheet. Useful for workbook structure analysis. sampleRows defaults to 0 (headers only) when allSheets is true to keep output compact."
        }
      },
      "required": []
    }
  },
  {
    "name": "searchData",
    "description": "Search for text or a value across the spreadsheet. Returns results grouped by sheet as [{sheet, cells}]; each cell entry has an addr (A1 notation) and value; formula is included only when the cell contains one. Supports case-sensitive, whole-cell, formula-text matching, Excel-style wildcards (* and ?) via options.wildcards, and JS regular expressions via options.regex. Searches all sheets by default. Use offset for pagination; follow nextOffset for the next page.",
    "parameters": {
      "type": "object",
      "properties": {
        "searchTerm": {
          "type": "string",
          "description": "The text to search for in the spreadsheet."
        },
        "sheetName": {
          "type": "string",
          "description": "Limit search to a specific worksheet by name. If omitted, searches all sheets."
        },
        "range": {
          "type": "string",
          "description": "A1 notation range to limit the search scope, e.g. 'A1:D100'. If omitted, searches the entire used range."
        },
        "offset": {
          "type": "number",
          "description": "Number of results to skip for pagination (default: 0). Use nextOffset from a previous response to get the next page."
        },
        "options": {
          "type": "object",
          "description": "Optional search configuration.",
          "properties": {
            "matchCase": {
              "type": "boolean",
              "description": "Case-sensitive search. Default: false."
            },
            "matchEntireCell": {
              "type": "boolean",
              "description": "Match entire cell content only (no partial matches). Default: false."
            },
            "matchFormulas": {
              "type": "boolean",
              "description": "Search in formula text instead of displayed values. Default: false."
            },
            "wildcards": {
              "type": "boolean",
              "description": "Excel-style wildcards in searchTerm: * = any sequence, ? = one character, ~ escapes a literal * or ?. Default: false."
            },
            "regex": {
              "type": "boolean",
              "description": "Treat searchTerm as JavaScript RegExp pattern text, without /.../ delimiters or flags (e.g. \\d{4,6}). Case sensitivity follows matchCase. Cannot be combined with wildcards. Default: false."
            },
            "maxResults": {
              "type": "number",
              "description": "Maximum number of results to return per page. Default: 500."
            }
          }
        }
      },
      "required": [
        "searchTerm"
      ]
    }
  },
  {
    "name": "setAltText",
    "description": "Saves the alternative text of an object on a sheet (a picture, a shape): its title and its description, the same fields as in the object properties. The previous alternative text of the object is replaced, so when it already has one, ask the user before calling this tool. Take objectId (the shape id) from getSheetObjects or from the object analyzed by describeImage. The alternative text of a chart cannot be saved.",
    "parameters": {
      "type": "object",
      "properties": {
        "objectId": {
          "type": "string",
          "description": "The id of the shape, from getSheetObjects."
        },
        "title": {
          "type": "string",
          "description": "A short title of the alternative text."
        },
        "description": {
          "type": "string",
          "description": "The description of the alternative text: what the object shows."
        }
      },
      "required": [
        "objectId",
        "title",
        "description"
      ]
    }
  },
  {
    "name": "setAutoFilter",
    "description": "Applies autofilter to a data range, enabling dropdown filters on column headers. Supports filtering by column number or column name (with fuzzy matching). Offers multiple filter types: value comparison operators (greater than, less than, equals), multiple value selection, top/bottom N items or percentage, color-based filtering (cell background or font color), and dynamic filters. Can be used to filter active selection or specific ranges. IMPORTANT: when filtering by specific named values (xlFilterValues operator with criteria1Array), ALWAYS call getRangeData first to read the column data and pass them in criteria1Array.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to apply autofilter (e.g., 'A1:D10'). If the user explicitly says 'current selection' or 'selected cells' — omit this parameter. Otherwise, ALWAYS determine the range explicitly: if the user named a column or table — call readSheetContext to resolve it; if no range is mentioned at all — call readSheetContext to find the relevant data range on the sheet. Do NOT silently fall back to selection when the user hasn't asked for it."
        },
        "field": {
          "type": "number",
          "description": "Field number for filtering (starting from 1, left-most field)."
        },
        "fieldName": {
          "type": "string",
          "description": "Column name/header for filtering (e.g., 'Name', 'Age'). Will automatically find the column number."
        },
        "criteria1": {
          "type": "string",
          "description": "Filter criteria string for operators (e.g., '>10'), JSON-encoded color object ('{\"r\":255,\"g\":0,\"b\":0}') for color filters, or dynamic filter constant. For multiple value selection use criteria1Array instead."
        },
        "criteria1Array": {
          "type": "array",
          "items": {
            "type": [
              "string",
              "number"
            ]
          },
          "description": "Array of values to filter by when using xlFilterValues operator (e.g., ['John', 'Jane'] or [1, 2, 3])."
        },
        "operator": {
          "type": "string",
          "description": "Filter operator.",
          "enum": [
            "xlAnd",
            "xlOr",
            "xlFilterValues",
            "xlTop10Items",
            "xlTop10Percent",
            "xlBottom10Items",
            "xlBottom10Percent",
            "xlFilterCellColor",
            "xlFilterFontColor",
            "xlFilterDynamic"
          ]
        },
        "criteria2": {
          "type": "string",
          "description": "Second criteria for compound filters (used with xlAnd/xlOr operators)."
        },
        "visibleDropDown": {
          "type": "boolean",
          "description": "Show/hide filter dropdown arrow (default: true)."
        }
      },
      "required": []
    }
  },
  {
    "name": "setCellData",
    "description": "Writes values, formulas, styles, and notes to cells by A1 address, and verifies written formulas by reading back their calculated values.\nUse for simple, targeted edits: plain cell writes, formula fills across a bounded range via copyTo, style / number-format passes. For multi-step logic or anything not covered here, use executeOfficeJs.\nIf cell addresses are not given, call readSheetContext first. Do NOT guess.\nThe response returns writtenCount (total cells written, including copyTo-filled cells), formulaChecks (sampled {addr, formula, value} from written formulas - ALWAYS inspect before reporting success; it checks computed values only, number formats and styles are not read back; formulaChecksTruncated=true means the sample was capped), and duplicates (addresses that appeared more than once in cells[]; last occurrence wins, earlier ones are dropped silently). On any validation error nothing is written - the call is atomic. Fields the tool does not read (bold, fontSize, fill_color, ...) are not applied and come back in ignoredFields; formatting goes into the CSS string style. formulasRejected lists the addresses whose formula Excel did not accept (the cell kept its old content, its copyTo was not filled, it is not counted).",
    "parameters": {
      "type": "object",
      "properties": {
        "sheet": {
          "type": "string",
          "description": "Sheet name. Omit to use the active sheet."
        },
        "cells": {
          "type": "array",
          "description": "Array of cell specs. Each spec has an address (single cell like \"B3\" for values, formulas, notes, and per-cell styles; or a bounded range like \"B3:D6\" for style-only / range-formatting over existing cells). To repeat a value or formula across a range, use a single seed spec plus copyTo — prefer this over many individual specs.",
          "items": {
            "type": "object",
            "properties": {
              "address": {
                "type": "string",
                "description": "Target A1 address. Single cell (e.g. \"B3\") for normal writes, or bounded range (e.g. \"B3:D6\") for style-only / range-formatting over existing cells. A range address cannot be combined with an explicit copyTo field."
              },
              "value": {
                "type": [
                  "string",
                  "number",
                  "boolean",
                  "null"
                ],
                "description": "Static cell value for a single cell, or for a seed cell that will be expanded with copyTo. Do NOT put formulas here; use the formula field instead. Ignored when formula is also set. A numeric-looking string (\"0010\", \"1e3\") is stored as a number, as if typed into the cell; to keep it as text add numberFormat \"@\" in the same spec (applied before the value) or prefix the value with an apostrophe."
              },
              "formula": {
                "type": "string",
                "description": "Formula to set on a single cell, or on a seed cell that will be expanded with copyTo. Overrides value. Do NOT put formulas in value — use this field instead."
              },
              "style": {
                "type": "string",
                "description": "Optional CSS declarations. Include only the properties that should change; omitted properties keep existing formatting.\nBest-supported: font-family, font-size (pt), font-weight, font-style, text-decoration, color, background-color, text-align, vertical-align, white-space, border, border-top/right/bottom/left. Use hex colors like #4472c4. Prefer background-color over background. Use vertical-align:middle, not center.\nWrap (white-space:normal) on long text needs an explicit column width (via executeOfficeJs range.format.columnWidth, in points) plus resizeRows:true - autofit does not widen wrapped cells.\nUnsupported layout CSS (margin, display, flex, border-radius, gradients, box-shadow) is ignored.\nStyle-only (no value or formula) is supported — existing cell content is preserved; can be combined with copyTo to format a range without touching values.\nFor range addresses, borders apply to the range's outer Top/Right/Bottom/Left edges only (inside borders not supported)."
              },
              "note": {
                "type": [
                  "string",
                  "null"
                ],
                "description": "Cell comment text. Pass null to delete an existing comment."
              },
              "numberFormat": {
                "type": "string",
                "description": "Spreadsheet number format string, e.g. #,##0.00 or 0% or mm/dd/yyyy. Applied to the seed and the full copyTo range when copyTo is set, or to the full bounded range when address is a range. It is applied before the value, so \"@\" keeps a numeric-looking string as text; to write a number into a \"@\" cell, also pass a numeric format such as General."
              },
              "copyTo": {
                "type": "string",
                "description": "Single-cell seed only. Fills the seed's value or formula down/across a bounded A1 range (e.g. \"J2:J81\"). Whole-column (J:J), whole-row (2:2), comma-separated, and multi-range values are not supported. Prefer making the seed cell the first cell of the copyTo range (e.g. seed at J2 with copyTo \"J2:J81\"); if the seed is outside the range, it is still written on its own and then copied to the range.\nFormula references follow standard spreadsheet rules: relative refs shift with position, absolute refs with $ stay pinned. Example: seed \"=E2*0.2\" at J2 with copyTo \"J2:J81\" becomes \"=E3*0.2\" at J3, …, \"=E81*0.2\" at J81. Use $ to lock: $E2 locks the column, E$2 locks the row, $E$2 locks both (so \"=$E$2*B2\" keeps $E$2 fixed while B2 shifts with the row)."
              }
            },
            "required": [
              "address"
            ],
            "additionalProperties": false
          }
        },
        "resizeColumns": {
          "type": "boolean",
          "description": "If true, auto-fit all written columns after the write. Prefer this over guessing fixed widths. Pass false for pure style / number-format passes so existing column widths are preserved. Default: true.",
          "default": true
        },
        "resizeRows": {
          "type": "boolean",
          "description": "If true, auto-fit all written rows after the write. Default: false.",
          "default": false
        },
        "validateFormulas": {
          "type": "boolean",
          "description": "If true, sample up to 50 written formulas into formulaChecks as {addr, formula, value}. Plain single-cell formulas take 1 sample each; copyTo ranges take 2 (first and last cell). When the cap is hit, formulaChecksTruncated=true in the response. Default: true.",
          "default": true
        }
      },
      "required": [
        "cells"
      ]
    }
  },
  {
    "name": "setMultiSort",
    "description": "Sorts data by multiple columns (up to 3 levels). Use this when you need to sort by primary, secondary, and tertiary sort keys. Each level can have its own sort order.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to sort (e.g., 'A1:D10'). REQUIRED unless the user explicitly says 'current selection' or 'selected cells'. If the user identified the column or table by name, call readSheetContext first to get the address, then pass it here."
        },
        "key1": {
          "type": [
            "string",
            "number"
          ],
          "description": "First sort field - cell reference, column index (1-based), or column name. If omitted, uses first column."
        },
        "sortOrder1": {
          "type": "string",
          "description": "Sort order for key1: 'xlAscending' or 'xlDescending'.",
          "enum": [
            "xlAscending",
            "xlDescending"
          ],
          "default": "xlAscending"
        },
        "key2": {
          "type": [
            "string",
            "number"
          ],
          "description": "Second sort field - cell reference, column index, or column name. If omitted, uses second column."
        },
        "sortOrder2": {
          "type": "string",
          "description": "Sort order for key2: 'xlAscending' or 'xlDescending'.",
          "enum": [
            "xlAscending",
            "xlDescending"
          ],
          "default": "xlAscending"
        },
        "key3": {
          "type": [
            "string",
            "number"
          ],
          "description": "Third sort field - cell reference, column index, or column name."
        },
        "sortOrder3": {
          "type": "string",
          "description": "Sort order for key3: 'xlAscending' or 'xlDescending'.",
          "enum": [
            "xlAscending",
            "xlDescending"
          ],
          "default": "xlAscending"
        },
        "header": {
          "type": "string",
          "description": "Specifies if first row contains headers: 'xlYes' or 'xlNo'.",
          "enum": [
            "xlYes",
            "xlNo"
          ],
          "default": "xlYes"
        }
      },
      "required": []
    }
  },
  {
    "name": "setSort",
    "description": "Sorts data in a range by a single column in ascending or descending order.",
    "parameters": {
      "type": "object",
      "properties": {
        "range": {
          "type": "string",
          "description": "Cell range to sort (e.g., 'A1:D10'). REQUIRED unless the user explicitly says 'current selection' or 'selected cells'. If the user identified the column or table by name, call readSheetContext first to get the address, then pass it here."
        },
        "key1": {
          "type": [
            "string",
            "number"
          ],
          "description": "Sort field - cell reference (e.g., 'A1'), column index (1-based), or column name. If omitted, uses first column."
        },
        "sortOrder1": {
          "type": "string",
          "description": "Sort order: 'xlAscending' or 'xlDescending'.",
          "enum": [
            "xlAscending",
            "xlDescending"
          ],
          "default": "xlAscending"
        },
        "header": {
          "type": "string",
          "description": "Specifies if first row contains headers: 'xlYes' or 'xlNo'.",
          "enum": [
            "xlYes",
            "xlNo"
          ],
          "default": "xlYes"
        }
      },
      "required": []
    }
  },
  {
    "name": "executeOfficeJs",
    "description": "Executes JavaScript against the open workbook using the Office.js Excel JavaScript API, inside an Excel.run batch.\nUse this tool to perform any spreadsheet operation when no other specialized tool is available.\nThe code is NOT VBA and NOT Office Scripts (no Sub/Dim, no main(workbook), no ExcelScript.*). To READ data: load the properties, await context.sync(), then return the value, e.g. const r = context.workbook.worksheets.getActiveWorksheet().getRange(\"A1\"); r.load(\"values\"); await context.sync(); return r.values[0][0];\n\nROUTING:\n- Before operating on EXISTING data, call \"readSheetContext\" first to get column metadata and table bounds. Do NOT guess column letters.\n- If a specialized tool exists for search/sort/filter/pivot/chart/highlight/format/cell writes, prefer it over executeOfficeJs. To FIND values/patterns: searchData (wildcards/regex), or just read a small sheet (up to ~500 data rows) with getRangeData; code searches only when searchData cannot express the logic, and returns ONLY matching rows (cap ~100).\n- For workbook structure/analysis, prefer readSheetContext({allSheets:true}) and answer in chat.\n- When unsure about any API area (data validation, pivots, named items, conditional formats, charts, freeze panes, protection, etc.) call \"getSpreadsheetApiReference\" with the relevant topics BEFORE writing the code. After an error from executeOfficeJs, call \"getSpreadsheetApiReference\" with topics:[\"gotchas\"] plus the topic for the failing area; for a suspected member-name error pass className (e.g. \"Excel.Range\") instead.\n- Charts over tabular data with headers → always addChart (NOT executeOfficeJs). If addChart failed, re-read headers with readSheetContext and retry addChart with exact names.",
    "parameters": {
      "type": "object",
      "properties": {
        "code": {
          "type": "string",
          "description": "The BODY of an async JavaScript function with the signature async (context, Excel) => { <code> }. context is the Excel.RequestContext of the current Excel.run; Excel is the Excel namespace (enums such as Excel.ChartType, Excel.BorderIndex). The tool performs a final context.sync() after your code returns.\n\nCRITICAL RULES (read these first):\n1. PREREQUISITE — EXISTING DATA: Call \"readSheetContext\" before any code that references existing columns. Use column letters, headerRow, and dataRowCount from tables[activeTableIndex] as literal constants in the code (e.g. const headerRow = 3; const salaryCol = \"L\";) — the JSON is not accessible inside. Do NOT re-derive these at runtime from getUsedRange() or by scanning. Verify column content from sampleRows; header: null in cols means the cell is empty.\n2. LOAD → SYNC → READ: Office.js objects are proxies. Before reading ANY property call obj.load(\"prop1,prop2\"), then await context.sync(), then read it. Reading an unloaded property throws PropertyNotLoaded. Writes (range.values = ..., range.format.font.bold = true) are queued and applied at the next sync — sync before reading back anything you wrote.\n3. FEW SYNCS: NEVER call context.sync() inside a loop over cells or rows. Read a whole block once (range.load(\"values\")), compute in JS, write it back once (range.values = matrix). Typical code needs 1–3 syncs. For very large writes (tens of thousands of cells) write in chunks of ~5,000 rows with a sync between chunks to stay under the request payload limit.\n4. 2-D ARRAYS: values, formulas, formulasR1C1 and numberFormat are ALWAYS 2-D arrays whose shape matches the range exactly: \"A1\" → [[\"x\"]], \"A1:C1\" → [[\"a\",\"b\",\"c\"]], \"A1:A3\" → [[\"a\"],[\"b\"],[\"c\"]]. A wrong shape throws InvalidArgument. Build repeated matrices with Array.from({length: rows}, (_, i) => [ ... ]).\n5. FORMULA PHILOSOPHY: For values derived from other cells (totals, rates, lookups), write a spreadsheet formula — range.formulas = [[\"=SUM(G2:G501)\"]] — NOT a JS-computed static number. Always use range.formulas (en-US syntax), never formulasLocal. To fill a formula down a column, build one formula per row in JS (e.g. formulas = Array.from({length: n}, (_, i) => [\"=E\" + (first + i) + \"*0.2\"])) or assign the same R1C1 formula to every row via formulasR1C1 (e.g. [[\"=RC[-2]*RC[-1]\"]] repeated). Avoid IFERROR unless the hidden error is expected. Do NOT write external-workbook links unless the workbook already uses them AND the user asked. Excel recalculates automatically in automatic mode; to read a computed result, load values after a sync (in manual mode call context.workbook.application.calculate(Excel.CalculationType.full) first). Skip verification reads unless high-stakes, the user asks, the result isn't visible, or a rule says ALWAYS verify (e.g. cross-sheet copy).\n6. LAST STEP: Prefer ending write/format code with targetRange.select() (after sheet.activate() if the sheet may not be active) so the user sees what changed. Skip for pure-read code. For new-sheet output or multi-area edits, select the most relevant result range. select() does not replace a needed verification read — combine them.\n7. SHEET SAFETY: worksheets.getItem(name) THROWS ItemNotFound for a missing sheet; worksheets.add(name) THROWS if the name exists. Create-or-get: let ws = context.workbook.worksheets.getItemOrNullObject(n); await context.sync(); if (ws.isNullObject) ws = context.workbook.worksheets.add(n);. If the sheet exists and the user did not explicitly say to overwrite, pick a unique name or add a numeric suffix — do NOT silently clear existing content.\n8. PURE OFFICE.JS: No VBA syntax (Sub/End Sub, Dim, Set, With, Range(\"A1\").Value), no Office Scripts (function main(workbook), ExcelScript.*). Do NOT wrap the code in Excel.run or in another function — write the body only; use await freely. Use 'return value;' to expose a result — return plain JSON-serializable data (numbers, strings, arrays, plain objects), NEVER Office.js proxy objects. Use 'throw new Error(\"reason\")' to report a failure.\n8a. NO COMMENTS: Return code only — no // comments, no /* */, no explanations, no markdown. Comments waste tokens and are never executed.\n9. REPORTS & ANALYSIS: When the user asks for a report, analysis, or data output with multiple rows/columns, write results to a NEW sheet with proper formatting (headers, borders, number formats). For brief factual answers (1-2 sentences), answer in chat instead.\n10. RETRY SIMPLICITY: On an error, rewrite using only primitives — sheet.getRange(address), range.values / range.formulas with exact 2-D shapes, range.format.font.bold, range.format.fill.color, range.numberFormat — with a single sync at the end. Drop helpers, chained logic and anything not essential.\n11. INCREMENTAL PROGRESS: Split multi-section work (distinct headers/blocks/formulas, whole report layouts) into multiple executeOfficeJs calls — one per logical section. Faster perceived progress; a late failure doesn't discard earlier sections. Do NOT split tightly-coupled pieces (header row + its data rows).\n12. NOT TRANSACTIONAL: If the code fails midway, everything already synced — and the commands queued before the failing one — STAYS applied; nothing is rolled back. Before retrying, check what already exists (sheet created? headers written?) and make setup idempotent (getItemOrNullObject, overwrite instead of append).\n\nContext & table rules:\n- For multi-table or multi-sheet tasks (joins, comparisons, pivots, workbook audits), use the exact metadata of all referenced tables/sheets from readSheetContext. Do NOT assume tables[activeTableIndex] is enough when another table is clearly referenced.\n- If multiple tables could match the request and the target is genuinely ambiguous, do NOT guess. Prefer the selected/current table only when the selection or activeTableIndex clearly disambiguates it.\n- Formula columns: do NOT overwrite — add a new column instead. Load range.formulas to read an expression, or check formulaExample in the column entry.\n- Row bounds: data rows span from headerRow+1 to headerRow+dataRowCount (1-based). Use these values — do NOT use getUsedRange() row counts, which cover the whole used area, not the data region.\n- Next free column: use nextCol from readSheetContext — NEVER compute it via String.fromCharCode(65+colCount), which only works for tables starting at A.\n- Column arithmetic: if you have a range object, use range.getOffsetRange(rows, cols) / range.getResizedRange(deltaRows, deltaCols) — no letter arithmetic needed. If you only have a column letter, use shiftCol (copy verbatim): function shiftCol(col,n){var v=0;for(var i=0;i<col.length;i++)v=v*26+(col.charCodeAt(i)-64);v+=n;var r='';while(v>0){r=String.fromCharCode((v-1)%26+65)+r;v=Math.floor((v-1)/26);}return r;}\n- Index-based addressing (getCell, getRangeByIndexes) is 0-based from the sheet origin: A1 = (0,0). Prefer sheet.getRange(colLetter + row) with letters from readSheetContext.\n- When adding new columns: write the header in headerRow (same row as existing headers), data from headerRow+1 to headerRow+dataRowCount. Use nextCol for the first column, shiftCol(nextCol,1) for the second, etc. Assign each column to a variable and use it consistently. Match adjacent header formatting (bold, fill, font).\n- Row filtering: prefer filtering in place (setAutoFilter). \"keep only matching\" → copy to NEW sheet. Explicit \"delete\" → delete bottom-up.\n\nAPI cheat-sheet (Office.js Excel):\n- Entry: const wb = context.workbook; wb.worksheets.getActiveWorksheet(), wb.worksheets.getItem(name) (throws if missing), wb.worksheets.getItemOrNullObject(name), wb.worksheets.add(name), sheet.activate(), sheet.name; wb.getSelectedRange(), wb.getActiveCell().\n- Ranges: sheet.getRange(\"A1:B10\"), sheet.getRange(\"C:C\"), sheet.getRange(\"5:5\"), sheet.getCell(row, col) and sheet.getRangeByIndexes(row, col, rowCount, colCount) (0-based), range.getCell(r, c), range.getOffsetRange(r, c), range.getResizedRange(dr, dc), range.getEntireRow(), range.getEntireColumn(), range.getSurroundingRegion(), sheet.getUsedRange(true). range.address (after load) includes the sheet name (\"Sheet1!A1:B2\").\n- Cell I/O: range.values (any[][]), range.text (displayed strings), range.formulas, range.formulasR1C1, range.numberFormat, range.valueTypes — all 2-D. Empty cells read as \"\" (empty string), not null.\n- Clear / insert / delete: range.clear(Excel.ClearApplyTo.contents | formats | all); range.insert(Excel.InsertShiftDirection.down | right); range.delete(Excel.DeleteShiftDirection.up | left). Whole rows: sheet.getRange(\"5:5\").delete(Excel.DeleteShiftDirection.up). Delete bottom-up.\n- Copy: dest.copyFrom(source, Excel.RangeCopyType.all | values | formats | formulas). The pasted block takes the SOURCE size, anchored at dest's top-left cell. BEFORE copying formulas to another sheet: relative refs to cells OUTSIDE the copied block re-point to the destination sheet (wrong values, no error) — copy those cells too, or copy values (Excel.RangeCopyType.values). Then ALWAYS verify: sync, load values of 1-2 destination formula cells; zeros/empty mean broken refs.\n- Number format: range.numberFormat = [[\"0.00\"]] (2-D, same shape), e.g. \"#,##0\", \"0%\", \"mm/dd/yyyy\", \"@\" for text. Set \"@\" BEFORE writing values that must stay text.\n- Sizing (points): range.format.columnWidth = 120; range.format.rowHeight = 20; range.format.autofitColumns(); range.format.autofitRows(). Wrapped text needs an explicit column width, then autofitRows() LAST.\n- Formatting: range.format.font.bold / .italic / .strikethrough = true; .underline = \"Single\" | \"None\"; .size = 12; .name = \"Arial\"; .color = \"#1F4E79\". Fill: range.format.fill.color = \"#FFF2CC\"; range.format.fill.clear(). Colors are \"#RRGGBB\" strings.\n- Borders: const b = range.format.borders.getItem(Excel.BorderIndex.edgeBottom); b.style = \"Continuous\"; b.weight = \"Thin\" | \"Medium\" | \"Thick\"; b.color = \"#000000\". Indexes: edgeTop, edgeBottom, edgeLeft, edgeRight, insideHorizontal, insideVertical, diagonalDown, diagonalUp. Styles: Continuous, Dash, DashDot, DashDotDot, Dot, Double, None.\n- Alignment: range.format.horizontalAlignment = \"Left\" | \"Center\" | \"Right\" | \"Justify\"; range.format.verticalAlignment = \"Top\" | \"Center\" | \"Bottom\"; range.format.wrapText = true. Merge: range.merge(false); range.unmerge().\n- Find / replace: range.findOrNullObject(text, { completeMatch, matchCase }) (find() throws when nothing matches); range.replaceAll(text, replacement, { completeMatch, matchCase }) returns a count (load via sync). Remove duplicates: range.removeDuplicates([0, 1], true) — 0-based column offsets within the range, includesHeader.\n- Data validation: range.dataValidation.rule = { list: { inCellDropDown: true, source: \"Opt1,Opt2,Opt3\" } }; range.dataValidation.errorAlert = { showAlert: true, style: \"Stop\", title: \"Invalid input\", message: \"Pick a value from the list.\" }; range.dataValidation.prompt = { showPrompt: true, title: \"Status\", message: \"Choose from the list.\" }; range.dataValidation.clear().\n- Named items: wb.names.add(\"Prices\", sheet.getRange(\"B2:B100\")); wb.names.getItemOrNullObject(\"Prices\"); sheet.names for sheet scope.\n- Comments / notes: wb.comments.add(\"Sheet1!B5\", \"text\"); wb.comments.getItemByCell(\"Sheet1!B5\") (throws if none); wb.notes.add(\"Sheet1!B5\", \"text\") where notes are supported.\n- Hyperlinks: range.hyperlink = { address: \"https://example.com\", textToDisplay: \"Site\", screenTip: \"Open\" } or { documentReference: \"Sheet2!A1\", textToDisplay: \"Go\" }.\n- Freeze panes: sheet.freezePanes.freezeRows(n), .freezeColumns(n), .freezeAt(range), .unfreeze(); sheet.freezePanes.getLocationOrNullObject().\n- Protection: sheet.protection.protect(options?, password?), sheet.protection.unprotect(password?); range.format.protection.locked = false. A protected sheet rejects writes (AccessDenied) — tell the user instead of unprotecting without permission.\n- Tables: const t = sheet.tables.add(\"A1:D10\", true); t.name = \"Sales\"; t.style = \"TableStyleLight9\"; t.showTotals = true; t.columns.getItem(\"Amount\").totals... ; t.rows.add(null, [[...]]). A table cannot overlap another table or merged cells.\n- Recalculation: context.workbook.application.calculate(Excel.CalculationType.full).\n\nSPECIALIZED TOOLS — use these instead of executeOfficeJs when the operation is the primary goal (not a minor step inside a larger executeOfficeJs):\n  Charts → addChart\n  Pivot tables → insertPivotTable supports rows, columns (cross-tabulation), and aggregation (Sum/Count/Average/Max/Min and more). Use executeOfficeJs only for filter fields, multiple value fields, or modifying an existing pivot (see pivots topic).\n  Sort → setSort / setMultiSort\n  AutoFilter → setAutoFilter\n  Cell values, formulas, styles, number formats, notes → setCellData\n  Conditional formatting → addCellValueCondition, addColorScale, addDataBars, addIconSet, addTop10Condition\n  For complex tasks, prefer sequential small tool calls over one giant executeOfficeJs. Trivial formatting (bold a header, set a number format) inside executeOfficeJs is fine — no need to split.\n  Cross-sheet: specialized action tools operate on the ACTIVE sheet unless they take a sheet parameter. If the target sheet is not active, first run executeOfficeJs({code: 'context.workbook.worksheets.getItem(\"SheetName\").activate(); await context.sync();'}). executeOfficeJs itself can access any sheet via worksheets.getItem(name).\nFALLBACK API — use ONLY when the specialized tool cannot handle the request:\n  Charts: const ch = sheet.charts.add(Excel.ChartType.columnClustered, sheet.getRange(\"A1:D10\"), Excel.ChartSeriesBy.auto); ch.title.text = \"Title\"; ch.setPosition(\"F2\", \"M20\"); series via ch.series.\n  Pivot: sheet.pivotTables.add(name, sourceRange, destinationRange); pivot.rowHierarchies.add(pivot.hierarchies.getItem(\"Region\")); const dh = pivot.dataHierarchies.add(pivot.hierarchies.getItem(\"Sales\")); dh.summarizeBy = Excel.AggregationFunction.sum; filters via pivot.filterHierarchies. Call getSheetObjects first to inspect existing pivots, then getSpreadsheetApiReference({topics:['pivots']}).\n  Sort: range.sort.apply([{ key: 2, ascending: false }], false, true) — key is the 0-based column offset within the range; last argument = hasHeaders.\n  AutoFilter: sheet.autoFilter.apply(range, columnIndex, { filterOn: Excel.FilterOn.values, values: [\"A\", \"B\"] }); sheet.autoFilter.clearCriteria(); sheet.autoFilter.remove().\n  Conditional formatting: const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue); cf.cellValue.format.fill.color = \"#FFC7CE\"; cf.cellValue.rule = { formula1: \"=100\", operator: \"GreaterThan\" }. Other types: colorScale, dataBar, iconSet, topBottom, presetCriteria (duplicateValues, uniqueValues, aboveAverage, ...), containsText, custom (formula). Remove: range.conditionalFormats.clearAll(). Inspect existing rules via getSheetObjects, then getSpreadsheetApiReference({topics:['cond_format']})."
        }
      },
      "required": [
        "code"
      ]
    }
  }
]
~~~~~