Execute Office.js JavaScript code to interact with the Excel document. The code receives a context parameter and runs inside Excel.run(). Use this for any document operations.

Freeze panes: to keep header rows or columns visible while the user scrolls, call `sheet.freezePanes.freezeRows(count)` / `sheet.freezePanes.freezeColumns(count)` — they freeze the top `count` rows / first `count` columns in place (e.g. `freezeRows(3)` when the header is on row 3, so the title rows above it stay visible too). Avoid `sheet.freezePanes.freezeAt(range)` for this: it makes `range` itself the frozen pane and can leave the rows above it and the columns to its left scrolled out of view.

A `conductor` global is available for sharing files with other agents:
  `conductor.writeFile("data.json", jsonString)`  — broadcast a file to all connected agents
  `conductor.readFile("excel-abc", "data.json")`  — read another agent's file
  `conductor.listFiles("excel-abc")`              — list an agent's shared files

LARGE DATA: When you need to send data to another agent (cell values, tables, extracted content, XML, etc.), write it to a conductor file instead of putting it in the send_message text. Use `execute_office_js` to read the data and call `conductor.writeFile()` with a descriptive filename (e.g. data.json, table.txt, chart.xml). The file is broadcast immediately to all peers and visible in their virtual filesystem. Then call send_message with a short message saying what you want done. Do NOT include instructions on how to read it; the receiving agent can access it directly.

USING RECEIVED FILES: When a peer has shared a file, read AND use it inside the SAME code block. NEVER `return conductor.readFile(...)` — that dumps the entire file into your context.

WORKFLOW: peek at structure via bash `head -5 /agents/<id>/files/<name>`, then come straight here and write the full processing code. The header line is enough — you do NOT need to see every row in your context.

  // ✅ GOOD: after bash head -5 showed "Country,Region,Rate,Year"
  const raw = conductor.readFile("desktop-41c", "data.csv");
  const lines = raw.trim().split("\n");
  const rows = lines.slice(1).map(r => r.split(","));
  const sheet = context.workbook.worksheets.add("Imported");
  sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length).values = rows;
  await context.sync();
  return { rowsWritten: rows.length };

  // ❌ BAD: read-only — whole file leaks into context
  return conductor.readFile("desktop-41c", "data.csv");

If structure is genuinely unclear after `head`, return a small probe (e.g. `return { keys: Object.keys(JSON.parse(raw)), len: raw.length }`) — never the full contents.

LARGE DATA: [see note below]

A `blobs` global provides in-memory storage local to this browser sandbox. It persists across execute_office_js calls for the conversation, but is NOT visible to the server-side `bash` or `code_execution` containers:
  `blobs.getText("name")`   — decoded UTF-8 string or null (for .csv, .json, .html, code files)
  `blobs.getJSON("name")`   — parsed JSON or null
  `blobs.get("name")`       — Uint8Array or null (raw bytes)
  `blobs.getBase64("name")` — base64 string or null (for insert*FromBase64)
  `blobs.has("name")`       — check if blob exists
  `blobs.keys()`            — list stored blob names
  `blobs.setText("name", s)` / `blobs.setJSON("name", v)` — stash data for later calls (5MB cap)

STASHING LARGE RESULTS: when a fetch (Graph, file read) would overflow the 40k-char tool-result limit, write the full payload to `blobs.setJSON()` and return only a small summary (`{count, sample: items.slice(0,3)}`). Later calls read it back with `blobs.getJSON()` inside the sandbox — the full data never enters context. `conductor.writeFile` is for broadcasting to OTHER agents; it is NOT a local stash and is NOT readable from the `bash` or `code_execution` sandboxes — use `blobs` for that.

The `blobs` store also holds SKILL files (auto-populated by `read_skill`) and user-uploaded images (pre-loaded, keyed by `file_id`). For other uploaded files (`container_upload`: .xlsx, .pdf, .csv, .docx), use `code_execution` (Python) — see the File Uploads section of the system prompt.

An `attachImage(base64, mediaType?)` global lets you return an image you can SEE in the tool result (as a vision content block, not text). Call it with the `.value` of any Office.js `getImage()` after `context.sync()`:
  `const img = sheet.getRange("A1:F20").getImage(); await context.sync(); attachImage(img.value);`
Also works for `chart.getImage()`. Use it when visual layout/formatting matters and cell values alone aren't answering the question. Keep rendered ranges under ~50 rows — larger produces illegible text.