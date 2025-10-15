function main(workbook: ExcelScript.Workbook, sheetName: string) {
  const START_COLUMN = "A";
  const END_COLUMN = "AB";  // bump to "AC" if needed
  const HEADER_ROW = 1;

  const ws = workbook.getWorksheet(sheetName);
  if (!ws) { console.log(`Worksheet "${sheetName}" not found`); return "[]"; }

  const used = ws.getUsedRange();
  if (!used) { console.log("No data found"); return "[]"; }

  const lastRow = used.getRowCount();
  const data = ws.getRange(`${START_COLUMN}${HEADER_ROW}:${END_COLUMN}${lastRow}`).getValues();

  if (data.length <= 1) { console.log("No data rows"); return "[]"; }

  // Helpers
  const toStr = (v: unknown) => (v ?? "").toString();
  const trimHeader = (v: unknown) =>
    toStr(v).replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim(); // collapse & trim, convert NBSP
  const isBlank = (v: unknown) =>
    v === null || v === undefined || (typeof v === "string" && v.trim() === "");

  // Normalize headers
  const rawHeaders = data[0];
  const headers: string[] = (rawHeaders as unknown[]).map((h) => trimHeader(h));
  console.log(`Headers(norm): ${JSON.stringify(headers)}`);

  // Map original->normalized header for logging/debug
  rawHeaders.forEach((h, i) => {
    if (trimHeader(h) !== toStr(h)) console.log(`Header normalized: "${toStr(h)}" -> "${trimHeader(h)}" @${i}`);
  });

  // Find columns that actually have data (after trimming), OR force-include CHUTE
  const activeCols: number[] = [];
  const forceInclude = new Set(["CHUTE"]); // add more keys you always want present

  for (let j = 0; j < headers.length; j++) {
    const name = headers[j];
    let hasData = false;

    // If forced, mark active regardless
    if (forceInclude.has(name)) {
      activeCols.push(j);
      console.log(`Force-included column: ${name} @ ${j}`);
      continue;
    }

    for (let i = 1; i < data.length; i++) {
      const cell = data[i][j];
      if (!isBlank(typeof cell === "string" ? cell.trim() : cell)) {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      activeCols.push(j);
      console.log(`Active column: ${name} @ ${j}`);
    }
  }

  // Build results
  const results: Array<Record<string, string | number | boolean>> = [];
  for (let i = 1; i < data.length; i++) {
    const rowObj: Record<string, string | number | boolean> = {};
    let any = false;

    for (const j of activeCols) {
      const name = headers[j];
      const val = data[i][j];

      // keep non-blanks; CHUTE gets included if forced and non-blank after trim
      if (!isBlank(typeof val === "string" ? val.trim() : val)) {
        rowObj[name] = typeof val === "string" ? val.trim() : val;
        any = true;
      }
    }

    if (any) results.push(rowObj);
  }

  console.log(`Returning ${results.length} rows`);
  // Return TEXT for Flow

  // Sort by CHUTE (A→Z), blanks last
  results.sort((a, b) => {
    const av = (a?.["CHUTE"] ?? "").toString().trim().toLowerCase();
    const bv = (b?.["CHUTE"] ?? "").toString().trim().toLowerCase();

    if (!av && !bv) return 0;
    if (!av) return 1;   // blanks last
    if (!bv) return -1;

    return av.localeCompare(bv);
  });

  return JSON.stringify(results);
}
