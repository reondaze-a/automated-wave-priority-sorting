function main(
  workbook: ExcelScript.Workbook,
  jsonText: string,
  targetYmd: string // "YYYY-MM-DD" from Flow
): string {
  // ---------- Types ----------
  type InputRow = Record<string, unknown>;
  interface WaveRow {
    "Wave #": string;
    "PCL Order Count": number;
    "LTL Order Count": number;
    "Source Codes": string;
    "Inducted ?": string;     // "✓" if CHUTE non-blank count > 5
    "Req Ship Date": string;  // echo targetYmd
  }

  // ---------- Expected headers ----------
  const COL_SOURCE = "Source Code";
  const COL_ORDER  = "Order Num";
  const COL_WAVE   = "Wave#";
  const COL_REQ    = "Req Ship Date";
  const COL_CHUTE  = "CHUTE";
  const COL_LTL_PCL = "LTL_PCL"; // guaranteed "LTL" or "PCL"

  // ---------- Helpers ----------
  const toStr = (v: unknown): string => (v ?? "").toString().trim();

  // Excel serial → Date
  const excelSerialToDate = (serial: number): Date =>
    new Date(Math.round((serial - 25569) * 86400 * 1000));
  const pad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));
  const ymd = (d: Date): string =>
    `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  const toYmd = (v: unknown): string => {
    if (typeof v === "number" && isFinite(v)) return ymd(excelSerialToDate(v));
    const s = toStr(v);
    if (!s) return "";
    const t = Date.parse(s);
    return Number.isNaN(t) ? "" : ymd(new Date(t));
  };

  // Type guards
  const isObject = (v: unknown): v is Record<string, unknown> =>
    typeof v === "object" && v !== null;
  const isRecordArray = (v: unknown): v is InputRow[] =>
    Array.isArray(v) && v.every((el) => typeof el === "object" && el !== null);

  // Date param (fallback to today if malformed)
  const ymdRegex = /^\d{4}-\d{2}-\d{2}$/;
  const effectiveYmd = ymdRegex.test(targetYmd) ? targetYmd : ymd(new Date());

  // ---------- Parse input JSON ----------
  let rows: InputRow[] = [];
  try {
    const parsed: unknown = JSON.parse(jsonText);
    if (isRecordArray(parsed)) rows = parsed;
    else if (isObject(parsed) && isRecordArray((parsed as { rows?: unknown }).rows))
      rows = (parsed as { rows?: unknown }).rows as InputRow[];
    else
      return JSON.stringify({ success: false, message: "Input JSON does not contain an array of rows.", rows: [] as WaveRow[] });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return JSON.stringify({ success: false, message: `Invalid JSON: ${msg}`, rows: [] as WaveRow[] });
  }

  if (rows.length === 0)
    return JSON.stringify({ success: true, message: "No input rows.", rows: [] as WaveRow[] });

  // ---------- Validate required headers ----------
  const firstKeys = Object.keys(rows[0]);
  const missing: string[] = [];
  for (const req of [COL_SOURCE, COL_ORDER, COL_WAVE, COL_REQ, COL_CHUTE, COL_LTL_PCL]) {
    if (!firstKeys.includes(req)) missing.push(req);
  }
  if (missing.length)
    return JSON.stringify({ success: false, message: `Missing required columns: ${missing.join(", ")}`, rows: [] as WaveRow[] });

  // ---------- Filter to target date ----------
  const filtered: InputRow[] = rows.filter((r) => toYmd(r[COL_REQ]) === effectiveYmd);
  if (filtered.length === 0)
    return JSON.stringify({ success: true, message: `No rows with "${COL_REQ}" = ${effectiveYmd}.`, rows: [] as WaveRow[] });

  // ---------- Aggregate by Wave ----------
  const waveSources = new Map<string, Set<string>>();   // Wave -> distinct Source Codes
  const waveChuteNonBlank = new Map<string, number>();  // Wave -> count of rows with non-blank CHUTE
  const wavePclOrders = new Map<string, Set<string>>(); // Wave -> distinct Order Num where LTL_PCL="PCL"
  const waveLtlOrders = new Map<string, Set<string>>(); // Wave -> distinct Order Num where LTL_PCL="LTL"

  for (const r of filtered) {
    const wave  = toStr(r[COL_WAVE]);
    const order = toStr(r[COL_ORDER]);
    const src   = toStr(r[COL_SOURCE]);
    const chute = toStr(r[COL_CHUTE]);
    const cat   = toStr(r[COL_LTL_PCL]); // guaranteed "LTL" or "PCL"

    if (!wave || !order) continue;

    // sources
    if (!waveSources.has(wave)) waveSources.set(wave, new Set<string>());
    if (src) waveSources.get(wave)!.add(src);

    // inducted counter
    if (!waveChuteNonBlank.has(wave)) waveChuteNonBlank.set(wave, 0);
    if (chute) waveChuteNonBlank.set(wave, (waveChuteNonBlank.get(wave) ?? 0) + 1);

    // category counts (distinct by order)
    if (cat === "PCL") {
      if (!wavePclOrders.has(wave)) wavePclOrders.set(wave, new Set<string>());
      wavePclOrders.get(wave)!.add(order);
    } else if (cat === "LTL") {
      if (!waveLtlOrders.has(wave)) waveLtlOrders.set(wave, new Set<string>());
      waveLtlOrders.get(wave)!.add(order);
    }
  }

  // ---------- Build result rows ----------
  const waves = new Set<string>([
    ...Array.from(waveSources.keys()),
    ...Array.from(wavePclOrders.keys()),
    ...Array.from(waveLtlOrders.keys())
  ]);

  const resultRows: WaveRow[] = Array.from(waves).map((wave) => {
    const pcl = wavePclOrders.get(wave)?.size ?? 0;
    const ltl = waveLtlOrders.get(wave)?.size ?? 0;
    const sources = Array.from(waveSources.get(wave) ?? new Set<string>()).sort().join(", ");
    const chuteCount = waveChuteNonBlank.get(wave) ?? 0;
    const inducted = chuteCount > 5 ? "✓" : "";
    return {
      "Wave #": wave,
      "PCL Order Count": pcl,
      "LTL Order Count": ltl,
      "Source Codes": sources,
      "Inducted ?": inducted,
      "Req Ship Date": effectiveYmd
    };
  });

  // ---------- Sort: PCL Order Count desc (only) ----------
  resultRows.sort((a, b) => b["PCL Order Count"] - a["PCL Order Count"]);

  return JSON.stringify({
    success: true,
    message: `Summarized ${resultRows.length} unique Waves for ${effectiveYmd}.`,
    rows: resultRows
  });
}
