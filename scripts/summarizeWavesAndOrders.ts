function main(
  workbook: ExcelScript.Workbook,
  jsonText: string,
  targetYmd: string // "YYYY-MM-DD" from Flow
): string {
  type InputRow = Record<string, unknown>;
  interface WaveRow {
    "Wave #": string;
    "Order Count": number;
    "Source Codes": string;
    "Inducted ?": string;     // "✓" if CHUTE non-blank count > 5
    "Req Ship Date": string;  // NEW: carries targetYmd
  }

  const COL_SOURCE = "Source Code";
  const COL_ORDER = "Order Num";
  const COL_WAVE = "Wave#";
  const COL_REQ = "Req Ship Date";
  const COL_CHUTE = "CHUTE";

  const toStr = (v: unknown): string => (v ?? "").toString().trim();
  const excelSerialToDate = (serial: number): Date => new Date(Math.round((serial - 25569) * 86400 * 1000));
  const pad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));
  const ymd = (d: Date): string => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  const toYmd = (v: unknown): string => {
    if (typeof v === "number" && isFinite(v)) return ymd(excelSerialToDate(v));
    const s = toStr(v);
    if (!s) return "";
    const t = Date.parse(s);
    return Number.isNaN(t) ? "" : ymd(new Date(t));
  };
  const isObject = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null;
  const isRecordArray = (v: unknown): v is InputRow[] => Array.isArray(v) && v.every((el) => typeof el === "object" && el !== null);

  const ymdRegex = /^\d{4}-\d{2}-\d{2}$/;
  const effectiveYmd = ymdRegex.test(targetYmd) ? targetYmd : ymd(new Date());

  let rows: InputRow[] = [];
  try {
    const parsed: unknown = JSON.parse(jsonText);
    if (isRecordArray(parsed)) rows = parsed;
    else if (isObject(parsed) && isRecordArray((parsed as { rows?: unknown }).rows)) rows = (parsed as { rows?: unknown }).rows as InputRow[];
    else return JSON.stringify({ success: false, message: "Input JSON does not contain an array of rows.", rows: [] as WaveRow[] });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return JSON.stringify({ success: false, message: `Invalid JSON: ${msg}`, rows: [] as WaveRow[] });
  }

  if (rows.length === 0) return JSON.stringify({ success: true, message: "No input rows.", rows: [] as WaveRow[] });

  const keys = Object.keys(rows[0]);
  const missing: string[] = [];
  for (const req of [COL_SOURCE, COL_ORDER, COL_WAVE, COL_REQ, COL_CHUTE]) if (!keys.includes(req)) missing.push(req);
  if (missing.length) return JSON.stringify({ success: false, message: `Missing required columns: ${missing.join(", ")}`, rows: [] as WaveRow[] });

  const filtered: InputRow[] = rows.filter((r) => toYmd(r[COL_REQ]) === effectiveYmd);
  if (filtered.length === 0) return JSON.stringify({ success: true, message: `No rows with "${COL_REQ}" = ${effectiveYmd}.`, rows: [] as WaveRow[] });

  const waveOrders = new Map<string, Set<string>>();
  const waveSources = new Map<string, Set<string>>();
  const waveChuteNonBlank = new Map<string, number>();

  for (const r of filtered) {
    const wave = toStr(r[COL_WAVE]);
    const order = toStr(r[COL_ORDER]);
    const src = toStr(r[COL_SOURCE]);
    const chute = toStr(r[COL_CHUTE]);
    if (!wave || !order) continue;

    if (!waveOrders.has(wave)) waveOrders.set(wave, new Set<string>());
    waveOrders.get(wave)!.add(order);

    if (!waveSources.has(wave)) waveSources.set(wave, new Set<string>());
    if (src) waveSources.get(wave)!.add(src);

    if (!waveChuteNonBlank.has(wave)) waveChuteNonBlank.set(wave, 0);
    if (chute) waveChuteNonBlank.set(wave, (waveChuteNonBlank.get(wave) ?? 0) + 1);
  }

  const resultRows: WaveRow[] = Array.from(waveOrders.keys()).map((wave) => {
    const orderCount = waveOrders.get(wave)?.size ?? 0;
    const sources = Array.from(waveSources.get(wave) ?? new Set<string>()).sort().join(", ");
    const chuteCount = waveChuteNonBlank.get(wave) ?? 0;
    const inducted = chuteCount > 5 ? "✓" : "";
    return {
      "Wave #": wave,
      "Order Count": orderCount,
      "Source Codes": sources,
      "Inducted ?": inducted,
      "Req Ship Date": effectiveYmd
    };
  });

  resultRows.sort((a, b) => b["Order Count"] - a["Order Count"]);

  return JSON.stringify({
    success: true,
    message: `Summarized ${resultRows.length} unique Waves for ${effectiveYmd}.`,
    rows: resultRows
  });
}
