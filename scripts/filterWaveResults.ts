function main(
    workbook: ExcelScript.Workbook,
    jsonText: string,
    sourceColumnName: string = "Source Code",
    excludeListCsv: string = "Standard",  // comma-separated list; default excludes "Standard"
    caseInsensitive: boolean = true,
    trimSpaces: boolean = true
): string {
    // Types
    type InputRow = Record<string, unknown>;

    // Helpers
    const isObject = (v: unknown): v is Record<string, unknown> =>
        typeof v === "object" && v !== null;
    const isRecordArray = (v: unknown): v is InputRow[] =>
        Array.isArray(v) && v.every((el) => typeof el === "object" && el !== null);
    const norm = (v: unknown): string => {
        let s = (v ?? "").toString();
        if (trimSpaces) s = s.trim();
        return caseInsensitive ? s.toLowerCase() : s;
    };

    // Parse input JSON -> must be an array of row objects (pass-through supported for {rows:[]})
    let rows: InputRow[] = [];
    try {
        const parsed: unknown = JSON.parse(jsonText);
        if (isRecordArray(parsed)) {
            rows = parsed;
        } else if (isObject(parsed) && isRecordArray((parsed as { rows?: unknown }).rows)) {
            // If someone ever passes {rows:[...]} we still handle it gracefully.
            rows = (parsed as { rows?: unknown }).rows as InputRow[];
        } else {
            // If not an array, return original unmodified to avoid breaking the flow
            return jsonText;
        }
    } catch {
        // If parse fails, return original unmodified
        return jsonText;
    }

    if (rows.length === 0) return JSON.stringify(rows);

    // Validate column presence; if missing, do a safe pass-through
    const firstKeys = Object.keys(rows[0]);
    if (!firstKeys.includes(sourceColumnName)) {
        return JSON.stringify(rows);
    }

    // Build exclusion set
    const excludeValues = excludeListCsv
        .split(",")
        .map(s => (trimSpaces ? s.trim() : s))
        .filter(s => s !== "");
    const excludeSet = new Set(excludeValues.map(v => caseInsensitive ? v.toLowerCase() : v));

    // Filter rows: keep if Source Code NOT in excludeSet
    const filtered = rows.filter(r => !excludeSet.has(norm(r[sourceColumnName])));

    return JSON.stringify(filtered);
}
