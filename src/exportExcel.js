import XLSX from "xlsx-js-style";

/* ═══════════════════ STYLE CONSTANTS ═══════════════════ */
const THIN_BORDER = {
  top:    { style: "thin", color: { rgb: "000000" } },
  bottom: { style: "thin", color: { rgb: "000000" } },
  left:   { style: "thin", color: { rgb: "000000" } },
  right:  { style: "thin", color: { rgb: "000000" } },
};

const LEGEND_HEADER_STYLE = {
  fill: { fgColor: { rgb: "000000" } },
  font: { bold: true, color: { rgb: "FFFFFF" } },
  border: THIN_BORDER,
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
};

const LEGEND_DATA_STYLE = {
  border: THIN_BORDER,
  alignment: { wrapText: true, vertical: "top" },
};

const META_LABEL_STYLE = {
  fill: { fgColor: { rgb: "87CEEB" } },
  font: { bold: true },
  border: THIN_BORDER,
  alignment: { vertical: "top", wrapText: true },
};

const META_VALUE_STYLE = {
  border: THIN_BORDER,
  alignment: { horizontal: "left", vertical: "top", wrapText: true },
};

const DATA_HEADER_STYLE = {
  fill: { fgColor: { rgb: "87CEEB" } },
  font: { bold: true },
  border: THIN_BORDER,
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
};

const DATA_CELL_STYLE = {
  border: THIN_BORDER,
  alignment: { wrapText: true, vertical: "top" },
};

const DATA_CELL_LEFT_STYLE = {
  border: THIN_BORDER,
  alignment: { horizontal: "left", wrapText: true, vertical: "top" },
};

/* ═══════════════════ HELPERS ═══════════════════ */
function str(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function autoHeight(text, colWidth, basePt) {
  const s = str(text);
  if (!s) return basePt;
  const lines = Math.ceil(s.length / Math.max(colWidth, 10));
  return Math.min(60, Math.max(basePt, lines * 15));
}

function clampWidth(len, min, max) {
  return Math.max(min, Math.min(max, len + 2));
}

/* ═══════════════════ LEGEND SHEET ═══════════════════ */
async function buildLegendSheet(wb) {
  const resp = await fetch("/legend.csv");
  const text = await resp.text();
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length === 0) return;

  // Parse CSV rows — handle quoted fields that may contain commas
  const rows = lines.map((line) => {
    const cells = [];
    let cur = "";
    let inQuote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuote) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') { inQuote = false; }
        else { cur += ch; }
      } else {
        if (ch === '"') { inQuote = true; }
        else if (ch === ',') { cells.push(cur.trim()); cur = ""; }
        else { cur += ch; }
      }
    }
    cells.push(cur.trim());
    return cells;
  });
  const headers = rows[0];
  const data = rows.slice(1);

  const ws = XLSX.utils.aoa_to_sheet([]);

  // Track max widths per column
  const colMaxLen = headers.map((h) => h.length);

  // Header row
  headers.forEach((h, c) => {
    const ref = XLSX.utils.encode_cell({ r: 0, c });
    ws[ref] = { v: h, t: "s", s: LEGEND_HEADER_STYLE };
  });

  // Data rows
  data.forEach((row, ri) => {
    row.forEach((val, ci) => {
      const ref = XLSX.utils.encode_cell({ r: ri + 1, c: ci });
      ws[ref] = { v: val, t: "s", s: LEGEND_DATA_STYLE };
      if (val.length > (colMaxLen[ci] || 0)) colMaxLen[ci] = val.length;
    });
  });

  // Set range
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: data.length, c: headers.length - 1 },
  });

  // Column widths (auto-fit, max 60)
  ws["!cols"] = colMaxLen.map((len) => ({ wch: clampWidth(len, 8, 60) }));

  // Row heights
  const rowHeights = [{ hpt: 20 }]; // header
  data.forEach((row) => {
    const maxLen = Math.max(...row.map((v) => v.length), 1);
    rowHeights.push({ hpt: Math.min(60, Math.max(16, Math.ceil(maxLen / 50) * 15)) });
  });
  ws["!rows"] = rowHeights;

  XLSX.utils.book_append_sheet(wb, ws, "Legend");
}

/* ═══════════════════ POC SHEET ═══════════════════ */
const DATA_COLUMNS = [
  { key: "empId",        label: "EmpID" },
  { key: "name",         label: "Name" },
  { key: "location",     label: "Location" },
  { key: "country",      label: "Country" },
  { key: "actPct",       label: "ACT/PCT" },
  { key: "skillSet",     label: "Skill Set" },
  { key: "verizonLevel", label: "Verizon Level Mapping" },
  { key: "classification", label: "Classification" },
  { key: "key",          label: "Key" },
  { key: "designation",  label: "Cognizant Designation" },
  { key: "serviceLine",  label: "Service Line" },
  { key: "timesheetHrs", label: "Timesheet" },
  { key: "rateInr",      label: "Hourly Rate(Rs)" },
  { key: "rateUsd",      label: "Hourly Rate($)" },
  { key: "projectedRate",label: "Projected Rate($)" },
  { key: "actualRate",   label: "Actual Rate" },
  { key: "variance",     label: "Variance" },
  { key: "jan26",        label: "Jan-26" },
  { key: "feb26",        label: "Feb-26" },
  { key: "mar26",        label: "Mar-26" },
];

function buildPocSheet(wb, pocName, pocRows) {
  const ws = {};

  // --- Derive meta from rows (find first non-empty value for each field) ---
  const findFirst = (field) => {
    for (const row of pocRows) {
      const v = str(row[field]);
      if (v) return v;
    }
    return "";
  };
  const meta = [
    { label: "ESA ID",                value: findFirst("esaId") },
    { label: "ESA Description",       value: findFirst("esaDesc") },
    { label: "Verizon TQ ID",         value: findFirst("vzTqId") },
    { label: "Verizon TQ Description",value: findFirst("vzTqDesc") },
    { label: "POC",                   value: findFirst("poc") },
  ];

  // Longest meta label / value for col width calcs
  let maxMetaLabel = 0;
  let maxMetaValue = 0;
  meta.forEach((m) => {
    if (m.label.length > maxMetaLabel) maxMetaLabel = m.label.length;
    if (m.value.length > maxMetaValue) maxMetaValue = m.value.length;
  });

  // --- Zone 1: Meta rows (rows 0-4 = Excel rows 1-5) ---
  meta.forEach((m, ri) => {
    const refA = XLSX.utils.encode_cell({ r: ri, c: 0 });
    const refB = XLSX.utils.encode_cell({ r: ri, c: 1 });
    ws[refA] = { v: m.label, t: "s", s: META_LABEL_STYLE };
    ws[refB] = { v: m.value, t: "s", s: META_VALUE_STYLE };
  });

  // --- Zone 2: Spacer rows (rows 5-6 = Excel rows 6-7) ---
  // Just leave them empty; we'll set row heights

  // --- Zone 3: Data header (row 7 = Excel row 8) ---
  const headerRow = 7;
  DATA_COLUMNS.forEach((col, ci) => {
    const ref = XLSX.utils.encode_cell({ r: headerRow, c: ci });
    ws[ref] = { v: col.label, t: "s", s: DATA_HEADER_STYLE };
  });

  // --- Zone 3: Data rows (row 8+ = Excel rows 9+) ---
  pocRows.forEach((row, ri) => {
    const excelRow = headerRow + 1 + ri;
    DATA_COLUMNS.forEach((col, ci) => {
      const ref = XLSX.utils.encode_cell({ r: excelRow, c: ci });
      const raw = row[col.key];
      const isNum = typeof raw === "number";
      const style = ci === 0 ? DATA_CELL_LEFT_STYLE : DATA_CELL_STYLE;
      ws[ref] = {
        v: isNum ? raw : str(raw),
        t: isNum ? "n" : "s",
        s: style,
      };
    });
  });

  // Set range
  const lastDataRow = headerRow + pocRows.length;
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: lastDataRow, c: DATA_COLUMNS.length - 1 },
  });

  // --- Column widths ---
  // Track max content length per column across header + data
  const colMaxLen = DATA_COLUMNS.map((col) => col.label.length);
  pocRows.forEach((row) => {
    DATA_COLUMNS.forEach((col, ci) => {
      const v = str(row[col.key]);
      if (v.length > colMaxLen[ci]) colMaxLen[ci] = v.length;
    });
  });

  const colWidths = colMaxLen.map((len) => ({ wch: clampWidth(len, 8, 50) }));
  // Col A min = longest meta label + 2, Col B min = longest meta value + 2 (max 80)
  colWidths[0].wch = Math.max(colWidths[0].wch, maxMetaLabel + 2);
  if (colWidths.length > 1) {
    colWidths[1].wch = Math.max(colWidths[1].wch, Math.min(maxMetaValue + 2, 80));
  }
  ws["!cols"] = colWidths;

  // --- Row heights ---
  const rowHeights = [];
  // Meta rows (0-4): auto height based on value length
  meta.forEach((m, ri) => {
    const colW = colWidths[1] ? colWidths[1].wch : 30;
    rowHeights[ri] = { hpt: autoHeight(m.value, colW, 18) };
  });
  // Spacer rows (5-6)
  rowHeights[5] = { hpt: 18 };
  rowHeights[6] = { hpt: 18 };
  // Header row (7)
  rowHeights[headerRow] = { hpt: 22 };
  // Data rows (8+)
  pocRows.forEach((row, ri) => {
    const excelRow = headerRow + 1 + ri;
    let maxCellLen = 0;
    DATA_COLUMNS.forEach((col) => {
      const v = str(row[col.key]);
      if (v.length > maxCellLen) maxCellLen = v.length;
    });
    const colW = 30; // average column width for estimation
    rowHeights[excelRow] = { hpt: Math.min(60, Math.max(16, Math.ceil(maxCellLen / colW) * 15)) };
  });
  ws["!rows"] = rowHeights;

  // --- Sheet name (max 31 chars for Excel) ---
  let sheetName = pocName.length > 31 ? pocName.slice(0, 31) : pocName;
  // Remove invalid Excel sheet-name characters
  sheetName = sheetName.replace(/[:\\/?*[\]]/g, "_");

  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

/* ═══════════════════ MAIN EXPORT ═══════════════════ */
export async function exportBurnsheetExcel(allData, region) {
  const wb = XLSX.utils.book_new();

  // 1) Legend sheet (always first)
  await buildLegendSheet(wb);

  // 2) Filter data by region if provided
  const data = region ? allData.filter((r) => r.country === region) : allData;

  // 3) Group by POC
  const pocMap = new Map();
  data.forEach((row) => {
    const poc = str(row.poc) || "Unknown";
    if (!pocMap.has(poc)) pocMap.set(poc, []);
    pocMap.get(poc).push(row);
  });

  // 4) One sheet per POC
  const usedNames = new Set(["Legend"]);
  for (const [pocName, pocRows] of pocMap) {
    // Deduplicate sheet names
    let name = pocName.length > 31 ? pocName.slice(0, 31) : pocName;
    name = name.replace(/[:\\/?*[\]]/g, "_");
    let baseName = name;
    let counter = 2;
    while (usedNames.has(name)) {
      const suffix = ` (${counter})`;
      name = baseName.slice(0, 31 - suffix.length) + suffix;
      counter++;
    }
    usedNames.add(name);

    buildPocSheet(wb, name, pocRows);
  }

  // 5) Write and trigger download
  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbOut], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Burnsheet_${region || "All"}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}
