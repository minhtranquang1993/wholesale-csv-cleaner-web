import Papa from "papaparse";
import ExcelJS from "exceljs";

export const COL_TEN_CANDIDATES = ["Tên Bệnh nhân", "Tên bệnh nhân"];
export const COL_NGUON_CANDIDATES = ["Nguồn"];
export const COL_TIEN_CANDIDATES = ["Tiền sau miễn giảm"];

const NO_SOURCE = "(Không rõ nguồn)";

const GROUP_COLORS = [
  "EBF5FB",
  "FEF9E7",
  "F5EEF8",
  "E8F8F5",
  "FDEDEC",
  "F0F3F4",
];

const collator = new Intl.Collator("vi", { sensitivity: "base", numeric: true });

// ─── Helpers ───

/** Normalize a column-name / key for flexible matching. */
export function normKey(s: unknown): string {
  return String(s ?? "")
    .normalize("NFC")
    .replace(/^Range\[(.+)\]$/, "$1")
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("vi");
}

/** Find the actual column name in `columns` matching any candidate. */
export function findColumn(
  columns: string[],
  candidates: string[]
): string | null {
  const map = new Map<string, string>();
  for (const c of columns) map.set(normKey(c), c);
  for (const cand of candidates) {
    const hit = map.get(normKey(cand));
    if (hit !== undefined) return hit;
  }
  return null;
}

/** Parse a money value handling VN formats ("1.000.000", "1.000.000,50") and numeric cells. */
export function parseMoney(value: unknown): number {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return isFinite(value) ? value : 0;
  if (typeof value === "object") {
    // exceljs formula/result cell
    const obj = value as { result?: unknown };
    if ("result" in obj) return parseMoney(obj.result);
    return 0;
  }
  let s = String(value).trim();
  if (!s || s.toLowerCase() === "nan") return 0;
  // keep only digits, separators and sign
  s = s.replace(/[^\d.,-]/g, "");
  if (!s || s === "-") return 0;

  const hasDot = s.includes(".");
  const hasComma = s.includes(",");

  if (hasDot && hasComma) {
    // last separator = decimal
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) {
      s = s.replace(/\./g, "").replace(",", "."); // VN: dot thousands, comma decimal
    } else {
      s = s.replace(/,/g, ""); // EN: comma thousands, dot decimal
    }
  } else if (hasComma) {
    // only commas → thousands separators if grouped in 3s, else decimal
    const parts = s.split(",");
    const groupedThousands =
      parts.length > 2 || (parts.length === 2 && parts[1].length === 3);
    if (groupedThousands) s = s.replace(/,/g, "");
    else s = s.replace(",", ".");
  } else if (hasDot) {
    // only dots → thousands separators if grouped in 3s, else decimal
    const parts = s.split(".");
    const groupedThousands =
      parts.length > 2 || (parts.length === 2 && parts[1].length === 3);
    if (groupedThousands) s = s.replace(/\./g, "");
  }

  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

export function isEmptyVal(value: unknown): boolean {
  if (value === null || value === undefined) return true;
  const s = String(value).trim();
  return s === "" || s.toLowerCase() === "nan";
}

export interface ParsedTable {
  columns: string[];
  rows: Record<string, unknown>[];
  /** Cột bị bỏ vì trùng normKey với một cột trước đó (giữ cột đầu tiên). */
  droppedColumns?: string[];
}

/** Ma trận thô: index 0 = dòng 1 trong Excel. Dòng rỗng được giữ để index luôn khớp. */
export type RawMatrix = unknown[][];

/** Số dòng đầu file được quét để tìm dòng header. */
export const HEADER_SCAN_ROWS = 30;

/** Gỡ các dạng cell của exceljs (formula / hyperlink / rich text) về giá trị thô. */
function unwrapCell(value: unknown): unknown {
  if (value === null || value === undefined) return null;
  if (typeof value !== "object") return value;
  if (value instanceof Date) return value;
  const o = value as {
    result?: unknown;
    richText?: { text?: string }[];
    text?: unknown;
    hyperlink?: unknown;
  };
  if ("result" in o) return unwrapCell(o.result);
  if (Array.isArray(o.richText)) {
    return o.richText.map((part) => part?.text ?? "").join("");
  }
  if ("text" in o) return o.text;
  return null;
}

/** Nhường main thread để React repaint được progress bar. */
export function yieldToPaint(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

const YIELD_EVERY_ROWS = 500;

export async function readRawXlsx(
  buffer: ArrayBuffer,
  onProgress?: (ratio: number) => void
): Promise<RawMatrix> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) return [];

  const width = ws.columnCount;
  const total = ws.rowCount;
  const matrix: RawMatrix = [];
  // Dùng eachCell (bỏ ô rỗng) thay vì getCell: getCell materialize cả ô trống nên
  // chậm ~42x trên file thật (5408ms vs 127ms) mà ma trận thu được y hệt.
  // Duyệt tới rowCount (không phải actualRowCount) để index tuyệt đối luôn khớp
  // số dòng người dùng thấy trong Excel khi họ nhập dòng header thủ công.
  for (let r = 1; r <= total; r++) {
    const arr: unknown[] = new Array(width).fill(null);
    ws.getRow(r).eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (colNumber <= width) arr[colNumber - 1] = unwrapCell(cell.value);
    });
    matrix.push(arr);
    if (onProgress && r % YIELD_EVERY_ROWS === 0) {
      onProgress(r / total);
      await yieldToPaint();
    }
  }
  onProgress?.(1);
  return matrix;
}

export function readRawCsv(text: string): RawMatrix {
  // header:false + skipEmptyLines:false để index dòng khớp tuyệt đối với file gốc.
  const parsed = Papa.parse<unknown[]>(text, {
    header: false,
    skipEmptyLines: false,
  });
  return (parsed.data as unknown[][]) || [];
}

/**
 * Tìm dòng header trong `HEADER_SCAN_ROWS` dòng đầu.
 * Chấm điểm 2 tầng: (1) số cột anchor khớp, (2) số giá trị non-empty phân biệt.
 * Tầng 2 cần thiết vì mỗi anchor chỉ có ít cách viết được chấp nhận — file đổi
 * tên cột một chút là mất anchor. Dòng metadata dùng merged cell nên chỉ còn
 * 1-2 giá trị phân biệt, còn dòng header thật có hàng chục.
 */
export function detectHeaderRow(matrix: RawMatrix): {
  headerIndex: number;
  anchorScore: number;
  needsManualHeader: boolean;
} {
  const anchors = [
    COL_TEN_CANDIDATES,
    COL_TIEN_CANDIDATES,
    COL_NGUON_CANDIDATES,
  ];
  const limit = Math.min(matrix.length, HEADER_SCAN_ROWS);

  let best = { headerIndex: 0, anchorScore: -1, distinct: -1 };
  for (let i = 0; i < limit; i++) {
    const row = matrix[i] || [];
    const keys = new Set<string>();
    const values = new Set<string>();
    for (const cell of row) {
      if (isEmptyVal(cell)) continue;
      keys.add(normKey(cell));
      values.add(String(cell).trim());
    }
    const anchorScore = anchors.reduce(
      (acc, cands) =>
        acc + (cands.some((c) => keys.has(normKey(c))) ? 1 : 0),
      0
    );
    const distinct = values.size;
    if (
      anchorScore > best.anchorScore ||
      (anchorScore === best.anchorScore && distinct > best.distinct)
    ) {
      best = { headerIndex: i, anchorScore, distinct };
    }
  }

  return {
    headerIndex: Math.max(0, best.headerIndex),
    anchorScore: Math.max(0, best.anchorScore),
    // Không có anchor nào khớp → không đoán bừa dòng 1 (sẽ ra lỗi "cột hiện có:
    // BỆNH VIỆN..."). Báo UI bắt người dùng nhập dòng header.
    needsManualHeader: best.anchorScore <= 0,
  };
}

/** Cắt ma trận thô thành bảng, lấy `headerIndex` (0-based) làm dòng header. */
export function tableFromMatrix(
  matrix: RawMatrix,
  headerIndex: number
): ParsedTable {
  const headerRow = matrix[headerIndex];
  if (!headerRow) return { columns: [], rows: [], droppedColumns: [] };

  const colDefs: { index: number; name: string }[] = [];
  const seen = new Set<string>();
  const droppedColumns: string[] = [];
  headerRow.forEach((cell, index) => {
    if (isEmptyVal(cell)) return;
    const name = String(cell).trim();
    const key = normKey(name);
    if (seen.has(key)) {
      droppedColumns.push(name);
      return;
    }
    seen.add(key);
    colDefs.push({ index, name });
  });

  const rows: Record<string, unknown>[] = [];
  for (let r = headerIndex + 1; r < matrix.length; r++) {
    const raw = matrix[r] || [];
    const obj: Record<string, unknown> = {};
    let hasAny = false;
    for (const { index, name } of colDefs) {
      const v = raw[index] ?? null;
      obj[name] = v;
      if (!isEmptyVal(v)) hasAny = true;
    }
    if (hasAny) rows.push(obj);
  }

  return { columns: colDefs.map((c) => c.name), rows, droppedColumns };
}

/**
 * Gộp nhiều bảng thành một. Cột được union theo `normKey`, lấy cách viết của
 * bảng đầu tiên làm canonical và **remap lại key của từng row**.
 * Bắt buộc remap: nếu file A ghi "Tên Bệnh nhân" và file B ghi "Tên bệnh nhân",
 * gộp thô sẽ khiến findColumn chọn một tên, rows của file kia trả undefined →
 * cộng thành 0 mà không báo lỗi gì.
 */
export function mergeTables(tables: ParsedTable[]): ParsedTable {
  const canonical = new Map<string, string>();
  const columns: string[] = [];
  for (const t of tables) {
    for (const col of t.columns) {
      const key = normKey(col);
      if (!canonical.has(key)) {
        canonical.set(key, col);
        columns.push(col);
      }
    }
  }

  const rows: Record<string, unknown>[] = [];
  for (const t of tables) {
    // Map tên cột của bảng này → tên canonical, tính 1 lần cho cả bảng.
    const remap: { from: string; to: string }[] = t.columns.map((col) => ({
      from: col,
      to: canonical.get(normKey(col)) || col,
    }));
    for (const row of t.rows) {
      const out: Record<string, unknown> = {};
      for (const { from, to } of remap) out[to] = row[from];
      rows.push(out);
    }
  }

  return { columns, rows };
}


// ─── Excel styling helpers ───

const HEADER_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF1F4E79" },
};
const HEADER_FONT: Partial<ExcelJS.Font> = {
  bold: true,
  color: { argb: "FFFFFFFF" },
  name: "Segoe UI",
  size: 11,
};
const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: "thin", color: { argb: "FFD5D8DC" } },
  left: { style: "thin", color: { argb: "FFD5D8DC" } },
  bottom: { style: "thin", color: { argb: "FFD5D8DC" } },
  right: { style: "thin", color: { argb: "FFD5D8DC" } },
};
const MONEY_FMT = "#,##0";

function solid(color: string): ExcelJS.Fill {
  return { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + color } };
}

function styleHeader(ws: ExcelJS.Worksheet) {
  ws.getRow(1).eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
}

// ─── Pivot builders ───

export function buildByPatient(
  rows: Record<string, unknown>[],
  colTen: string,
  colTien: string
): Promise<{ buffer: ArrayBuffer; uniquePatients: number }> {
  const sums = new Map<string, number>();
  for (const row of rows) {
    const name = String(row[colTen] ?? "").trim();
    if (!name) continue;
    sums.set(name, (sums.get(name) || 0) + parseMoney(row[colTien]));
  }

  const names = Array.from(sums.keys()).sort((a, b) => collator.compare(a, b));
  // Làm tròn từng dòng trước, rồi cộng các số ĐÃ làm tròn để dòng TỔNG CỘNG
  // luôn đúng bằng tổng các dòng hiển thị phía trên.
  const rounded = new Map(
    names.map((name) => [name, Math.round(sums.get(name) || 0)] as const)
  );
  const total = names.reduce((acc, name) => acc + (rounded.get(name) || 0), 0);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Pivot");
  ws.columns = [
    { header: "Tên bệnh nhân", key: "ten", width: 30 },
    { header: "Tổng tiền sau miễn giảm", key: "tien", width: 25 },
  ];
  styleHeader(ws);

  for (const name of names) {
    ws.addRow({ ten: name, tien: rounded.get(name) || 0 });
  }
  const totalRow = ws.addRow({ ten: "TỔNG CỘNG", tien: total });

  const lastRow = ws.rowCount;
  for (let r = 2; r <= lastRow; r++) {
    const cellB = ws.getCell(r, 2);
    cellB.numFmt = MONEY_FMT;
    cellB.alignment = { horizontal: "right" };
  }
  totalRow.eachCell((cell) => {
    cell.fill = solid("D9E1F2");
    cell.font = { bold: true };
  });
  ws.views = [{ state: "frozen", ySplit: 1 }];

  return wb.xlsx.writeBuffer().then((buffer) => ({
    buffer,
    uniquePatients: names.length,
  }));
}

export function buildBySource(
  rows: Record<string, unknown>[],
  colNguon: string,
  colTen: string,
  colTien: string
): Promise<{
  buffer: ArrayBuffer;
  uniquePatients: number;
  uniqueSources: number;
}> {
  // group by source -> patient -> sum
  const grouped = new Map<string, Map<string, number>>();
  for (const row of rows) {
    const name = String(row[colTen] ?? "").trim();
    if (!name) continue;
    let source = String(row[colNguon] ?? "").trim();
    if (!source || source.toLowerCase() === "nan") source = NO_SOURCE;
    if (!grouped.has(source)) grouped.set(source, new Map());
    const inner = grouped.get(source)!;
    inner.set(name, (inner.get(name) || 0) + parseMoney(row[colTien]));
  }

  const sources = Array.from(grouped.keys()).sort((a, b) =>
    collator.compare(a, b)
  );

  const wb = new ExcelJS.Workbook();

  // ── Sheet 1: Pivot Nguồn (detail with subtotals) ──
  const ws = wb.addWorksheet("Pivot Nguồn");
  ws.columns = [
    { header: "Nguồn", key: "nguon", width: 22 },
    { header: "Tên bệnh nhân", key: "ten", width: 35 },
    { header: "Tổng tiền sau miễn giảm", key: "tien", width: 28 },
  ];
  styleHeader(ws);

  const patientSet = new Set<string>();
  let grandTotal = 0;
  // Subtotal đã làm tròn của từng nguồn, dùng lại cho sheet 2 để hai sheet
  // luôn khớp nhau.
  const sourceSubtotals = new Map<string, number>();

  sources.forEach((source, sourceIdx) => {
    const inner = grouped.get(source)!;
    const names = Array.from(inner.keys()).sort((a, b) =>
      collator.compare(a, b)
    );
    let subtotal = 0;
    const color = GROUP_COLORS[sourceIdx % GROUP_COLORS.length];

    for (const name of names) {
      patientSet.add(name);
      const amount = Math.round(inner.get(name) || 0);
      subtotal += amount;
      const row = ws.addRow({ nguon: source, ten: name, tien: amount });
      row.eachCell((cell) => {
        cell.fill = solid(color);
        cell.font = { name: "Segoe UI", size: 10 };
        cell.border = THIN_BORDER;
      });
    }

    grandTotal += subtotal;
    sourceSubtotals.set(source, subtotal);
    const subRow = ws.addRow({
      nguon: "",
      ten: `Subtotal [${source}]`,
      tien: subtotal,
    });
    subRow.eachCell((cell) => {
      cell.fill = solid("D6EAF8");
      cell.font = {
        bold: true,
        name: "Segoe UI",
        size: 10,
        color: { argb: "FF1A5276" },
      };
      cell.border = THIN_BORDER;
    });
  });

  const grandRow = ws.addRow({
    nguon: "",
    ten: "TỔNG CỘNG",
    tien: grandTotal,
  });
  grandRow.eachCell((cell) => {
    cell.fill = solid("1F4E79");
    cell.font = {
      bold: true,
      name: "Segoe UI",
      size: 11,
      color: { argb: "FFFFFFFF" },
    };
    cell.border = THIN_BORDER;
  });

  // money format col C
  for (let r = 2; r <= ws.rowCount; r++) {
    const cellC = ws.getCell(r, 3);
    cellC.numFmt = MONEY_FMT;
    cellC.alignment = { horizontal: "right" };
  }
  ws.views = [{ state: "frozen", ySplit: 1 }];

  // ── Sheet 2: Tổng hợp Nguồn (summary per source) ──
  const ws2 = wb.addWorksheet("Tổng hợp Nguồn");
  ws2.columns = [
    { header: "Nguồn", key: "nguon", width: 28 },
    { header: "Tổng tiền sau miễn giảm", key: "tien", width: 28 },
  ];
  styleHeader(ws2);

  sources.forEach((source, idx) => {
    const sourceTotal = sourceSubtotals.get(source) || 0;
    const row = ws2.addRow({ nguon: source, tien: sourceTotal });
    const color = GROUP_COLORS[idx % GROUP_COLORS.length];
    row.eachCell((cell) => {
      cell.fill = solid(color);
      cell.font = { name: "Segoe UI", size: 10 };
      cell.border = THIN_BORDER;
    });
  });
  const totalRow2 = ws2.addRow({ nguon: "TỔNG CỘNG", tien: grandTotal });
  totalRow2.eachCell((cell) => {
    cell.fill = solid("1F4E79");
    cell.font = {
      bold: true,
      name: "Segoe UI",
      size: 11,
      color: { argb: "FFFFFFFF" },
    };
    cell.border = THIN_BORDER;
  });
  for (let r = 2; r <= ws2.rowCount; r++) {
    const cellB = ws2.getCell(r, 2);
    cellB.numFmt = MONEY_FMT;
    cellB.alignment = { horizontal: "right" };
  }
  ws2.views = [{ state: "frozen", ySplit: 1 }];

  return wb.xlsx.writeBuffer().then((buffer) => ({
    buffer,
    uniquePatients: patientSet.size,
    uniqueSources: sources.length,
  }));
}

// ─── Orchestrator (shared by browser + API route) ───

export type PivotMode = "patient" | "source";

export interface PivotOutput {
  buffer: ArrayBuffer;
  stats: { unique_patients: number; unique_sources?: number };
  filename: string;
}

/** Đọc file (xlsx/csv) thành ma trận thô. Throw message hiển thị cho user nếu định dạng sai. */
export async function readRawFile(
  fileName: string,
  read: {
    text: () => Promise<string>;
    arrayBuffer: () => Promise<ArrayBuffer>;
  },
  onProgress?: (ratio: number) => void
): Promise<RawMatrix> {
  const lower = fileName.toLowerCase();
  const isCsv = lower.endsWith(".csv");
  const isXlsx = lower.endsWith(".xlsx");
  if (!isCsv && !isXlsx) {
    throw new Error("Chỉ hỗ trợ file .xlsx hoặc .csv");
  }
  if (isCsv) {
    const matrix = readRawCsv(await read.text());
    onProgress?.(1);
    return matrix;
  }
  return readRawXlsx(await read.arrayBuffer(), onProgress);
}

/**
 * Chuyển 1-based (số dòng người dùng thấy trong Excel) sang index 0-based,
 * clamp vào [0, matrix.length-1]. Giá trị rỗng/không hợp lệ → `fallbackIndex`.
 */
export function resolveHeaderIndex(
  matrix: RawMatrix,
  headerRow1Based: number | null | undefined,
  fallbackIndex: number
): number {
  const maxIndex = Math.max(0, matrix.length - 1);
  if (
    headerRow1Based === null ||
    headerRow1Based === undefined ||
    !Number.isFinite(headerRow1Based)
  ) {
    return Math.min(Math.max(0, fallbackIndex), maxIndex);
  }
  const idx = Math.trunc(headerRow1Based) - 1;
  return Math.min(Math.max(0, idx), maxIndex);
}

/** Parse một file (xlsx/csv) thành bảng, tự nhận dòng header. */
export async function parseFile(
  fileName: string,
  read: {
    text: () => Promise<string>;
    arrayBuffer: () => Promise<ArrayBuffer>;
  }
): Promise<ParsedTable> {
  const matrix = await readRawFile(fileName, read);
  const { headerIndex } = detectHeaderRow(matrix);
  return tableFromMatrix(matrix, headerIndex);
}

/** Các cột bắt buộc theo từng mode, dùng để validate từng file trước khi gộp. */
export function missingColumns(columns: string[], mode: PivotMode): string[] {
  const missing: string[] = [];
  if (!findColumn(columns, COL_TEN_CANDIDATES)) missing.push("Tên Bệnh nhân");
  if (!findColumn(columns, COL_TIEN_CANDIDATES))
    missing.push("Tiền sau miễn giảm");
  if (mode === "source" && !findColumn(columns, COL_NGUON_CANDIDATES))
    missing.push("Nguồn");
  return missing;
}

/** Run the pivot over an already-parsed table. Throws a user-facing message when columns are missing. */
export async function runPivot(
  table: ParsedTable,
  mode: PivotMode
): Promise<PivotOutput> {
  if (table.rows.length === 0) {
    throw new Error("File không có dữ liệu.");
  }

  const colTen = findColumn(table.columns, COL_TEN_CANDIDATES);
  const colTien = findColumn(table.columns, COL_TIEN_CANDIDATES);

  if (!colTen) {
    throw new Error(
      `Không tìm thấy cột 'Tên Bệnh nhân'. Các cột hiện có: ${table.columns.join(
        ", "
      )}`
    );
  }
  if (!colTien) {
    throw new Error(
      `Không tìm thấy cột 'Tiền sau miễn giảm'. Các cột hiện có: ${table.columns.join(
        ", "
      )}`
    );
  }

  if (mode === "source") {
    const colNguon = findColumn(table.columns, COL_NGUON_CANDIDATES);
    if (!colNguon) {
      throw new Error(
        `Không tìm thấy cột 'Nguồn'. Các cột hiện có: ${table.columns.join(
          ", "
        )}`
      );
    }
    const res = await buildBySource(table.rows, colNguon, colTen, colTien);
    return {
      buffer: res.buffer,
      stats: {
        unique_patients: res.uniquePatients,
        unique_sources: res.uniqueSources,
      },
      filename: "pivot_nguon_benh_nhan.xlsx",
    };
  }

  const res = await buildByPatient(table.rows, colTen, colTien);
  return {
    buffer: res.buffer,
    stats: { unique_patients: res.uniquePatients },
    filename: "pivot_benh_nhan.xlsx",
  };
}
