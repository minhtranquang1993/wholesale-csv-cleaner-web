/**
 * Lưới verify cho pivot: so kết quả với baseline đã đo tay từ file export thật.
 *
 * Chạy:  node scripts/verify-pivot.ts [đường-dẫn-file.xlsx]
 * Mặc định tìm file ở ~/Downloads/export_data_*.xlsx
 *
 * Node ≥22 tự strip type annotation nên không cần thêm dependency nào.
 */

import { readFileSync, readdirSync, existsSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import ExcelJS from "exceljs";
import {
  detectHeaderRow,
  mergeTables,
  readRawXlsx,
  resolveHeaderIndex,
  runPivot,
  tableFromMatrix,
  type ParsedTable,
} from "../app/lib/pivot.ts";

// ─── Baseline đã đo tay bằng cách parse trực tiếp XML của file mẫu ───
const BASELINE = {
  headerRow: 7, // 1-based
  dataRows: 2207,
  columnCount: 44,
  uniquePatients: 156,
  patientTotal: 478735371,
  uniqueSources: 20,
  sourceGrandTotal: 478735371,
  // Gộp file 2 lần. KHÔNG phải 2 x 478735371 = 957470742: buildByPatient làm
  // tròn sau khi cộng theo từng bệnh nhân, 24/156 bệnh nhân có tổng lẻ.
  mergedRows: 4414,
  mergedPatientTotal: 957470743,
  mergedSourceGrandTotal: 957470743,
};

let failures = 0;
let checks = 0;

function check(name: string, actual: unknown, expected: unknown) {
  checks++;
  const ok = actual === expected;
  if (!ok) failures++;
  const fmt = (v: unknown) =>
    typeof v === "number" ? v.toLocaleString("en-US") : String(v);
  console.log(
    `  ${ok ? "✓" : "✗"} ${name}: ${fmt(actual)}${
      ok ? "" : `  (mong đợi ${fmt(expected)})`
    }`
  );
}

function findSampleFile(): string | null {
  const fromArgv = process.argv[2];
  if (fromArgv) return existsSync(fromArgv) ? fromArgv : null;
  const dir = join(homedir(), "Downloads");
  if (!existsSync(dir)) return null;
  const hit = readdirSync(dir).find(
    (f) => f.startsWith("export_data_") && f.endsWith(".xlsx")
  );
  return hit ? join(dir, hit) : null;
}

/** Đọc lại tổng tiền + số dòng từ file XLSX mà runPivot vừa tạo. */
async function readPivotOutput(buffer: ArrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheetNames = wb.worksheets.map((ws) => ws.name);
  const ws = wb.worksheets[0];
  const lastRow = ws.getRow(ws.rowCount);
  const totalCell = lastRow.getCell(ws.columns.length);
  const headerValues = ws
    .getRow(1)
    .values as unknown[];
  return {
    sheetNames,
    total: Number(totalCell.value ?? 0),
    totalLabel: String(lastRow.getCell(ws.columns.length - 1).value ?? ""),
    // numFmt của ô tiền ở dòng dữ liệu đầu tiên (đây mới là ô cần đúng format).
    moneyNumFmt: ws.getCell(2, ws.columns.length).numFmt,
    // exceljs đọc lại view kèm đủ field default nên chỉ so 2 field mình set.
    frozen: `${ws.views?.[0]?.state}/${
      (ws.views?.[0] as { ySplit?: number } | undefined)?.ySplit
    }`,
    headers: headerValues
      .filter((v) => v !== undefined && v !== null)
      .join("|"),
  };
}

async function main() {
  const path = findSampleFile();
  if (!path) {
    console.error(
      "Không tìm thấy file mẫu. Truyền đường dẫn: node scripts/verify-pivot.ts <file.xlsx>"
    );
    process.exit(2);
  }
  console.log(`File mẫu: ${path}\n`);

  const bytes = readFileSync(path);
  const buf = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength
  ) as ArrayBuffer;

  // ── 1. Header detection ──
  console.log("1. Header detection");
  const matrix = await readRawXlsx(buf);
  const detected = detectHeaderRow(matrix);
  check("dòng header (1-based)", detected.headerIndex + 1, BASELINE.headerRow);
  check("cần nhập header thủ công", detected.needsManualHeader, false);

  const table = tableFromMatrix(matrix, detected.headerIndex);
  check("số cột", table.columns.length, BASELINE.columnCount);
  check("số dòng dữ liệu", table.rows.length, BASELINE.dataRows);
  check("cột trùng tên bị bỏ", (table.droppedColumns || []).length, 0);

  // ── 2. Pivot theo bệnh nhân ──
  console.log("\n2. Pivot mode 'patient'");
  const p1 = await runPivot(table, "patient");
  check("số bệnh nhân unique", p1.stats.unique_patients, BASELINE.uniquePatients);
  const out1 = await readPivotOutput(p1.buffer);
  check("TỔNG CỘNG", out1.total, BASELINE.patientTotal);
  check("nhãn dòng tổng", out1.totalLabel, "TỔNG CỘNG");
  check("sheet", out1.sheetNames.join("|"), "Pivot");
  check("money numFmt", out1.moneyNumFmt, "#,##0");
  check("freeze pane", out1.frozen, "frozen/1");
  check(
    "header labels",
    out1.headers,
    "Tên bệnh nhân|Tổng tiền sau miễn giảm"
  );

  // ── 3. Pivot theo nguồn ──
  console.log("\n3. Pivot mode 'source'");
  const p2 = await runPivot(table, "source");
  check("số bệnh nhân unique", p2.stats.unique_patients, BASELINE.uniquePatients);
  check("số nguồn", p2.stats.unique_sources, BASELINE.uniqueSources);
  const out2 = await readPivotOutput(p2.buffer);
  check("GRAND TOTAL", out2.total, BASELINE.sourceGrandTotal);
  check(
    "2 sheet",
    out2.sheetNames.join("|"),
    "Pivot Nguồn|Tổng hợp Nguồn"
  );
  check("freeze pane", out2.frozen, "frozen/1");
  check(
    "header labels",
    out2.headers,
    "Nguồn|Tên bệnh nhân|Tổng tiền sau miễn giảm"
  );

  // Gate phụ: 2 mode khớp nhau trên ĐÚNG input này. Không phải invariant chung —
  // 2 mode làm tròn ở granularity khác nhau (per-bệnh-nhân vs per-(nguồn,bệnh nhân)).
  check("patient total == source total", out1.total, out2.total);

  // ── 4. Gộp 2 file (cùng file add 2 lần) ──
  console.log("\n4. Gộp 2 file");
  const merged = mergeTables([table, table]);
  check("số dòng sau gộp", merged.rows.length, BASELINE.mergedRows);
  check("số cột sau gộp", merged.columns.length, BASELINE.columnCount);
  const m1 = await runPivot(merged, "patient");
  check("số bệnh nhân unique", m1.stats.unique_patients, BASELINE.uniquePatients);
  const mout1 = await readPivotOutput(m1.buffer);
  check("TỔNG CỘNG", mout1.total, BASELINE.mergedPatientTotal);
  const m2 = await runPivot(merged, "source");
  const mout2 = await readPivotOutput(m2.buffer);
  check("GRAND TOTAL", mout2.total, BASELINE.mergedSourceGrandTotal);
  check("patient total == source total", mout1.total, mout2.total);

  // ── 5. Gộp 2 file lệch cách viết tên cột ──
  console.log("\n5. Gộp 2 file lệch cách viết tên cột");
  // Đổi "Tên Bệnh nhân" -> "Tên bệnh nhân" ở bảng thứ 2. mergeTables phải remap
  // key về canonical, nếu không thì rows của bảng 2 trả undefined -> mất tiền.
  const renamed: ParsedTable = {
    columns: table.columns.map((c) =>
      c === "Tên Bệnh nhân" ? "Tên bệnh nhân" : c
    ),
    rows: table.rows.map((r) => {
      const out: Record<string, unknown> = { ...r };
      out["Tên bệnh nhân"] = r["Tên Bệnh nhân"];
      delete out["Tên Bệnh nhân"];
      return out;
    }),
  };
  check(
    "bảng 2 đã đổi tên cột",
    renamed.columns.includes("Tên bệnh nhân"),
    true
  );
  const mergedMixed = mergeTables([table, renamed]);
  check("số cột (không nhân đôi)", mergedMixed.columns.length, BASELINE.columnCount);
  const mm = await runPivot(mergedMixed, "patient");
  check("số bệnh nhân unique", mm.stats.unique_patients, BASELINE.uniquePatients);
  const mmout = await readPivotOutput(mm.buffer);
  check("TỔNG CỘNG (không mất tiền)", mmout.total, BASELINE.mergedPatientTotal);

  // ── 6. Header ở dòng 1 (file đã sạch) không regress ──
  console.log("\n6. File header dòng 1 (không regress)");
  const cleanMatrix = matrix.slice(detected.headerIndex);
  const cleanDetect = detectHeaderRow(cleanMatrix);
  check("dòng header", cleanDetect.headerIndex + 1, 1);
  const cleanTable = tableFromMatrix(cleanMatrix, cleanDetect.headerIndex);
  check("số dòng dữ liệu", cleanTable.rows.length, BASELINE.dataRows);
  const cp = await runPivot(cleanTable, "patient");
  check("số bệnh nhân unique", cp.stats.unique_patients, BASELINE.uniquePatients);
  const cpout = await readPivotOutput(cp.buffer);
  check("TỔNG CỘNG", cpout.total, BASELINE.patientTotal);

  // ── 7. Không detect được header ──
  console.log("\n7. File không có cột anchor nào");
  const junk = [
    ["Tiêu đề báo cáo", null, null],
    ["a", "b", "c"],
    ["1", "2", "3"],
  ];
  const junkDetect = detectHeaderRow(junk);
  check("cần nhập header thủ công", junkDetect.needsManualHeader, true);
  check("anchor score", junkDetect.anchorScore, 0);

  // ── 8. resolveHeaderIndex clamp ──
  console.log("\n8. Override dòng header ngoài range");
  check("nhập 0", resolveHeaderIndex(matrix, 0, 6), 0);
  check("nhập 99999", resolveHeaderIndex(matrix, 99999, 6), matrix.length - 1);
  check("nhập null -> fallback", resolveHeaderIndex(matrix, null, 6), 6);
  check("nhập NaN -> fallback", resolveHeaderIndex(matrix, NaN, 6), 6);
  check("nhập 7 -> index 6", resolveHeaderIndex(matrix, 7, 0), 6);
  check(
    "matrix rỗng không crash",
    resolveHeaderIndex([], 7, 6),
    0
  );

  console.log(
    `\n${failures === 0 ? "✅ PASS" : "❌ FAIL"} — ${checks - failures}/${checks} check`
  );
  process.exit(failures === 0 ? 0 : 1);
}

main().catch((err) => {
  console.error("Lỗi khi chạy verify:", err);
  process.exit(2);
});
