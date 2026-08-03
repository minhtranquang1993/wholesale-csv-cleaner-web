import Papa from "papaparse";
import * as XLSX from "xlsx";

// Port nguyên logic từ app/api/clean/route.ts để chạy trên trình duyệt:
// không upload file lên server nên không dính giới hạn body 4.5MB của Vercel,
// và đo được tiến độ thật theo từng phase.

const BASE_PREFIX = "https://qhdistribution.com/wholesale/?";

const TIME_COL_CANDIDATES = [
  "submission_date",
  "submission date",
  "created_at",
  "created at",
  "timestamp",
  "time",
  "date",
  "datetime",
  "date_time",
  "date time",
  "created_time",
  "created time",
];

function findColumn(columns: string[], candidates: string[]): string | null {
  const lowered: Record<string, string> = {};
  for (const c of columns) {
    lowered[c.toLowerCase()] = c;
  }
  for (const candidate of candidates) {
    if (candidate.toLowerCase() in lowered) {
      return lowered[candidate.toLowerCase()];
    }
  }
  return null;
}

function isEmpty(value: unknown): boolean {
  if (value === null || value === undefined) return true;
  const s = String(value).trim();
  return s === "" || s.toLowerCase() === "nan";
}

function extractUtm(urlValue: unknown, paramName: string): string | null {
  try {
    const raw = String(urlValue).trim();
    if (!raw || raw.toLowerCase() === "nan") return null;

    let fullUrl = raw;
    if (!raw.startsWith("http")) {
      fullUrl = "https://x.com/?" + raw;
    }
    const parsed = new URL(fullUrl);
    const value = parsed.searchParams.get(paramName);
    return value ? value.trim() : null;
  } catch {
    return null;
  }
}

export interface CleanStats {
  first_source_url: number;
  first_user_source: number;
  first_user_medium: number;
}

export interface CleanOutput {
  buffer: ArrayBuffer;
  stats: CleanStats;
  rowCount: number;
}

export interface CleanPhase {
  /** 0-100 */
  percent: number;
  label: string;
  /** Phase blocking không chia nhỏ được → bar chạy sọc thay vì %. */
  indeterminate?: boolean;
}

/** Nhường main thread để React repaint được progress bar. */
function yieldToPaint(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * Làm sạch CSV wholesale. Band % theo cost profile đo thật trên CSV 20k dòng:
 * parse 4.9% · sort 5.6% · rule 5.4% · json_to_sheet 17.8% · XLSX.write 66.4%.
 */
export async function cleanWholesaleCsv(
  file: File,
  onPhase?: (phase: CleanPhase) => void
): Promise<CleanOutput> {
  const report = async (phase: CleanPhase) => {
    onPhase?.(phase);
    await yieldToPaint();
  };

  if (!file.name.toLowerCase().endsWith(".csv")) {
    throw new Error("Chỉ hỗ trợ file .csv");
  }

  await report({ percent: 0, label: "Đang đọc file..." });
  const text = await file.text();

  await report({ percent: 10, label: "Đang phân tích CSV..." });
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });

  if (parsed.errors.length > 0 && parsed.data.length === 0) {
    throw new Error("Không thể đọc file CSV");
  }

  const rows = parsed.data as Record<string, unknown>[];
  const columns = parsed.meta.fields || [];

  const firstSourceCol = findColumn(columns, [
    "first_source_url",
    "first source url",
  ]);
  const urlCol = findColumn(columns, ["url"]);
  const firstUserSourceCol = findColumn(columns, [
    "first_user_source",
    "first user source",
  ]);
  const firstUserMediumCol = findColumn(columns, [
    "first_user_medium",
    "first user medium",
  ]);

  if (!firstSourceCol) {
    throw new Error("Không tìm thấy cột 'first_source_url' trong file CSV.");
  }
  if (!urlCol) {
    throw new Error("Không tìm thấy cột 'url' trong file CSV.");
  }

  await report({ percent: 20, label: "Đang sắp xếp theo thời gian..." });
  const timeCol = findColumn(columns, TIME_COL_CANDIDATES);
  if (timeCol) {
    rows.sort((a, b) => {
      const da = new Date(String(a[timeCol] || "")).getTime();
      const db = new Date(String(b[timeCol] || "")).getTime();
      if (isNaN(da) && isNaN(db)) return 0;
      if (isNaN(da)) return 1;
      if (isNaN(db)) return -1;
      return da - db;
    });
  }

  await report({
    percent: 25,
    label: `Đang áp dụng 3 rule cho ${rows.length.toLocaleString("vi-VN")} dòng...`,
  });

  let countSourceUrl = 0;
  let countUserSource = 0;
  let countUserMedium = 0;

  for (const row of rows) {
    const sourceEmpty = isEmpty(row[firstSourceCol]);
    const urlNotEmpty = !isEmpty(row[urlCol]);

    if (sourceEmpty && urlNotEmpty) {
      // Rule 1
      row[firstSourceCol] = BASE_PREFIX + String(row[urlCol]).trim();
      countSourceUrl++;

      // Rule 2
      if (firstUserSourceCol) {
        const utm = extractUtm(row[urlCol], "utm_source");
        if (utm) {
          row[firstUserSourceCol] = utm;
          countUserSource++;
        }
      }

      // Rule 3
      if (firstUserMediumCol) {
        const utm = extractUtm(row[urlCol], "utm_medium");
        if (utm) {
          row[firstUserMediumCol] = utm;
          countUserMedium++;
        }
      }
    }
  }

  await report({ percent: 35, label: "Đang dựng bảng tính..." });
  const ws = XLSX.utils.json_to_sheet(rows, { header: columns });

  // XLSX.write chiếm ~66% thời gian và là call đồng bộ không có progress
  // callback → báo indeterminate trước khi vào call thay vì % giả.
  await report({
    percent: 50,
    label: "Đang ghi file XLSX (có thể mất vài giây)...",
    indeterminate: true,
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  const written = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const buffer = written as ArrayBuffer;

  await report({ percent: 100, label: "Hoàn tất" });

  return {
    buffer,
    rowCount: rows.length,
    stats: {
      first_source_url: countSourceUrl,
      first_user_source: countUserSource,
      first_user_medium: countUserMedium,
    },
  };
}
