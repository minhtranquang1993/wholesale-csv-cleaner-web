import { NextRequest, NextResponse } from "next/server";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const BASE_PREFIX = "https://qhdistribution.com/wholesale/?";

function findColumn(
  columns: string[],
  candidates: string[]
): string | null {
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

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;

    if (!file) {
      return NextResponse.json({ error: "Chưa chọn file" }, { status: 400 });
    }

    if (!file.name.toLowerCase().endsWith(".csv")) {
      return NextResponse.json(
        { error: "Chỉ hỗ trợ file .csv" },
        { status: 400 }
      );
    }

    const text = await file.text();
    const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });

    if (parsed.errors.length > 0 && parsed.data.length === 0) {
      return NextResponse.json(
        { error: "Không thể đọc file CSV" },
        { status: 400 }
      );
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
      return NextResponse.json(
        { error: "Không tìm thấy cột 'first_source_url' trong file CSV." },
        { status: 400 }
      );
    }
    if (!urlCol) {
      return NextResponse.json(
        { error: "Không tìm thấy cột 'url' trong file CSV." },
        { status: 400 }
      );
    }

    // Sort by time if possible
    const timeCol = findColumn(columns, [
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
    ]);

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

    // Create XLSX
    const ws = XLSX.utils.json_to_sheet(rows, { header: columns });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const xlsxBuffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    const stats = JSON.stringify({
      first_source_url: countSourceUrl,
      first_user_source: countUserSource,
      first_user_medium: countUserMedium,
    });

    return new NextResponse(xlsxBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="wholesale_cleaned.xlsx"',
        "X-Clean-Stats": stats,
      },
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Lỗi không xác định";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
