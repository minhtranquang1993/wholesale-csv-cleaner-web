import { NextRequest, NextResponse } from "next/server";
import { parseFile, runPivot, type PivotMode } from "@/app/lib/pivot";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    const mode = String(formData.get("mode") || "patient") as PivotMode;

    if (!file) {
      return NextResponse.json({ error: "Chưa chọn file" }, { status: 400 });
    }

    const table = await parseFile(file.name, {
      text: () => file.text(),
      arrayBuffer: () => file.arrayBuffer(),
    });

    const { buffer, stats, filename } = await runPivot(table, mode);

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "X-Pivot-Stats": JSON.stringify(stats),
      },
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Lỗi không xác định";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
