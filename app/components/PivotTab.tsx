"use client";

import { useState, useCallback, useRef } from "react";
import ProgressBar from "./ProgressBar";
import {
  detectHeaderRow,
  mergeTables,
  missingColumns,
  readRawFile,
  resolveHeaderIndex,
  runPivot,
  tableFromMatrix,
  yieldToPaint,
  type PivotMode,
  type RawMatrix,
} from "../lib/pivot";

interface PivotResult {
  unique_patients: number;
  unique_sources?: number;
  total_rows: number;
  file_count: number;
}

/** Một file đã đọc xong: giữ matrix trong RAM để đổi dòng header không cần đọc lại. */
interface LoadedFile {
  id: string;
  name: string;
  matrix: RawMatrix;
  /** 1-based, đúng số dòng người dùng thấy trong Excel. */
  headerRow: number;
  autoHeaderRow: number;
  needsManualHeader: boolean;
  dataRows: number;
  columnCount: number;
  droppedColumns: string[];
}

const MODE_INFO: Record<
  PivotMode,
  { label: string; desc: string; filename: string }
> = {
  patient: {
    label: "Theo bệnh nhân",
    desc: "Gộp theo Tên Bệnh nhân, tính tổng Tiền sau miễn giảm + dòng TỔNG CỘNG.",
    filename: "pivot_benh_nhan.xlsx",
  },
  source: {
    label: "Theo nguồn + bệnh nhân",
    desc: "Gộp theo Nguồn + Tên Bệnh nhân, subtotal mỗi nguồn + sheet tổng hợp nguồn.",
    filename: "pivot_nguon_benh_nhan.xlsx",
  },
};

export default function PivotTab() {
  const [mode, setMode] = useState<PivotMode>("patient");
  const [files, setFiles] = useState<LoadedFile[]>([]);
  const [loading, setLoading] = useState(false);
  const [reading, setReading] = useState(false);
  const [phase, setPhase] = useState<{
    percent: number;
    label: string;
  } | null>(null);
  const [result, setResult] = useState<PivotResult | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const idRef = useRef(0);
  // Gate bằng ref, không dùng state: onDrop là useCallback deps [] nên đọc
  // state sẽ luôn thấy giá trị của lần render đầu.
  const busyRef = useRef(false);

  const busy = loading || reading;

  /**
   * Đổi link tải, tự revoke link cũ. Mọi chỗ đổi downloadUrl phải đi qua đây để
   * blob của file XLSX trước không bị giữ lại trong bộ nhớ.
   */
  const setDownload = (url: string | null) => {
    setDownloadUrl((prev) => {
      if (prev && prev !== url) URL.revokeObjectURL(prev);
      return url;
    });
  };

  /** Dựng lại thông tin bảng của 1 file theo dòng header hiện tại. */
  const describe = (matrix: RawMatrix, headerRow: number) => {
    const idx = resolveHeaderIndex(matrix, headerRow, 0);
    const table = tableFromMatrix(matrix, idx);
    return {
      dataRows: table.rows.length,
      columnCount: table.columns.length,
      droppedColumns: table.droppedColumns || [],
    };
  };

  const addFiles = async (incoming: File[]) => {
    // Không cho 2 lượt đọc chạy song song: lượt xong trước sẽ set reading=false
    // giữa lúc lượt kia còn đang thêm file, khiến nút chạy pivot mở ra sớm và
    // pivot thiếu file (mất trọn số tiền của file đó).
    if (busyRef.current) return;

    const accepted = incoming.filter((f) => {
      const n = f.name.toLowerCase();
      return n.endsWith(".xlsx") || n.endsWith(".csv");
    });
    if (accepted.length !== incoming.length) {
      setError("Chỉ hỗ trợ file .xlsx hoặc .csv");
    } else {
      setError(null);
    }
    if (!accepted.length) return;

    busyRef.current = true;
    setReading(true);
    setResult(null);
    setDownload(null);

    try {
      // Phần lớn thời gian nằm ở đây (~460ms/file): đọc file + dựng matrix.
      // Progress chia đều cho từng file, mỗi file tự báo tiến độ nội bộ.
      for (let i = 0; i < accepted.length; i++) {
        const f = accepted[i];
        const base = (i / accepted.length) * 100;
        const span = 100 / accepted.length;
        setPhase({
          percent: base,
          label: `Đang đọc ${f.name} (${i + 1}/${accepted.length})...`,
        });
        await yieldToPaint();

        const matrix = await readRawFile(
          f.name,
          {
            text: () => f.text(),
            arrayBuffer: () => f.arrayBuffer(),
          },
          (ratio) => setPhase({
            percent: base + ratio * span,
            label: `Đang đọc ${f.name} (${i + 1}/${accepted.length})...`,
          })
        );

        const detected = detectHeaderRow(matrix);
        const headerRow = detected.headerIndex + 1;
        const info = describe(matrix, headerRow);
        idRef.current += 1;
        const loaded: LoadedFile = {
          id: `f${idRef.current}`,
          name: f.name,
          matrix,
          headerRow,
          autoHeaderRow: headerRow,
          needsManualHeader: detected.needsManualHeader,
          ...info,
        };
        setFiles((prev) => [...prev, loaded]);
      }
      setPhase({ percent: 100, label: "Đã đọc xong" });
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : "Không đọc được file");
    } finally {
      busyRef.current = false;
      setReading(false);
      setPhase(null);
    }
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files.length) {
      addFiles(Array.from(e.dataTransfer.files));
    }
  }, []);

  const changeHeaderRow = (id: string, value: string) => {
    setFiles((prev) =>
      prev.map((f) => {
        if (f.id !== id) return f;
        const parsed = parseInt(value, 10);
        const idx = resolveHeaderIndex(
          f.matrix,
          Number.isFinite(parsed) ? parsed : null,
          f.autoHeaderRow - 1
        );
        const headerRow = idx + 1;
        return { ...f, headerRow, ...describe(f.matrix, headerRow) };
      })
    );
    setResult(null);
    setDownload(null);
  };

  const removeFile = (id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
    setResult(null);
    setDownload(null);
    setError(null);
  };

  const handleSubmit = async () => {
    if (!files.length || busyRef.current) return;
    busyRef.current = true;
    setLoading(true);
    setError(null);
    setResult(null);
    setDownload(null);

    try {
      // Matrix đã ở RAM nên phase này rất nhanh (~26ms): chỉ cắt bảng, gộp,
      // pivot rồi ghi file. Progress ở đây gần như tức thì là đúng.
      setPhase({ percent: 0, label: "Đang cắt bảng theo dòng header..." });
      await yieldToPaint();

      const tables = files.map((f) => {
        const idx = resolveHeaderIndex(f.matrix, f.headerRow, 0);
        return { file: f, table: tableFromMatrix(f.matrix, idx) };
      });

      // Validate từng file để báo rõ file nào thiếu cột, thay vì lỗi chung
      // sau khi đã gộp.
      const problems = tables
        .map(({ file, table }) => {
          const missing = missingColumns(table.columns, mode);
          return missing.length
            ? `• ${file.name} (header dòng ${file.headerRow}): thiếu cột ${missing
                .map((m) => `'${m}'`)
                .join(", ")}`
            : null;
        })
        .filter(Boolean);

      if (problems.length) {
        throw new Error(
          `Không tìm thấy cột bắt buộc:\n${problems.join(
            "\n"
          )}\n\nHãy kiểm tra lại dòng header của các file trên.`
        );
      }

      setPhase({ percent: 30, label: "Đang gộp dữ liệu..." });
      await yieldToPaint();
      const merged = mergeTables(tables.map((t) => t.table));

      setPhase({
        percent: 40,
        label: `Đang pivot ${merged.rows.length.toLocaleString("vi-VN")} dòng...`,
      });
      await yieldToPaint();

      setPhase({ percent: 50, label: "Đang ghi file XLSX..." });
      await yieldToPaint();
      const { buffer, stats } = await runPivot(merged, mode);

      setPhase({ percent: 100, label: "Hoàn tất" });
      setResult({
        ...stats,
        total_rows: merged.rows.length,
        file_count: files.length,
      });
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      setDownload(URL.createObjectURL(blob));
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : "Đã xảy ra lỗi");
    } finally {
      busyRef.current = false;
      setLoading(false);
      setPhase(null);
    }
  };

  const reset = () => {
    setFiles([]);
    setResult(null);
    setDownload(null);
    setError(null);
  };

  const changeMode = (m: PivotMode) => {
    if (m === mode) return;
    setMode(m);
    setResult(null);
    setDownload(null);
    setError(null);
  };

  const totalRows = files.reduce((acc, f) => acc + f.dataRows, 0);

  return (
    <div>
      <p className="text-gray-500 text-sm mb-4">
        Thêm một hoặc nhiều file Excel/CSV danh sách bệnh nhân. Tool tự bỏ các
        dòng tiêu đề ở đầu file, gộp dữ liệu rồi pivot tổng tiền và trả file
        XLSX có định dạng.
      </p>

      {/* Segmented control */}
      <div className="flex bg-gray-100 rounded-lg p-1 mb-4">
        {(Object.keys(MODE_INFO) as PivotMode[]).map((m) => (
          <button
            key={m}
            onClick={() => changeMode(m)}
            className={`flex-1 py-2 px-3 text-sm font-semibold rounded-md transition-colors ${
              mode === m
                ? "bg-white text-blue-600 shadow-sm"
                : "text-gray-500 hover:text-gray-700"
            }`}
          >
            {MODE_INFO[m].label}
          </button>
        ))}
      </div>

      <div className="bg-blue-50 rounded-lg p-4 mb-6 text-sm text-gray-700 leading-relaxed">
        {MODE_INFO[mode].desc}
      </div>

      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 rounded-lg p-3 mb-4 text-sm whitespace-pre-wrap">
          {error}
        </div>
      )}

      {!result ? (
        <>
          <div
            className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all mb-4 ${
              dragOver
                ? "border-blue-500 bg-blue-50"
                : "border-gray-300 hover:border-blue-400 hover:bg-gray-50"
            }`}
            onDragOver={(e) => {
              e.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
            onClick={() => !busy && fileInputRef.current?.click()}
          >
            <svg
              className="w-10 h-10 text-gray-400 mx-auto mb-2"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={1.5}
                d="M12 16V4m0 0L8 8m4-4l4 4M4 14v4a2 2 0 002 2h12a2 2 0 002-2v-4"
              />
            </svg>
            <p className="text-gray-500 text-sm">
              Kéo thả file .xlsx / .csv vào đây hoặc{" "}
              <span className="font-semibold text-blue-600">
                click để chọn
              </span>
            </p>
            <p className="text-gray-400 text-xs mt-1">
              Chọn được nhiều file cùng lúc, hoặc thêm dần từng file
            </p>
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.csv"
            multiple
            className="hidden"
            onChange={(e) => {
              if (e.target.files?.length) {
                addFiles(Array.from(e.target.files));
              }
              e.target.value = "";
            }}
          />

          {files.length > 0 && (
            <div className="border border-gray-200 rounded-lg overflow-hidden mb-4">
              <table className="w-full text-sm">
                <thead className="bg-gray-50 text-gray-500 text-xs uppercase">
                  <tr>
                    <th className="text-left px-3 py-2 font-semibold">File</th>
                    <th className="text-center px-2 py-2 font-semibold w-28">
                      Dòng header
                    </th>
                    <th className="text-right px-2 py-2 font-semibold w-20">
                      Số dòng
                    </th>
                    <th className="text-right px-2 py-2 font-semibold w-16">
                      Số cột
                    </th>
                    <th className="w-8" />
                  </tr>
                </thead>
                <tbody>
                  {files.map((f) => (
                    <tr key={f.id} className="border-t border-gray-200">
                      <td className="px-3 py-2">
                        <span
                          className="block truncate max-w-[15rem]"
                          title={f.name}
                        >
                          {f.name}
                        </span>
                        {f.needsManualHeader && (
                          <span className="block text-xs text-amber-600 mt-0.5">
                            Không tự nhận được dòng header — vui lòng nhập dòng
                            header thủ công
                          </span>
                        )}
                        {f.droppedColumns.length > 0 && (
                          <span className="block text-xs text-amber-600 mt-0.5">
                            {f.droppedColumns.length} cột trùng tên bị bỏ:{" "}
                            {f.droppedColumns.join(", ")}
                          </span>
                        )}
                      </td>
                      <td className="px-2 py-2 text-center">
                        <input
                          type="number"
                          min={1}
                          max={Math.max(1, f.matrix.length)}
                          value={f.headerRow}
                          disabled={busy}
                          onChange={(e) =>
                            changeHeaderRow(f.id, e.target.value)
                          }
                          className={`w-16 px-2 py-1 border rounded text-center tabular-nums ${
                            f.needsManualHeader
                              ? "border-amber-400 bg-amber-50"
                              : "border-gray-300"
                          }`}
                        />
                      </td>
                      <td className="px-2 py-2 text-right tabular-nums text-gray-700">
                        {f.dataRows.toLocaleString("vi-VN")}
                      </td>
                      <td className="px-2 py-2 text-right tabular-nums text-gray-700">
                        {f.columnCount}
                      </td>
                      <td className="px-2 py-2 text-center">
                        <button
                          onClick={() => removeFile(f.id)}
                          disabled={busy}
                          title="Bỏ file này"
                          className="text-gray-400 hover:text-red-600 disabled:opacity-40 text-lg leading-none"
                        >
                          ×
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot className="bg-gray-50 border-t border-gray-200">
                  <tr>
                    <td className="px-3 py-2 font-semibold text-gray-700">
                      {files.length} file
                    </td>
                    <td />
                    <td className="px-2 py-2 text-right font-bold text-blue-600 tabular-nums">
                      {totalRows.toLocaleString("vi-VN")}
                    </td>
                    <td colSpan={2} />
                  </tr>
                </tfoot>
              </table>
            </div>
          )}

          {phase && (
            <ProgressBar percent={phase.percent} label={phase.label} />
          )}

          <button
            onClick={handleSubmit}
            disabled={!files.length || busy}
            className="w-full py-3 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
          >
            {loading
              ? "Đang xử lý..."
              : reading
                ? "Đang đọc file..."
                : files.length > 1
                  ? `Gộp ${files.length} file, chạy Pivot và tải XLSX`
                  : "Chạy Pivot và tải XLSX"}
          </button>
        </>
      ) : (
        <div className="text-center">
          <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
            <svg
              className="w-8 h-8 text-green-600"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2.5}
                d="M5 13l4 4L19 7"
              />
            </svg>
          </div>
          <h2 className="text-xl font-bold text-gray-900 mb-4">
            Pivot thành công!
          </h2>

          <div className="bg-gray-50 rounded-lg p-4 mb-6 text-left text-sm">
            <div className="flex justify-between py-2">
              <span className="text-gray-600">Số file đã gộp</span>
              <span className="font-bold text-gray-900">
                {result.file_count}
              </span>
            </div>
            <div className="flex justify-between py-2 border-t border-gray-200">
              <span className="text-gray-600">Tổng số dòng dữ liệu</span>
              <span className="font-bold text-gray-900">
                {result.total_rows.toLocaleString("vi-VN")}
              </span>
            </div>
            <div className="flex justify-between py-2 border-t border-gray-200">
              <span className="text-gray-600">Số bệnh nhân unique</span>
              <span className="font-bold text-blue-600">
                {result.unique_patients}
              </span>
            </div>
            {result.unique_sources !== undefined && (
              <div className="flex justify-between py-2 border-t border-gray-200">
                <span className="text-gray-600">Số nguồn</span>
                <span className="font-bold text-blue-600">
                  {result.unique_sources}
                </span>
              </div>
            )}
          </div>

          {downloadUrl && (
            <a
              href={downloadUrl}
              download={MODE_INFO[mode].filename}
              className="inline-block w-full py-3 bg-green-600 text-white font-semibold rounded-lg hover:bg-green-700 transition-colors mb-3 text-center"
            >
              ⬇️ Tải file XLSX
            </a>
          )}
          <button
            onClick={() => {
              setResult(null);
              setDownload(null);
            }}
            className="w-full py-3 bg-blue-50 text-blue-700 font-medium rounded-lg hover:bg-blue-100 transition-colors mb-3"
          >
            ← Đổi mode / thêm file (giữ {files.length} file đang có)
          </button>
          <button
            onClick={reset}
            className="w-full py-3 bg-gray-100 text-gray-600 font-medium rounded-lg hover:bg-gray-200 transition-colors"
          >
            Xoá hết, bắt đầu lại
          </button>
        </div>
      )}
    </div>
  );
}
