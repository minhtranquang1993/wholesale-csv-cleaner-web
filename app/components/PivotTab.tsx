"use client";

import { useState, useCallback, useRef } from "react";
import { parseFile, runPivot, type PivotMode } from "../lib/pivot";

interface PivotResult {
  unique_patients: number;
  unique_sources?: number;
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
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<PivotResult | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = (f: File) => {
    const name = f.name.toLowerCase();
    if (!name.endsWith(".xlsx") && !name.endsWith(".csv")) {
      setError("Chỉ hỗ trợ file .xlsx hoặc .csv");
      return;
    }
    setFile(f);
    setError(null);
    setResult(null);
    setDownloadUrl(null);
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files.length) {
      handleFile(e.dataTransfer.files[0]);
    }
  }, []);

  const handleSubmit = async () => {
    if (!file) return;
    setLoading(true);
    setError(null);
    setResult(null);
    setDownloadUrl(null);

    try {
      // Xử lý ngay trên trình duyệt: không upload file lên server nên
      // không dính giới hạn body 4.5MB của Vercel, và dữ liệu bệnh nhân
      // không rời khỏi máy người dùng.
      const table = await parseFile(file.name, {
        text: () => file.text(),
        arrayBuffer: () => file.arrayBuffer(),
      });
      const { buffer, stats } = await runPivot(table, mode);

      setResult(stats);
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      setDownloadUrl(URL.createObjectURL(blob));
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : "Đã xảy ra lỗi");
    } finally {
      setLoading(false);
    }
  };

  const reset = () => {
    setFile(null);
    setResult(null);
    setDownloadUrl(null);
    setError(null);
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  };

  const changeMode = (m: PivotMode) => {
    if (m === mode) return;
    setMode(m);
    setResult(null);
    setDownloadUrl(null);
    setError(null);
  };

  return (
    <div>
      <p className="text-gray-500 text-sm mb-4">
        Upload file Excel/CSV danh sách bệnh nhân, tool sẽ pivot tổng tiền và
        trả file XLSX có định dạng.
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
            className={`border-2 border-dashed rounded-xl p-10 text-center cursor-pointer transition-all mb-6 ${
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
            onClick={() => fileInputRef.current?.click()}
          >
            <svg
              className="w-12 h-12 text-gray-400 mx-auto mb-3"
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
            {file && (
              <p className="text-blue-600 font-semibold mt-2">{file.name}</p>
            )}
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.csv"
            className="hidden"
            onChange={(e) => {
              if (e.target.files?.[0]) handleFile(e.target.files[0]);
            }}
          />
          <button
            onClick={handleSubmit}
            disabled={!file || loading}
            className="w-full py-3 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
          >
            {loading ? "Đang xử lý..." : "Chạy Pivot và tải XLSX"}
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
            onClick={reset}
            className="w-full py-3 bg-gray-100 text-gray-600 font-medium rounded-lg hover:bg-gray-200 transition-colors"
          >
            ← Pivot file khác
          </button>
        </div>
      )}
    </div>
  );
}
