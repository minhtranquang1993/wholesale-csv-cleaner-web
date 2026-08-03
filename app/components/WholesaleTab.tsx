"use client";

import { useState, useCallback, useRef } from "react";
import ProgressBar from "./ProgressBar";
import { cleanWholesaleCsv, type CleanPhase, type CleanStats } from "../lib/clean";

interface CleanResult extends CleanStats {
  rowCount: number;
}

export default function WholesaleTab() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [phase, setPhase] = useState<CleanPhase | null>(null);
  const [result, setResult] = useState<CleanResult | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  // Gate bằng ref, không dùng state: onDrop là useCallback deps [] nên đọc
  // state sẽ luôn thấy giá trị của lần render đầu.
  const busyRef = useRef(false);

  /** Đổi link tải, tự revoke link cũ để blob file XLSX trước không bị giữ lại. */
  const setDownload = (url: string | null) => {
    setDownloadUrl((prev) => {
      if (prev && prev !== url) URL.revokeObjectURL(prev);
      return url;
    });
  };

  const handleFile = (f: File) => {
    // Đang xử lý thì không cho đổi file: cleanWholesaleCsv đã giữ file cũ nên
    // đổi tên hiển thị sẽ khiến kết quả không khớp file người dùng thấy.
    if (busyRef.current) return;
    if (!f.name.toLowerCase().endsWith(".csv")) {
      setError("Chỉ hỗ trợ file .csv");
      return;
    }
    setFile(f);
    setError(null);
    setResult(null);
    setDownload(null);
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files.length) {
      handleFile(e.dataTransfer.files[0]);
    }
  }, []);

  const handleSubmit = async () => {
    if (!file || busyRef.current) return;
    busyRef.current = true;
    setLoading(true);
    setError(null);
    setResult(null);
    setDownload(null);
    setPhase({ percent: 0, label: "Đang bắt đầu..." });

    try {
      // Xử lý ngay trên trình duyệt: không dính giới hạn body 4.5MB của Vercel
      // và đo được tiến độ thật theo từng phase.
      const { buffer, stats, rowCount } = await cleanWholesaleCsv(file, setPhase);

      setResult({ ...stats, rowCount });
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
    setFile(null);
    setResult(null);
    setDownload(null);
    setError(null);
  };

  return (
    <div>
      <p className="text-gray-500 text-sm mb-6">
        Upload file CSV, tool sẽ làm sạch dữ liệu và trả file XLSX cho bạn
        download.
      </p>

      <div className="bg-blue-50 rounded-lg p-4 mb-6 text-sm text-gray-700 leading-relaxed">
        <p>
          <span className="font-semibold text-blue-600">Rule 1:</span> Nếu
          first_source_url rỗng và url không rỗng → gán first_source_url =
          prefix + url
        </p>
        <p>
          <span className="font-semibold text-blue-600">Rule 2:</span> Lấy
          utm_source từ URL gán vào first_user_source
        </p>
        <p>
          <span className="font-semibold text-blue-600">Rule 3:</span> Lấy
          utm_medium từ URL gán vào first_user_medium
        </p>
      </div>

      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 rounded-lg p-3 mb-4 text-sm">
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
              Kéo thả file CSV vào đây hoặc{" "}
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
            accept=".csv"
            className="hidden"
            onChange={(e) => {
              if (e.target.files?.[0]) handleFile(e.target.files[0]);
            }}
          />
          {phase && (
            <ProgressBar
              percent={phase.percent}
              label={phase.label}
              indeterminate={phase.indeterminate}
            />
          )}
          <button
            onClick={handleSubmit}
            disabled={!file || loading}
            className="w-full py-3 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
          >
            {loading ? "Đang xử lý..." : "Làm sạch và tải XLSX"}
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
            Làm sạch thành công!
          </h2>

          <div className="bg-gray-50 rounded-lg p-4 mb-6 text-left text-sm">
            <div className="flex justify-between py-2">
              <span className="text-gray-600">Tổng số dòng</span>
              <span className="font-bold text-gray-900">
                {result.rowCount.toLocaleString("vi-VN")} dòng
              </span>
            </div>
            <div className="flex justify-between py-2 border-t border-gray-200">
              <span className="text-gray-600">first_source_url cập nhật</span>
              <span className="font-bold text-blue-600">
                {result.first_source_url} dòng
              </span>
            </div>
            <div className="flex justify-between py-2 border-t border-gray-200">
              <span className="text-gray-600">first_user_source cập nhật</span>
              <span className="font-bold text-blue-600">
                {result.first_user_source} dòng
              </span>
            </div>
            <div className="flex justify-between py-2 border-t border-gray-200">
              <span className="text-gray-600">first_user_medium cập nhật</span>
              <span className="font-bold text-blue-600">
                {result.first_user_medium} dòng
              </span>
            </div>
          </div>

          {downloadUrl && (
            <a
              href={downloadUrl}
              download="wholesale_cleaned.xlsx"
              className="inline-block w-full py-3 bg-green-600 text-white font-semibold rounded-lg hover:bg-green-700 transition-colors mb-3 text-center"
            >
              ⬇️ Tải file XLSX
            </a>
          )}
          <button
            onClick={reset}
            className="w-full py-3 bg-gray-100 text-gray-600 font-medium rounded-lg hover:bg-gray-200 transition-colors"
          >
            ← Làm sạch file khác
          </button>
        </div>
      )}
    </div>
  );
}
