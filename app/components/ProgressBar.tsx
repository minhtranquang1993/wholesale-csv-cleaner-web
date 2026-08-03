"use client";

interface ProgressBarProps {
  /** 0-100 */
  percent: number;
  label: string;
  /** Phase blocking không chia nhỏ được → chạy sọc thay vì hiện %. */
  indeterminate?: boolean;
}

export default function ProgressBar({
  percent,
  label,
  indeterminate,
}: ProgressBarProps) {
  const clamped = Math.min(100, Math.max(0, percent));

  return (
    <div className="mb-4">
      <div className="flex justify-between items-center mb-1.5 text-sm">
        <span className="text-gray-600">{label}</span>
        {!indeterminate && (
          <span className="font-semibold text-blue-600 tabular-nums">
            {Math.round(clamped)}%
          </span>
        )}
      </div>
      <div className="h-2 bg-gray-200 rounded-full overflow-hidden">
        {indeterminate ? (
          <div className="h-full w-1/3 bg-blue-600 rounded-full animate-progress-slide" />
        ) : (
          <div
            className="h-full bg-blue-600 rounded-full transition-[width] duration-200"
            style={{ width: `${clamped}%` }}
          />
        )}
      </div>
    </div>
  );
}
