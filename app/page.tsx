"use client";

import { useState } from "react";
import WholesaleTab from "./components/WholesaleTab";
import PivotTab from "./components/PivotTab";

type Tab = "wholesale" | "pivot";

const TABS: { id: Tab; label: string }[] = [
  { id: "wholesale", label: "Wholesale CSV" },
  { id: "pivot", label: "Pivot Bệnh Nhân" },
];

export default function Home() {
  const [tab, setTab] = useState<Tab>("wholesale");

  return (
    <main className="flex items-center justify-center min-h-screen p-4">
      <div className="bg-white rounded-2xl shadow-lg p-8 sm:p-12 max-w-lg w-full">
        <h1 className="text-2xl font-bold text-gray-900 mb-4">
          📊 Data Cleaner
        </h1>

        {/* Tab bar */}
        <div className="flex border-b border-gray-200 mb-6">
          {TABS.map((t) => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className={`px-4 py-2 -mb-px text-sm font-semibold border-b-2 transition-colors ${
                tab === t.id
                  ? "border-blue-600 text-blue-600"
                  : "border-transparent text-gray-500 hover:text-gray-700"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>

        {tab === "wholesale" ? <WholesaleTab /> : <PivotTab />}
      </div>
    </main>
  );
}
