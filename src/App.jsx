import React, { useState } from "react";
import JournalTool from "./JournalTool";
import AccountsTool from "./AccountsTool";

const COLORS = { teal: "#0E3B36", line: "#D9D2BE", paper: "#F7F4EC" };

export default function App() {
  const [tab, setTab] = useState("accounts");

  return (
    <div dir="rtl" style={{ background: COLORS.paper, fontFamily: "'Segoe UI', Tahoma, Arial, sans-serif" }}>
      <div className="mx-auto max-w-5xl px-4 pt-6 sm:px-8">
        <div className="flex gap-2 border-b" style={{ borderColor: COLORS.line }}>
          <button
            onClick={() => setTab("entries")}
            className="px-4 py-2 text-sm font-semibold"
            style={{
              color: tab === "entries" ? COLORS.teal : "#8A8163",
              borderBottom: tab === "entries" ? `2px solid ${COLORS.teal}` : "2px solid transparent",
            }}
          >
            استيراد القيود
          </button>
          <button
            onClick={() => setTab("accounts")}
            className="px-4 py-2 text-sm font-semibold"
            style={{
              color: tab === "accounts" ? COLORS.teal : "#8A8163",
              borderBottom: tab === "accounts" ? `2px solid ${COLORS.teal}` : "2px solid transparent",
            }}
          >
            استيراد الشجرة
          </button>
        </div>
      </div>
      {tab === "entries" ? <JournalTool /> : <AccountsTool />}
    </div>
  );
}
