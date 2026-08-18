import React, { useState, useMemo, useRef, useCallback } from "react";
import {
  Upload, FileSpreadsheet, CheckCircle2, AlertTriangle, XCircle, Loader2,
  Download, ChevronDown, ChevronUp, Info, RefreshCcw, Copy, Sparkles,
} from "lucide-react";
import { readWorkbookRows, parseChartFile, parseEntriesFile, buildParentInfo, validateEntryStructure, buildSemanticsPrompt } from "./lib/excelCore";
import { buildImportFile, downloadBlob, buildPasteText } from "./lib/excelExport";
import { callClaude, parseJsonResponse } from "./lib/claudeProxy";

const COLORS = {
  paper: "#F7F4EC", ink: "#1F2A24", teal: "#0E3B36", tealLight: "#155850",
  gold: "#B9852F", amber: "#C97A2B", red: "#A6382C", green: "#2F6F4E", line: "#D9D2BE",
};

function UploadCard({ title, subtitle, fileName, ok, count, onFile, busy }) {
  const inputRef = useRef(null);
  return (
    <div
      className="rounded-lg border-2 border-dashed p-4 text-center transition cursor-pointer"
      style={{ borderColor: ok ? COLORS.green : COLORS.line, background: "#FFFDF7" }}
      onClick={() => inputRef.current?.click()}
    >
      <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden"
        onChange={(e) => e.target.files[0] && onFile(e.target.files[0])} />
      <div className="flex flex-col items-center gap-1">
        {busy ? <Loader2 size={26} className="animate-spin" style={{ color: COLORS.gold }} />
          : ok ? <CheckCircle2 size={26} style={{ color: COLORS.green }} /> : <Upload size={26} style={{ color: COLORS.gold }} />}
        <p className="text-sm font-semibold">{title}</p>
        <p className="text-xs" style={{ color: "#5B5340" }}>{subtitle}</p>
        {fileName && (
          <p className="mt-1 flex items-center gap-1 text-xs" style={{ color: COLORS.tealLight }}>
            <FileSpreadsheet size={12} /> {fileName} {count && `— ${count}`}
          </p>
        )}
      </div>
    </div>
  );
}

function SummaryStat({ label, value, color }) {
  return (
    <div className="rounded-lg border px-4 py-3 text-center" style={{ borderColor: COLORS.line, background: "#FFFDF7" }}>
      <p className="text-2xl font-bold" style={{ color }}>{value}</p>
      <p className="text-xs" style={{ color: "#5B5340" }}>{label}</p>
    </div>
  );
}

function AccountPicker({ accounts, value, onChange, hasError, parentCodes }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const selectable = useMemo(() => {
    return accounts
      .filter((a) => !parentCodes.has(a.code))
      .slice()
      .sort((a, b) => a.code.localeCompare(b.code, undefined, { numeric: true }));
  }, [accounts, parentCodes]);
  const effectiveQuery = query !== "" ? query : value || "";
  const filtered = useMemo(() => {
    const q = effectiveQuery.trim().toLowerCase();
    if (!q) return selectable;
    return selectable.filter((a) => a.code.toLowerCase().includes(q) || a.name.toLowerCase().includes(q));
  }, [effectiveQuery, selectable]);
  return (
    <div className="relative">
      <input
        value={query || value || ""}
        onChange={(e) => { setQuery(e.target.value); onChange(e.target.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
        placeholder="رمز الحساب" dir="ltr"
        className="w-full rounded-md border px-2 py-1.5 text-sm font-mono text-left focus:outline-none"
        style={{ borderColor: hasError ? COLORS.red : COLORS.line }}
      />
      {open && filtered.length > 0 && (
        <div className="absolute z-20 mt-1 max-h-56 w-72 overflow-y-auto rounded-md border bg-white shadow-lg" style={{ borderColor: COLORS.line, right: 0 }}>
          {filtered.map((a) => (
            <div key={a.code} onMouseDown={() => { onChange(a.code); setQuery(""); setOpen(false); }}
              className="cursor-pointer px-3 py-1.5 text-sm hover:bg-[#F3EFE2] text-right">
              <span className="font-mono ml-2" style={{ color: COLORS.teal }}>{a.code}</span> {a.name}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function EntryCard({ entry, issues, isOpen, onToggle, chartAccountsList, onUpdateRow, onUpdateMeta, parentCodes, onApplySuggestion, onDismiss }) {
  const hasErrors = issues.length > 0;
  const statusColor = hasErrors ? COLORS.red : COLORS.green;
  return (
    <div className="overflow-hidden rounded-lg border" style={{ borderColor: COLORS.line, background: "#FFFDF7" }}>
      <button onClick={onToggle} className="flex w-full items-center justify-between gap-3 px-4 py-3 text-right">
        <div className="flex items-center gap-3">
          <span className="flex h-6 w-6 items-center justify-center rounded-full text-xs font-bold text-white" style={{ background: statusColor }}>
            {issues.length === 0 ? <CheckCircle2 size={14} /> : issues.length}
          </span>
          <div>
            <p className="text-sm font-semibold">قيد #{entry.seq} — {entry.desc || "بدون وصف"}</p>
            <p className="text-xs" style={{ color: "#5B5340" }}>{entry.date || "بدون تاريخ"}</p>
          </div>
        </div>
        {isOpen ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
      </button>
      {isOpen && (
        <div className="border-t px-4 py-3" style={{ borderColor: COLORS.line }}>
          <div className="mb-3 flex flex-wrap items-center gap-3 text-xs">
            <label className="flex items-center gap-1">التاريخ:
              <input dir="ltr" value={entry.date || ""} onChange={(e) => onUpdateMeta(entry.seq, "date", e.target.value)}
                placeholder="dd/mm/yyyy" className="rounded border px-2 py-1 font-mono" style={{ borderColor: COLORS.line, width: 100 }} />
            </label>
            <label className="flex flex-1 items-center gap-1">الوصف:
              <input value={entry.desc || ""} onChange={(e) => onUpdateMeta(entry.seq, "desc", e.target.value)}
                className="flex-1 rounded border px-2 py-1" style={{ borderColor: COLORS.line }} />
            </label>
          </div>
          <table className="mb-3 w-full text-xs">
            <thead><tr style={{ color: "#5B5340" }}>
              <th className="pb-1 text-right font-medium">الرمز</th>
              <th className="pb-1 text-right font-medium">اسم الحساب</th>
              <th className="pb-1 text-right font-medium">مدين</th>
              <th className="pb-1 text-right font-medium">دائن</th>
              <th className="pb-1 text-right font-medium">تعليق</th>
            </tr></thead>
            <tbody>
              {entry.rows.map((r) => {
                const rowIssue = issues.find((i) => i.rowIndex === r._rowIndex);
                const acc = chartAccountsList.find((a) => a.code === r.code);
                return (
                  <tr key={r._rowIndex} className="border-t" style={{ borderColor: COLORS.line }}>
                    <td className="py-1.5 pl-2" style={{ width: 130 }}>
                      <AccountPicker accounts={chartAccountsList} value={r.code} hasError={!!rowIssue} parentCodes={parentCodes}
                        onChange={(v) => onUpdateRow(entry.seq, r._rowIndex, "code", v)} />
                    </td>
                    <td className="py-1.5 pl-2 text-right" style={{ color: !acc ? COLORS.red : rowIssue?.type === "parent_account" ? COLORS.amber : COLORS.ink }}>
                      {acc ? acc.name : "—"}{rowIssue?.type === "parent_account" && " (رئيسي)"}
                    </td>
                    <td className="py-1.5 pl-2">
                      <input dir="ltr" value={r.debit ?? ""} onChange={(e) => onUpdateRow(entry.seq, r._rowIndex, "debit", e.target.value === "" ? null : parseFloat(e.target.value))}
                        className="w-20 rounded border px-1.5 py-1 font-mono text-left" style={{ borderColor: COLORS.line }} />
                    </td>
                    <td className="py-1.5 pl-2">
                      <input dir="ltr" value={r.credit ?? ""} onChange={(e) => onUpdateRow(entry.seq, r._rowIndex, "credit", e.target.value === "" ? null : parseFloat(e.target.value))}
                        className="w-20 rounded border px-1.5 py-1 font-mono text-left" style={{ borderColor: COLORS.line }} />
                    </td>
                    <td className="py-1.5">
                      <input value={r.comment || ""} onChange={(e) => onUpdateRow(entry.seq, r._rowIndex, "comment", e.target.value)}
                        className="w-full rounded border px-1.5 py-1" style={{ borderColor: COLORS.line }} />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {issues.map((issue) => (
            <div key={issue.id} className="mb-2 flex items-start justify-between gap-2 rounded-md border px-3 py-2 text-xs"
              style={{ borderColor: issue.severity === "warning" ? COLORS.amber : COLORS.red, background: issue.severity === "warning" ? "#FBF3E6" : "#FBEDEA" }}>
              <div className="flex items-start gap-2">
                <AlertTriangle size={14} className="mt-0.5 shrink-0" style={{ color: issue.severity === "warning" ? COLORS.amber : COLORS.red }} />
                <p>{issue.message}</p>
              </div>
              {issue.type === "semantic_mismatch" && (
                <div className="flex shrink-0 gap-1">
                  {issue.suggestedCode && (
                    <button onClick={() => onApplySuggestion(entry.seq, issue.rowIndex, issue.suggestedCode, issue.id)}
                      className="rounded px-2 py-1 text-white" style={{ background: COLORS.green }}>تطبيق</button>
                  )}
                  <button onClick={() => onDismiss(issue.id)} className="rounded border px-2 py-1" style={{ borderColor: COLORS.line }}>تجاهل</button>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default function JournalTool() {
  const [chartAccounts, setChartAccounts] = useState(null);
  const [chartFileName, setChartFileName] = useState("");
  const [chartBusy, setChartBusy] = useState(false);
  const [entries, setEntries] = useState(null);
  const [entriesFileName, setEntriesFileName] = useState("");
  const [entriesBusy, setEntriesBusy] = useState(false);
  const [parseError, setParseError] = useState("");
  const [expanded, setExpanded] = useState({});
  const [copyStatus, setCopyStatus] = useState("");
  const [showManualCopy, setShowManualCopy] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [includeAI, setIncludeAI] = useState(true);
  const [analyzingAI, setAnalyzingAI] = useState(false);
  const [aiProgress, setAiProgress] = useState({ done: 0, total: 0 });
  const [semanticIssues, setSemanticIssues] = useState({});
  const [resolvedIds, setResolvedIds] = useState({});
  const [aiError, setAiError] = useState("");

  const chartMap = useMemo(() => {
    if (!chartAccounts) return {};
    const m = {};
    chartAccounts.forEach((a) => (m[a.code] = a));
    return m;
  }, [chartAccounts]);

  const parentInfo = useMemo(() => (chartAccounts ? buildParentInfo(chartAccounts) : { parentCodes: new Set(), childrenByParent: {} }), [chartAccounts]);

  const structuralIssuesBySeq = useMemo(() => {
    if (!entries || !chartAccounts) return {};
    const out = {};
    entries.forEach((entry) => (out[entry.seq] = validateEntryStructure(entry, chartMap, parentInfo)));
    return out;
  }, [entries, chartMap, parentInfo, chartAccounts]);

  const issuesBySeq = useMemo(() => {
    const out = {};
    (entries || []).forEach((entry) => {
      const struct = structuralIssuesBySeq[entry.seq] || [];
      const sem = semanticIssues[entry.seq] || [];
      out[entry.seq] = [...struct, ...sem].filter((i) => !resolvedIds[i.id]);
    });
    return out;
  }, [entries, structuralIssuesBySeq, semanticIssues, resolvedIds]);

  const handleChartUpload = async (file) => {
    setParseError(""); setChartFileName(file.name); setChartBusy(true);
    try {
      const rows = await readWorkbookRows(file);
      setChartAccounts(parseChartFile(rows));
    } catch (err) {
      setParseError("خطأ في قراءة ملف شجرة الحسابات: " + err.message);
      setChartAccounts(null);
    } finally { setChartBusy(false); }
  };

  const handleEntriesUpload = async (file) => {
    setParseError(""); setEntriesFileName(file.name); setEntriesBusy(true);
    try {
      const rows = await readWorkbookRows(file);
      const grouped = parseEntriesFile(rows);
      if (grouped.length === 0) throw new Error("تم التعرف على تنسيق الملف لكن لم يتم العثور على أي قيود بداخله");
      setEntries(grouped);
      setSemanticIssues({});
      setResolvedIds({});
      setAiError("");
    } catch (err) {
      setParseError("خطأ في قراءة ملف القيود: " + err.message);
      setEntries(null);
    } finally { setEntriesBusy(false); }
  };

  function chunkArray(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  }

  const runAIReview = useCallback(async () => {
    if (!entries || !chartAccounts || !includeAI) return;
    setAnalyzingAI(true);
    setAiError("");
    const batches = chunkArray(entries, 12);
    setAiProgress({ done: 0, total: batches.length });
    const merged = {};
    try {
      for (const batch of batches) {
        try {
          const prompt = buildSemanticsPrompt(batch, chartAccounts, parentInfo.parentCodes);
          const data = await callClaude({ model: "claude-sonnet-4-6", max_tokens: 4000, messages: [{ role: "user", content: prompt }] });
          const result = parseJsonResponse(data);
          (result || []).forEach((item) => {
            if (!item.hasIssue) return;
            const entry = entries.find((e) => String(e.seq) === String(item.seq));
            if (!entry) return;
            const row = entry.rows[item.rowIndex];
            if (!row) return;
            if (!merged[entry.seq]) merged[entry.seq] = [];
            merged[entry.seq].push({
              id: `${entry.seq}-row${item.rowIndex}-semantic`,
              type: "semantic_mismatch",
              severity: "warning",
              rowIndex: row._rowIndex,
              suggestedCode: item.suggestedCode,
              suggestedName: item.suggestedName,
              message: `السطر ${item.rowIndex + 1}: الحساب المستخدم لا يتوافق منطقياً مع وصف القيد — ${item.reason || ""}`,
            });
          });
        } catch (batchErr) {
          console.error("batch failed", batchErr);
        }
        setAiProgress((p) => ({ ...p, done: p.done + 1 }));
      }
      setSemanticIssues(merged);
    } catch (err) {
      setAiError("تعذر إتمام المراجعة الذكية: " + err.message);
    } finally {
      setAnalyzingAI(false);
    }
  }, [entries, chartAccounts, includeAI, parentInfo]);

  const updateRow = (seq, rowIndex, field, value) => {
    setEntries((prev) => prev.map((entry) => entry.seq !== seq ? entry :
      { ...entry, rows: entry.rows.map((r) => r._rowIndex === rowIndex ? { ...r, [field]: value } : r) }));
  };
  const updateEntryMeta = (seq, field, value) => {
    setEntries((prev) => prev.map((entry) => (entry.seq === seq ? { ...entry, [field]: value } : entry)));
  };

  const applySuggestion = (seq, rowIndex, suggestedCode, issueId) => {
    updateRow(seq, rowIndex, "code", suggestedCode);
    setResolvedIds((prev) => ({ ...prev, [issueId]: true }));
  };
  const dismissIssue = (issueId) => {
    setResolvedIds((prev) => ({ ...prev, [issueId]: true }));
  };

  const totalEntries = entries ? entries.length : 0;
  const entriesWithIssues = entries ? entries.filter((e) => (issuesBySeq[e.seq] || []).length > 0).length : 0;
  const totalOpenIssues = Object.values(issuesBySeq).reduce((s, arr) => s + arr.length, 0);
  const ready = chartAccounts && entries;

  const handleDownload = async () => {
    setDownloading(true);
    try {
      const blob = await buildImportFile(entries);
      downloadBlob(blob, "قيود_جاهزة_للاستيراد.xlsx");
    } catch (err) {
      setParseError("تعذر إنشاء الملف: " + err.message);
    } finally { setDownloading(false); }
  };

  const copyToClipboard = async () => {
    const text = buildPasteText(entries);
    let success = false;
    if (navigator.clipboard?.writeText) {
      try { await navigator.clipboard.writeText(text); success = true; } catch { success = false; }
    }
    if (!success) {
      try {
        const ta = document.createElement("textarea");
        ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0";
        document.body.appendChild(ta); ta.focus(); ta.select();
        success = document.execCommand("copy");
        document.body.removeChild(ta);
      } catch { success = false; }
    }
    if (success) { setCopyStatus("copied"); setShowManualCopy(false); setTimeout(() => setCopyStatus(""), 3000); }
    else { setCopyStatus("failed"); setShowManualCopy(true); }
  };

  const resetAll = () => {
    setChartAccounts(null); setChartFileName(""); setEntries(null); setEntriesFileName("");
    setParseError(""); setExpanded({}); setSemanticIssues({}); setResolvedIds({}); setAiError("");
  };

  return (
    <div dir="rtl" className="min-h-screen w-full" style={{ background: COLORS.paper, color: COLORS.ink, fontFamily: "'Segoe UI', Tahoma, Arial, sans-serif" }}>
      <div className="mx-auto max-w-5xl px-4 py-6 sm:px-8">
        <div className="mb-6 flex items-center justify-between border-b pb-4" style={{ borderColor: COLORS.line }}>
          <div>
            <h1 className="text-xl font-bold tracking-tight" style={{ color: COLORS.teal, fontFamily: "Georgia, serif" }}>مدقّق استيراد القيود — قيود</h1>
            <p className="mt-1 text-sm" style={{ color: "#5B5340" }}>ارفع شجرة الحسابات وملف القيود، وسيتم فحصها وتجهيزها للاستيراد تلقائياً</p>
          </div>
          {(chartAccounts || entries) && (
            <button onClick={resetAll} className="flex items-center gap-1 rounded-md border px-3 py-1.5 text-xs" style={{ borderColor: COLORS.line, color: COLORS.tealLight }}>
              <RefreshCcw size={14} /> بدء من جديد
            </button>
          )}
        </div>

        <div className="mb-6 flex items-start gap-2 rounded-lg border px-4 py-3 text-xs leading-relaxed" style={{ borderColor: COLORS.line, background: "#FFFDF7" }}>
          <Info size={16} className="mt-0.5 shrink-0" style={{ color: COLORS.gold }} />
          <div style={{ color: "#5B5340" }}>
            المعايير المعتمدة: صيغة التاريخ dd/mm/yyyy · مدين = دائن لكل قيد · صف فارغ إلزامي بين كل قيد وآخر · لا يجوز الترحيل
            على حساب رئيسي له حسابات فرعية · المراجعة الذكية بالذكاء الاصطناعي متاحة (تحتاج مفتاح Anthropic API مُعد على
            الخادم — راجع README).
          </div>
        </div>

        <div className="mb-6 grid gap-4 sm:grid-cols-2">
          <UploadCard title="شجرة الحسابات" subtitle="ملف الحسابات الخاص بالعميل" fileName={chartFileName} ok={!!chartAccounts} busy={chartBusy}
            count={chartAccounts ? `${chartAccounts.length} حساب` : ""} onFile={handleChartUpload} />
          <UploadCard title="القيود المراد استيرادها" subtitle="ملف القيود بصيغة قيود (مسودة)" fileName={entriesFileName} ok={!!entries} busy={entriesBusy}
            count={entries ? `${entries.length} قيد` : ""} onFile={handleEntriesUpload} />
        </div>

        {parseError && (
          <div className="mb-6 flex items-center gap-2 rounded-md border px-4 py-2 text-sm" style={{ borderColor: COLORS.red, color: COLORS.red }}>
            <XCircle size={16} /> {parseError}
          </div>
        )}

        {ready && (
          <>
            <div className="mb-6 flex flex-wrap items-center gap-3">
              <button onClick={runAIReview} disabled={analyzingAI || !includeAI}
                className="flex items-center gap-2 rounded-md px-4 py-2 text-sm font-medium text-white disabled:opacity-50" style={{ background: COLORS.teal }}>
                {analyzingAI ? <Loader2 size={16} className="animate-spin" /> : <Sparkles size={16} />}
                {analyzingAI ? `جاري المراجعة الذكية... (${aiProgress.done}/${aiProgress.total})` : "تشغيل المراجعة الذكية"}
              </button>
              <label className="flex items-center gap-2 text-xs" style={{ color: "#5B5340" }}>
                <input type="checkbox" checked={includeAI} onChange={(e) => setIncludeAI(e.target.checked)} />
                تضمين المراجعة الذكية للمنطق المحاسبي
              </label>
            </div>
            {aiError && (
              <div className="mb-6 flex items-center gap-2 rounded-md border px-4 py-2 text-sm" style={{ borderColor: COLORS.amber, color: COLORS.amber }}>
                <AlertTriangle size={16} /> {aiError}
              </div>
            )}

            <div className="mb-6 grid grid-cols-3 gap-3">
              <SummaryStat label="إجمالي القيود" value={totalEntries} color={COLORS.teal} />
              <SummaryStat label="قيود سليمة" value={totalEntries - entriesWithIssues} color={COLORS.green} />
              <SummaryStat label="قيود بها أخطاء" value={entriesWithIssues} color={totalOpenIssues > 0 ? COLORS.red : COLORS.green} />
            </div>

            <div className="space-y-3">
              {entries.map((entry) => {
                const issues = issuesBySeq[entry.seq] || [];
                const isOpen = expanded[entry.seq] ?? issues.length > 0;
                return (
                  <EntryCard key={entry.seq} entry={entry} issues={issues} isOpen={isOpen} parentCodes={parentInfo.parentCodes}
                    onToggle={() => setExpanded((p) => ({ ...p, [entry.seq]: !isOpen }))}
                    chartAccountsList={chartAccounts} onUpdateRow={updateRow} onUpdateMeta={updateEntryMeta}
                    onApplySuggestion={applySuggestion} onDismiss={dismissIssue} />
                );
              })}
            </div>

            <div className="mt-8 flex flex-col items-center gap-2 border-t pt-6" style={{ borderColor: COLORS.line }}>
              {totalOpenIssues > 0 && (
                <p className="flex items-center gap-1 text-xs" style={{ color: COLORS.amber }}>
                  <AlertTriangle size={14} /> ما زال هناك {totalOpenIssues} خطأ لم يُحل — يمكنك التنزيل بعد المراجعة
                </p>
              )}
              <div className="flex flex-wrap items-center justify-center gap-3">
                <button onClick={handleDownload} disabled={downloading}
                  className="flex items-center gap-2 rounded-md px-6 py-2.5 text-sm font-semibold text-white disabled:opacity-50" style={{ background: COLORS.gold }}>
                  {downloading ? <Loader2 size={16} className="animate-spin" /> : <Download size={16} />} تنزيل الملف الجاهز للاستيراد
                </button>
                <button onClick={copyToClipboard} className="flex items-center gap-2 rounded-md border px-5 py-2.5 text-sm font-semibold" style={{ borderColor: COLORS.green, color: COLORS.green }}>
                  <Copy size={16} /> نسخ البيانات
                </button>
              </div>
              {copyStatus === "copied" && <span className="flex items-center gap-1 text-xs" style={{ color: COLORS.green }}><CheckCircle2 size={14} /> تم النسخ</span>}
              {showManualCopy && (
                <textarea readOnly dir="ltr" value={buildPasteText(entries)} onFocus={(e) => e.target.select()}
                  className="mt-1 h-32 w-full max-w-md rounded border p-2 font-mono text-xs" style={{ borderColor: COLORS.line }} />
              )}
              <p className="mx-auto mt-1 max-w-md text-center text-[11px] leading-relaxed" style={{ color: "#8A8163" }}>
                الملف مبني مباشرة فوق نسخة قالب قيود الرسمي المرفقة بالتطبيق — التنسيق (الألوان، الخط، عرض الأعمدة) محفوظ 100% تلقائياً.
              </p>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
