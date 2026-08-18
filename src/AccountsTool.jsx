import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import {
  Upload, FileSpreadsheet, CheckCircle2, AlertTriangle, XCircle, Loader2,
  Download, Info, Copy, Sparkles,
} from "lucide-react";
import { normalizeCode, extractParentCode } from "./lib/excelCore";
import { callClaude, parseJsonResponse, arrayBufferToBase64 } from "./lib/claudeProxy";

const COLORS = {
  paper: "#F7F4EC", ink: "#1F2A24", teal: "#0E3B36", tealLight: "#155850",
  gold: "#B9852F", amber: "#C97A2B", red: "#A6382C", green: "#2F6F4E", line: "#D9D2BE",
};

const LEVEL1_TYPES = ["الاصول", "الالتزامات", "المصاريف", "حقوق الملاك", "الايرادات"];

const LEVEL2_TYPES = [
  "الأصول المتداولة", "الأصول غير المتداولة", "رأس المال المصدر", "حقوق الملاك الأخرى",
  "الأرباح المبقاة", "التكلفة المباشرة", "تكاليف غير تشغيلية", "تكاليف تشغيلية",
  "الالتزامات المتداولة", "الالتزامات غير المتداولة", "الإيرادات الأخرى", "المبيعات",
];

const LEVEL3_TYPES = [
  "المدينون", "حساب البنك", "سلف موظفين", "المخزون", "النقدية ومافي حكمها", "أصول متداولة أخرى",
  "عهد نقدية", "مصروفات مقدمة", "أصول غير ملموسة", "أصول غير متداولة أخرى", "عقارات وآلات ومعدات",
  "رأس المال الإضافي المدفوع", "حقوق الموظفين", "رأس المال", "حقوق ملكية أخرى", "الاحتياطيات",
  "الأرباح المبقاة (أو الخسائر)", "تكلفة المبيعات", "تكاليف مباشرة أخرى", "ضرائب", "مصاريف الإطفاء",
  "مصاريف الاستهلاك", "مجمع الإطفاء", "مكافآت وحوافز", "مصاريف عمومية وإدارية", "مصاريف تسويقية",
  "تكاليف تشغيلية أخرى", "الرواتب", "مصاريف تقنية واستشارية", "الدائنون", "مصاريف مستحقة",
  "الرواتب والمبالغ المستحقة للموظفين", "مجمع الاستهلاك", "مخصص الديون المشكوك في تحصيلها",
  "مخصص مكافأة نهاية الخدمة", "التزامات متداولة أخرى", "مخصصات", "قروض قصيرة الأجل",
  "الضرائب المستحقة", "الإيرادات المقدمة", "قروض طويلة الأجل", "التزامات غير متداولة أخرى",
  "إيرادات أخرى", "المبيعات", "الضريبة", "الزكاة", "استثمارات بشركة تابعة", "مكاسب/خسائر بيع أصول",
  "توزيع الأرباح", "الزكاة المستحقة", "ترجمة عملات أجنبية", "مصروف فوائد",
  "ضريبة القيمة المضافة المستحقة", "فوائد مستحقة", "مكاسب/خسائر بيع أصول غير ملموسة",
  "مصاريف البحث والتطوير", "ضمان حسن التنفيذ", "الجزء المتداول من التزامات طويلة أجل", "مشاريع تحت التنفيذ",
];

const EQUITY_LEVEL2_TYPES = ["رأس المال المصدر", "حقوق الملاك الأخرى", "الأرباح المبقاة"];

function typesForLevel(level) {
  const lvl = Number(level);
  if (lvl === 2) return LEVEL2_TYPES;
  if (lvl >= 3) return LEVEL3_TYPES;
  return LEVEL1_TYPES;
}

function findHeaderRowIndex(rows, mustInclude) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    if (row.some((c) => typeof c === "string" && c.trim() === mustInclude)) return i;
  }
  return -1;
}
function colIndex(headerRow, ...candidates) {
  for (const cand of candidates) {
    const idx = headerRow.findIndex((h) => typeof h === "string" && h.trim().includes(cand));
    if (idx !== -1) return idx;
  }
  return -1;
}

function parseExistingTreeFile(rows) {
  const hIdx = findHeaderRowIndex(rows, "الرمز");
  if (hIdx === -1) throw new Error("لم يتم العثور على عمود 'الرمز' في ملف الشجرة الافتراضية");
  const header = rows[hIdx];
  const cCode = colIndex(header, "الرمز");
  const cNameAr = colIndex(header, "الاسم العربي", "اسم الحساب");
  const cNameEn = colIndex(header, "الاسم الانجليزي", "الاسم الإنجليزي");
  const cType = colIndex(header, "نوع الحساب", "النوع");
  const cLevel = colIndex(header, "المستوى");
  const cParent = colIndex(header, "الحساب الرئيسي", "Parent", "الحساب الأب");
  const cDesc = colIndex(header, "الوصف");
  const cPay = colIndex(header, "الدفع والتحصيل", "يمكن الدفع");

  const accounts = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const code = normalizeCode(r[cCode]);
    if (!code) continue;
    accounts.push({
      code,
      nameAr: cNameAr !== -1 ? (r[cNameAr] || "").toString().trim() : "",
      nameEn: cNameEn !== -1 ? (r[cNameEn] || "").toString().trim() : "",
      type: cType !== -1 ? (r[cType] || "").toString().trim() : "",
      level: cLevel !== -1 ? (r[cLevel] || "").toString().trim() : "",
      parentCode: cParent !== -1 ? extractParentCode((r[cParent] || "").toString().trim()) : "",
      description: cDesc !== -1 ? (r[cDesc] || "").toString().trim() : "",
      canPay: cPay !== -1 ? (r[cPay] || "").toString().trim() : "",
    });
  }
  return accounts;
}

async function readClientTreeInput(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
    const text = rows
      .filter((r) => Array.isArray(r) && r.some((c) => c !== null && String(c).trim() !== ""))
      .map((r) => r.map((c) => (c === null ? "" : String(c).trim())).join(" | "))
      .join("\n");
    return { kind: "text", text };
  }
  if (name.endsWith(".pdf")) {
    const buf = await file.arrayBuffer();
    return { kind: "pdf", base64: arrayBufferToBase64(buf) };
  }
  if (name.endsWith(".docx")) {
    const buf = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: buf });
    return { kind: "text", text: result.value };
  }
  throw new Error("صيغة الملف غير مدعومة. الصيغ المدعومة: Excel (.xlsx)، PDF، أو Word (.docx)");
}

function buildAccountsInstructions(existingAccounts) {
  const existingCodesText = existingAccounts.length
    ? existingAccounts.map((a) => `${a.code} — ${a.nameAr || a.nameEn}`).join("\n")
    : "(لا توجد حسابات افتراضية مرفقة)";

  return `أنت محاسب خبير بنظام "قيود" (Qoyod) متخصص ببناء أشجار الحسابات.

مهمتك: اقرأ شجرة حسابات العميل المرفقة (قد تكون بأي شكل: جدول، قائمة، تقرير PDF...) وحوّلها إلى شجرة حسابات متوافقة تمامًا مع نظام قيود، جاهزة للرفع في "قالب استيراد شجرة الحسابات".

قواعد قيود الثابتة (التزم بها حرفياً، تم التأكد منها من القالب الرسمي لقيود):
1. المستوى الأول (5 أنواع) نظامي وثابت بقيود تلقائياً — لا تُنشئ له صفوفاً في مخرجاتك، ابدأ مباشرة من المستوى الثاني: ${LEVEL1_TYPES.join("، ")}
2. المستوى الثاني: نوع الحساب يجب أن يكون واحداً بالضبط من هذه الـ12 نوع فقط: ${LEVEL2_TYPES.join("، ")}
3. المستوى الثالث وما فوق (حتى السابع): نوع الحساب يجب أن يكون واحداً بالضبط من هذه الـ59 نوع فقط: ${LEVEL3_TYPES.join("، ")}
4. أي حساب من المستوى الرابع حتى السابع يرث حرفياً نفس "نوع الحساب" الخاص بجدّه من المستوى الثالث.
5. نوع "رأس المال" لا يصح إلا ضمن حسابات حقوق الملاك (تحت: رأس المال المصدر / حقوق الملاك الأخرى / الأرباح المبقاة).
6. لا يوجد نوع "صافي الأصول" بقيود إطلاقاً — لحسابات صافي الأصول بالجمعيات غير الربحية استخدم "التزامات غير متداولة أخرى".
7. لا تكرر أي رمز موجود مسبقاً بالشجرة الافتراضية المذكورة أدناه.
8. أعطِ كل حساب رمزاً رقمياً هرمياً منطقياً يعكس المستوى والأب (مثال: 11 ثم 1101 ثم 110101).
9. "الحساب الرئيسي (الرمز)" = رمز الأب المباشر فقط (أعلى بمستوى واحد). للمستوى الثاني اتركه فارغاً.

الحسابات الموجودة افتراضياً بمنشأة العميل حالياً (لا تكرر أكوادها):
${existingCodesText}

أجب حصراً بصيغة JSON (مصفوفة كائنات) بدون أي نص أو Markdown إضافي، كل عنصر بهذا الشكل بالضبط:
{"code":"1101","nameAr":"النقدية والبنوك","nameEn":"Cash and Banks","level":2,"parentCode":"","type":"الأصول المتداولة","description":"","canPay":"No"}`;
}

async function callClaudeForAccounts(existingAccounts, clientTreeInput) {
  const instructions = buildAccountsInstructions(existingAccounts);
  let content;
  if (clientTreeInput.kind === "pdf") {
    content = [
      { type: "document", source: { type: "base64", media_type: "application/pdf", data: clientTreeInput.base64 } },
      { type: "text", text: instructions + "\n\nشجرة حسابات العميل مرفقة كملف PDF أعلاه." },
    ];
  } else {
    content = instructions + "\n\nشجرة حسابات العميل (نص مستخرج من الملف):\n" + clientTreeInput.text.slice(0, 60000);
  }
  const data = await callClaude({ model: "claude-sonnet-4-6", max_tokens: 8000, messages: [{ role: "user", content }] });
  return parseJsonResponse(data);
}

function validateAccountsList(accounts, existingCodesSet) {
  const byCode = {};
  accounts.forEach((a) => (byCode[a.code] = a));
  const issues = {};

  function ancestorAtLevel(acc, targetLevel) {
    let cur = acc;
    let guard = 0;
    while (cur && guard < 12) {
      if (Number(cur.level) === targetLevel) return cur;
      cur = byCode[cur.parentCode];
      guard++;
    }
    return null;
  }

  accounts.forEach((a) => {
    const rowIssues = [];
    const lvl = Number(a.level);
    if (!a.code) rowIssues.push("رمز الحساب مفقود");
    if (existingCodesSet.has(a.code)) rowIssues.push(`الرمز "${a.code}" مستخدم مسبقاً بالشجرة الافتراضية`);
    if (lvl === 2 && !LEVEL2_TYPES.includes(a.type)) rowIssues.push(`نوع غير صالح للمستوى الثاني: "${a.type}"`);
    if (lvl >= 3 && !LEVEL3_TYPES.includes(a.type)) rowIssues.push(`نوع غير صالح للمستوى الثالث فأعلى: "${a.type}"`);
    if (lvl >= 4) {
      const anc = ancestorAtLevel(a, 3);
      if (anc && anc.type !== a.type) rowIssues.push(`يجب أن يطابق نوع الحساب نوع جدّه بالمستوى الثالث ("${anc.type}")`);
    }
    if (a.type === "رأس المال") {
      const l2 = ancestorAtLevel(a, 2);
      if (!l2 || !EQUITY_LEVEL2_TYPES.includes(l2.type)) rowIssues.push(`نوع "رأس المال" يصح فقط ضمن حسابات حقوق الملاك`);
    }
    if (rowIssues.length) issues[a.code] = rowIssues;
  });
  return issues;
}

function buildAccountsPasteText(accounts) {
  return accounts
    .map((a) => [a.code, a.nameEn || "", a.nameAr || "", a.level, a.parentCode || "", a.type, a.description || "", a.canPay || "No"].join("\t"))
    .join("\n");
}

function UploadCard({ title, subtitle, fileName, ok, count, onFile, accept }) {
  const inputRef = useRef(null);
  return (
    <div className="rounded-lg border-2 border-dashed p-4 text-center transition cursor-pointer"
      style={{ borderColor: ok ? COLORS.green : COLORS.line, background: "#FFFDF7" }}
      onClick={() => inputRef.current?.click()}>
      <input ref={inputRef} type="file" accept={accept} className="hidden" onChange={(e) => e.target.files[0] && onFile(e.target.files[0])} />
      <div className="flex flex-col items-center gap-1">
        {ok ? <CheckCircle2 size={26} style={{ color: COLORS.green }} /> : <Upload size={26} style={{ color: COLORS.gold }} />}
        <p className="text-sm font-semibold">{title}</p>
        <p className="text-xs" style={{ color: "#5B5340" }}>{subtitle}</p>
        {fileName && <p className="mt-1 flex items-center gap-1 text-xs" style={{ color: COLORS.tealLight }}><FileSpreadsheet size={12} /> {fileName} {count && `— ${count}`}</p>}
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

export default function AccountsTool() {
  const [defaultTreeName, setDefaultTreeName] = useState("");
  const [defaultAccounts, setDefaultAccounts] = useState(null);
  const [clientTreeFile, setClientTreeFile] = useState(null);
  const [clientTreeName, setClientTreeName] = useState("");
  const [proposedAccounts, setProposedAccounts] = useState(null);
  const [analyzing, setAnalyzing] = useState(false);
  const [error, setError] = useState("");
  const [copyStatus, setCopyStatus] = useState("");
  const [showManualCopy, setShowManualCopy] = useState(false);

  const handleDefaultTreeUpload = async (file) => {
    setError(""); setDefaultTreeName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
      setDefaultAccounts(parseExistingTreeFile(rows));
    } catch (err) {
      setError("خطأ في قراءة الشجرة الافتراضية: " + err.message);
      setDefaultAccounts(null);
    }
  };

  const handleClientTreeUpload = (file) => {
    setError(""); setClientTreeName(file.name); setClientTreeFile(file); setProposedAccounts(null);
  };

  const runAnalysis = async () => {
    if (!clientTreeFile) return;
    setAnalyzing(true); setError("");
    try {
      const clientInput = await readClientTreeInput(clientTreeFile);
      const result = await callClaudeForAccounts(defaultAccounts || [], clientInput);
      setProposedAccounts(result);
    } catch (err) {
      setError("تعذر تحليل شجرة العميل: " + err.message);
    } finally {
      setAnalyzing(false);
    }
  };

  const existingCodesSet = useMemo(() => new Set((defaultAccounts || []).map((a) => a.code)), [defaultAccounts]);
  const issuesByCode = useMemo(() => (proposedAccounts ? validateAccountsList(proposedAccounts, existingCodesSet) : {}), [proposedAccounts, existingCodesSet]);
  const totalIssues = Object.values(issuesByCode).reduce((s, arr) => s + arr.length, 0);

  const updateAccount = (code, field, value) => {
    setProposedAccounts((prev) => prev.map((a) => (a.code === code ? { ...a, [field]: value } : a)));
  };
  const deleteAccount = (code) => {
    setProposedAccounts((prev) => prev.filter((a) => a.code !== code));
  };
  const addBlankAccount = () => {
    setProposedAccounts((prev) => [...(prev || []), { code: "", nameAr: "", nameEn: "", level: 2, parentCode: "", type: LEVEL2_TYPES[0], description: "", canPay: "No" }]);
  };

  const copyToClipboard = async () => {
    const text = buildAccountsPasteText(proposedAccounts);
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

  const downloadFile = () => {
    const header = ["الرمز", "الاسم الانجليزي", "الاسم العربي", "المستوى", "الحساب الرئيسي (الرمز)", "نوع الحساب", "الوصف", "يمكن الدفع والتحصيل بهذا الحساب"];
    const data = [header, ...proposedAccounts.map((a) => [a.code, a.nameEn || "", a.nameAr || "", a.level, a.parentCode || "", a.type, a.description || "", a.canPay || "No"])];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [{ wch: 14.4 }, { wch: 32.4 }, { wch: 27 }, { wch: 18 }, { wch: 30 }, { wch: 23.4 }, { wch: 14.4 }, { wch: 50 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Accounts Upload Template");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "شجرة_حسابات_جاهزة.xlsx";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div dir="rtl" className="min-h-screen w-full" style={{ background: COLORS.paper, color: COLORS.ink, fontFamily: "'Segoe UI', Tahoma, Arial, sans-serif" }}>
      <div className="mx-auto max-w-5xl px-4 py-6 sm:px-8">
        <div className="mb-6 flex items-start gap-2 rounded-lg border px-4 py-3 text-xs leading-relaxed" style={{ borderColor: COLORS.line, background: "#FFFDF7" }}>
          <Info size={16} className="mt-0.5 shrink-0" style={{ color: COLORS.gold }} />
          <div style={{ color: "#5B5340" }}>
            يقرأ الأداة شجرة حسابات العميل (أي صيغة: Excel، PDF، أو Word) ويحوّلها لشجرة متوافقة مع قيود: يبدأ من المستوى
            الثاني، يلتزم بالـ12 نوع للمستوى الثاني والـ59 نوع للثالث فأعلى، ويطبّق وراثة النوع من المستوى الثالث حتى
            السابع. <b>ملاحظة:</b> قيود نفسها لا تربط أنواع مستوى ثاني محددة بأنواع مستوى ثالث محددة (تأكدنا من ملف
            القالب الرسمي) — لو عندك جدول ربط دقيق إضافي أرسله لإضافته كفحص.
          </div>
        </div>

        <div className="mb-6 grid gap-4 sm:grid-cols-2">
          <UploadCard title="الشجرة الافتراضية الحالية بمنشأة العميل" subtitle="تصدير الحسابات الموجودة فعلياً بقيود حالياً (Excel)"
            fileName={defaultTreeName} ok={!!defaultAccounts} count={defaultAccounts ? `${defaultAccounts.length} حساب` : ""}
            onFile={handleDefaultTreeUpload} accept=".xlsx,.xls" />
          <UploadCard title="شجرة حسابات العميل" subtitle="أي صيغة: Excel، PDF، أو Word"
            fileName={clientTreeName} ok={!!clientTreeFile} count="" onFile={handleClientTreeUpload} accept=".xlsx,.xls,.pdf,.docx" />
        </div>

        {error && (
          <div className="mb-6 flex items-center gap-2 rounded-md border px-4 py-2 text-sm" style={{ borderColor: COLORS.red, color: COLORS.red }}>
            <XCircle size={16} /> {error}
          </div>
        )}

        {clientTreeFile && (
          <div className="mb-6">
            <button onClick={runAnalysis} disabled={analyzing}
              className="flex items-center gap-2 rounded-md px-4 py-2 text-sm font-medium text-white disabled:opacity-50" style={{ background: COLORS.teal }}>
              {analyzing ? <Loader2 size={16} className="animate-spin" /> : <Sparkles size={16} />}
              {analyzing ? "جاري التحليل والبناء..." : "تحليل وبناء الشجرة"}
            </button>
          </div>
        )}

        {proposedAccounts && (
          <>
            <div className="mb-6 grid grid-cols-3 gap-3">
              <SummaryStat label="إجمالي الحسابات المقترحة" value={proposedAccounts.length} color={COLORS.teal} />
              <SummaryStat label="حسابات سليمة" value={proposedAccounts.length - Object.keys(issuesByCode).length} color={COLORS.green} />
              <SummaryStat label="حسابات فيها ملاحظات" value={Object.keys(issuesByCode).length} color={totalIssues > 0 ? COLORS.red : COLORS.green} />
            </div>

            <div className="mb-3 flex justify-end">
              <button onClick={addBlankAccount} className="rounded-md border px-3 py-1.5 text-xs" style={{ borderColor: COLORS.line, color: COLORS.tealLight }}>
                + إضافة حساب يدوياً
              </button>
            </div>

            <div className="overflow-x-auto rounded-lg border" style={{ borderColor: COLORS.line, background: "#FFFDF7" }}>
              <table className="w-full text-xs">
                <thead>
                  <tr style={{ color: "#5B5340" }}>
                    <th className="p-2 text-right font-medium">الرمز</th>
                    <th className="p-2 text-right font-medium">الاسم العربي</th>
                    <th className="p-2 text-right font-medium">المستوى</th>
                    <th className="p-2 text-right font-medium">الحساب الرئيسي</th>
                    <th className="p-2 text-right font-medium">نوع الحساب</th>
                    <th className="p-2 text-right font-medium">الدفع والتحصيل</th>
                    <th className="p-2"></th>
                  </tr>
                </thead>
                <tbody>
                  {proposedAccounts.map((a) => {
                    const rowIssues = issuesByCode[a.code] || [];
                    return (
                      <React.Fragment key={a.code || Math.random()}>
                        <tr className="border-t" style={{ borderColor: COLORS.line, background: rowIssues.length ? "#FBEDEA" : undefined }}>
                          <td className="p-2 font-mono">
                            <input dir="ltr" value={a.code} onChange={(e) => updateAccount(a.code, "code", e.target.value)}
                              className="w-24 rounded border px-1.5 py-1 font-mono text-left" style={{ borderColor: COLORS.line }} />
                          </td>
                          <td className="p-2">
                            <input value={a.nameAr || ""} onChange={(e) => updateAccount(a.code, "nameAr", e.target.value)}
                              className="w-full min-w-[140px] rounded border px-1.5 py-1" style={{ borderColor: COLORS.line }} />
                          </td>
                          <td className="p-2">
                            <input dir="ltr" value={a.level} onChange={(e) => updateAccount(a.code, "level", e.target.value)}
                              className="w-14 rounded border px-1.5 py-1 text-center" style={{ borderColor: COLORS.line }} />
                          </td>
                          <td className="p-2">
                            <input dir="ltr" value={a.parentCode || ""} onChange={(e) => updateAccount(a.code, "parentCode", e.target.value)}
                              className="w-24 rounded border px-1.5 py-1 font-mono text-left" style={{ borderColor: COLORS.line }} />
                          </td>
                          <td className="p-2">
                            <select value={a.type} onChange={(e) => updateAccount(a.code, "type", e.target.value)}
                              className="w-full min-w-[160px] rounded border px-1.5 py-1" style={{ borderColor: COLORS.line }}>
                              {typesForLevel(a.level).map((t) => (<option key={t} value={t}>{t}</option>))}
                            </select>
                          </td>
                          <td className="p-2">
                            <select value={a.canPay || "No"} onChange={(e) => updateAccount(a.code, "canPay", e.target.value)}
                              className="rounded border px-1.5 py-1" style={{ borderColor: COLORS.line }}>
                              <option value="Yes">Yes</option>
                              <option value="No">No</option>
                            </select>
                          </td>
                          <td className="p-2">
                            <button onClick={() => deleteAccount(a.code)} className="rounded border px-2 py-1 text-xs" style={{ borderColor: COLORS.red, color: COLORS.red }}>حذف</button>
                          </td>
                        </tr>
                        {rowIssues.length > 0 && (
                          <tr>
                            <td colSpan={7} className="px-2 pb-2">
                              {rowIssues.map((msg, i) => (
                                <div key={i} className="flex items-center gap-1 text-xs" style={{ color: COLORS.red }}>
                                  <AlertTriangle size={12} /> {msg}
                                </div>
                              ))}
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="mt-8 flex flex-col items-center gap-2 border-t pt-6" style={{ borderColor: COLORS.line }}>
              {totalIssues > 0 && (
                <p className="flex items-center gap-1 text-xs" style={{ color: COLORS.amber }}>
                  <AlertTriangle size={14} /> ما زال هناك {totalIssues} ملاحظة — راجعها قبل الرفع لقيود
                </p>
              )}
              <div className="flex flex-wrap items-center justify-center gap-3">
                <button onClick={copyToClipboard} className="flex items-center gap-2 rounded-md px-6 py-2.5 text-sm font-semibold text-white" style={{ background: COLORS.green }}>
                  <Copy size={16} /> نسخ البيانات (للصق بالقالب الرسمي من الخلية A3)
                </button>
                <button onClick={downloadFile} className="flex items-center gap-2 rounded-md border px-5 py-2.5 text-sm font-semibold" style={{ borderColor: COLORS.line, color: COLORS.tealLight }}>
                  <Download size={16} /> أو تنزيل ملف مباشرة
                </button>
              </div>
              {copyStatus === "copied" && <span className="flex items-center gap-1 text-xs" style={{ color: COLORS.green }}><CheckCircle2 size={14} /> تم النسخ</span>}
              {showManualCopy && (
                <textarea readOnly dir="ltr" value={buildAccountsPasteText(proposedAccounts)} onFocus={(e) => e.target.select()}
                  className="mt-1 h-32 w-full max-w-md rounded border p-2 font-mono text-xs" style={{ borderColor: COLORS.line }} />
              )}
              <p className="mx-auto mt-1 max-w-md text-center text-[11px] leading-relaxed" style={{ color: "#8A8163" }}>
                للحفاظ على تنسيق القالب الرسمي (الأوراق المخفية والقوائم المنسدلة) بالكامل: افتح نسخة فارغة من "قالب استيراد
                شجرة الحسابات" الرسمي، اضغط الخلية A3، والصق.
              </p>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
