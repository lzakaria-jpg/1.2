import * as XLSX from "xlsx";

// ---------- low-level helpers ----------
export function normalizeDateGuess(raw) {
  if (raw === null || raw === undefined || raw === "") return "";
  const s = String(raw).trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[3].padStart(2, "0")}/${m[2].padStart(2, "0")}/${m[1]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[1].padStart(2, "0")}/${m[2].padStart(2, "0")}/${m[3]}`;
  if (/^\d+(\.\d+)?$/.test(s) && Number(s) > 20000 && Number(s) < 80000) {
    const d = XLSX.SSF.parse_date_code(Number(s));
    if (d) return `${String(d.d).padStart(2, "0")}/${String(d.m).padStart(2, "0")}/${d.y}`;
  }
  return s;
}

export function normalizeCode(raw) {
  if (raw === null || raw === undefined || raw === "") return "";
  if (typeof raw === "number") return String(raw);
  let s = String(raw).trim();
  s = s.replace(/\.0$/, "");
  return s;
}

export function parseAmount(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number") return raw;
  const s = String(raw).replace(/,/g, "").trim();
  if (s === "" || s === "-") return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function cellText(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

// Reads an uploaded file into a plain array-of-arrays (1 row = 1 array, 0-indexed columns).
// Uses SheetJS rather than ExcelJS here: real-world exports (e.g. some accounting-system
// reports) sometimes use non-default XML namespace prefixes in workbook.xml, which ExcelJS's
// parser fails on (throws "Cannot read properties of undefined (reading 'sheets')") but SheetJS
// handles fine. ExcelJS is used later only for writing the final file (see excelExport.js),
// where we control the source template and don't hit this issue.
export async function readWorkbookRows(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
}

// ---------- AI semantic review (journal entries) ----------
export function buildChartText(chartAccounts, parentCodes) {
  return chartAccounts
    .map((a) => {
      const flag = parentCodes.has(a.code) ? " | [حساب رئيسي - غير قابل للترحيل المباشر]" : "";
      return `${a.code} | ${a.name} | نوع: ${a.type}${a.parentCode ? " | تابع لـ: " + a.parentCode : ""}${flag}`;
    })
    .join("\n");
}

export function buildEntriesPromptText(entries) {
  return entries
    .map((e) => {
      const rowsText = e.rows
        .map((r, i) => `  سطر ${i}: رمز=${r.code || "فارغ"} مدين=${r.debit ?? ""} دائن=${r.credit ?? ""} تعليق=${r.comment || ""}`)
        .join("\n");
      return `قيد رقم ${e.seq} — التاريخ: ${e.date || ""} — الوصف: ${e.desc || ""}\n${rowsText}`;
    })
    .join("\n\n");
}

export function buildSemanticsPrompt(entriesBatch, chartAccounts, parentCodes) {
  return `أنت محاسب خبير بنظام "قيود" (Qoyod) تراجع قيود يومية مقابل شجرة حسابات عميل معيّن.
مهمتك: افحص كل سطر في كل قيد، وحدد إن كان رمز الحساب المستخدم يتوافق منطقياً مع وصف القيد وطبيعة العملية — حتى لو كان الرمز موجوداً فعلياً في الشجرة.
مهم جداً: الحسابات المعلّمة بـ "[حساب رئيسي - غير قابل للترحيل المباشر]" ممنوع نهائياً اقتراحها كبديل.
إن لم تجد بديلاً مناسباً بثقة، لا تقترح شيئاً واترك hasIssue بقيمة true مع suggestedCode فارغ.
لا تُبلغ عن الأخطاء الهيكلية (رمز غير موجود، عدم توازن، ترحيل على حساب رئيسي) فهذه تُفحص برمجياً بشكل منفصل.

شجرة الحسابات (رمز | اسم | نوع):
${buildChartText(chartAccounts, parentCodes)}

القيود المطلوب مراجعتها (رقم السطر يبدأ من صفر داخل كل قيد):
${buildEntriesPromptText(entriesBatch)}

أجب بصيغة JSON فقط بدون أي نص أو Markdown إضافي، وفق الشكل التالي بالضبط (مصفوفة فارغة إن لم تجد أي مشكلة):
[{"seq": "1", "rowIndex": 0, "hasIssue": true, "currentCode": "110103", "suggestedCode": "110402", "suggestedName": "إيجار مقدم", "reason": "سبب موجز بالعربي"}]`;
}

function findHeaderRowIndex(rows, mustInclude) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    if (row.some((c) => cellText(c).trim() === mustInclude)) return i;
  }
  return -1;
}

function colIndex(headerRow, ...candidates) {
  for (const cand of candidates) {
    const idx = headerRow.findIndex((h) => cellText(h).includes(cand));
    if (idx !== -1) return idx;
  }
  return -1;
}

// ---------- chart of accounts ----------
export function parseChartFile(rows) {
  const hIdx = findHeaderRowIndex(rows, "الرمز");
  if (hIdx === -1) throw new Error("لم يتم العثور على عمود 'الرمز' في ملف شجرة الحسابات");
  const header = rows[hIdx].map(cellText);
  const cCode = colIndex(header, "الرمز");
  const cName = colIndex(header, "اسم الحساب");
  const cType = colIndex(header, "النوع");
  const cDesc = colIndex(header, "الوصف");
  const cParent = colIndex(header, "Parent", "الحساب الأب");
  const cPay = colIndex(header, "الدفع والتحصيل", "يمكن الدفع");

  const accounts = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const code = normalizeCode(r[cCode]);
    if (!code) continue;
    accounts.push({
      code,
      name: cName !== -1 ? cellText(r[cName]).trim() : "",
      type: cType !== -1 ? cellText(r[cType]).trim() : "",
      description: cDesc !== -1 ? cellText(r[cDesc]).trim() : "",
      parentCode: cParent !== -1 ? cellText(r[cParent]).trim() : "",
      canPay: cPay !== -1 ? cellText(r[cPay]).trim() : "",
    });
  }
  return accounts;
}

export function extractParentCode(raw) {
  if (!raw) return "";
  const s = String(raw).trim();
  if (!s) return "";
  return s.split(/\s+/)[0].replace(/\.0$/, "");
}

export function buildParentInfo(chartAccounts) {
  const parentCodes = new Set();
  const childrenByParent = {};
  chartAccounts.forEach((a) => {
    const pc = extractParentCode(a.parentCode);
    if (!pc) return;
    parentCodes.add(pc);
    if (!childrenByParent[pc]) childrenByParent[pc] = [];
    childrenByParent[pc].push(a);
  });
  // Root/category-level accounts (e.g. Level 1 "الأصول") carry no parent code of their own,
  // so nothing above ever lists them as a parent via the loop above — but they are never valid
  // posting targets either. Any account with a blank parent code is itself a category header.
  chartAccounts.forEach((a) => {
    if (!extractParentCode(a.parentCode)) parentCodes.add(a.code);
  });
  return { parentCodes, childrenByParent };
}

// ---------- journal entries: two supported input schemas ----------
function parseTemplateSchema(rows, hIdx) {
  const header = rows[hIdx].map(cellText);
  const cSeq = colIndex(header, "تسلسل القيد");
  const cDate = colIndex(header, "التاريخ");
  const cDesc = colIndex(header, "وصف القيد");
  const cAccType = colIndex(header, "نوع الحساب");
  const cCode = colIndex(header, "رمز الحساب");
  const cContact = colIndex(header, "جهة اتصال");
  const cDebit = colIndex(header, "مدين");
  const cCredit = colIndex(header, "دائن");
  const cComment = colIndex(header, "التعليقات");

  const flat = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    flat.push({
      seq: cellText(r[cSeq]).trim(),
      date: cDate !== -1 ? normalizeDateGuess(r[cDate]) : "",
      desc: cDesc !== -1 ? cellText(r[cDesc]).trim() : "",
      accType: cAccType !== -1 ? cellText(r[cAccType]).trim() : "حسابات دفتر الاستاذ",
      code: cCode !== -1 ? normalizeCode(r[cCode]) : "",
      contact: cContact !== -1 ? cellText(r[cContact]).trim() : "",
      debit: cDebit !== -1 ? parseAmount(r[cDebit]) : null,
      credit: cCredit !== -1 ? parseAmount(r[cCredit]) : null,
      comment: cComment !== -1 ? cellText(r[cComment]).trim() : "",
    });
  }
  return groupEntries(flat);
}

function parseRawLedgerSchema(rows, hIdx) {
  const header = rows[hIdx].map(cellText);
  const cOp = colIndex(header, "رقم العملية", "رقم القيد");
  const cDate = colIndex(header, "تاريخ");
  const cCode = colIndex(header, "رمز الحساب");
  const cDesc = colIndex(header, "تعريف", "وصف القيد", "البيان");
  const cDebit = colIndex(header, "مدين");
  const cCredit = colIndex(header, "دائن");
  const cNotes = colIndex(header, "ملاحظ");
  const cCCCode = colIndex(header, "رمز مركز");
  const cCCName = colIndex(header, "اسم مركز");

  const groups = [];
  let current = null;
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const op = cOp !== -1 ? cellText(r[cOp]).trim() : "";
    const code = cCode !== -1 ? normalizeCode(r[cCode]) : "";
    const descRaw = cDesc !== -1 ? cellText(r[cDesc]).trim() : "";

    if (descRaw.includes("اجمالي")) continue;
    if (!op && !code) continue;
    if (!op) continue;

    if (!current || current.seq !== op) {
      current = { seq: op, date: cDate !== -1 ? normalizeDateGuess(r[cDate]) : "", desc: descRaw, rows: [] };
      groups.push(current);
    }
    if (!current.date && cDate !== -1) {
      const d = normalizeDateGuess(r[cDate]);
      if (d) current.date = d;
    }
    if (!current.desc && descRaw) current.desc = descRaw;

    const ccCode = cCCCode !== -1 ? cellText(r[cCCCode]).trim() : "";
    const ccName = cCCName !== -1 ? cellText(r[cCCName]).trim() : "";
    const notes = cNotes !== -1 ? cellText(r[cNotes]).trim() : "";
    let comment = notes;
    if (ccCode) comment = (comment ? comment + " | " : "") + `مركز التكلفة: ${ccCode}${ccName ? " - " + ccName : ""}`;

    current.rows.push({
      seq: op,
      date: current.date,
      desc: current.desc,
      accType: "حسابات دفتر الاستاذ",
      code,
      contact: "",
      debit: cDebit !== -1 ? parseAmount(r[cDebit]) : null,
      credit: cCredit !== -1 ? parseAmount(r[cCredit]) : null,
      comment,
      _rowIndex: i,
    });
  }
  return groups;
}

export function parseEntriesFile(rows) {
  let hIdx = findHeaderRowIndex(rows, "تسلسل القيد");
  if (hIdx !== -1) return parseTemplateSchema(rows, hIdx);

  hIdx = findHeaderRowIndex(rows, "رقم العملية");
  if (hIdx !== -1 && colIndex(rows[hIdx].map(cellText), "رمز الحساب") !== -1) {
    return parseRawLedgerSchema(rows, hIdx);
  }

  const candidateRow =
    rows.find((r) => Array.isArray(r) && r.some((c) => cellText(c).includes("رمز الحساب"))) ||
    rows.slice(0, 8).find((r) => Array.isArray(r) && r.filter((c) => cellText(c).trim() !== "").length >= 3) ||
    rows[0] ||
    [];
  const foundCols = candidateRow.map(cellText).filter((c) => c.trim() !== "").join("، ");
  throw new Error(
    `لم يتم التعرف على تنسيق ملف القيود.` +
      (foundCols ? ` الأعمدة الموجودة بالملف: ${foundCols}.` : "") +
      ` الصيغ المدعومة حالياً: (1) قالب استيراد القيود الرسمي لقيود، أو (2) تقرير قيود اليومية.`
  );
}

function groupEntries(flatRows) {
  const groups = [];
  let current = null;
  flatRows.forEach((r, idx) => {
    const isFullyEmpty = !r.seq && !r.date && !r.desc && !r.code && r.debit === null && r.credit === null && !r.comment;
    if (isFullyEmpty) {
      current = null;
      return;
    }
    if (r.seq) {
      if (!current || current.seq !== r.seq) {
        current = { seq: r.seq, date: r.date, desc: r.desc, rows: [] };
        groups.push(current);
      }
    }
    if (!current) {
      current = { seq: `?${idx}`, date: r.date, desc: r.desc, rows: [] };
      groups.push(current);
    }
    if (!current.date && r.date) current.date = r.date;
    if (!current.desc && r.desc) current.desc = r.desc;
    current.rows.push({ ...r, _rowIndex: idx });
  });
  return groups;
}

// ---------- validation ----------
export function validateEntryStructure(entry, chartMap, parentInfo) {
  const issues = [];
  const totalDebit = entry.rows.reduce((s, r) => s + (parseFloat(r.debit) || 0), 0);
  const totalCredit = entry.rows.reduce((s, r) => s + (parseFloat(r.credit) || 0), 0);

  if (Math.abs(totalDebit - totalCredit) > 0.005) {
    issues.push({
      id: `${entry.seq}-balance`,
      type: "unbalanced",
      severity: "error",
      message: `القيد غير متزن: مجموع المدين ${totalDebit.toLocaleString()} لا يساوي مجموع الدائن ${totalCredit.toLocaleString()}`,
    });
  }
  if (!entry.date || !/^\d{2}\/\d{2}\/\d{4}$/.test(entry.date)) {
    issues.push({
      id: `${entry.seq}-date`,
      type: "date_format",
      severity: "error",
      message: `تاريخ القيد مفقود أو غير مطابق لصيغة dd/mm/yyyy (القيمة الحالية: "${entry.date || "فارغ"}")`,
    });
  }
  if (!entry.desc) {
    issues.push({ id: `${entry.seq}-desc`, type: "missing_desc", severity: "error", message: "وصف القيد مفقود في السطر الأول" });
  }
  if (entry.rows.length < 2) {
    issues.push({
      id: `${entry.seq}-rows`,
      type: "too_few_rows",
      severity: "error",
      message: "القيد يحتوي على سطر واحد فقط، يجب أن يحتوي على سطرين على الأقل (مدين ودائن)",
    });
  }

  entry.rows.forEach((r, i) => {
    const hasDebit = r.debit !== null && !isNaN(r.debit) && r.debit > 0;
    const hasCredit = r.credit !== null && !isNaN(r.credit) && r.credit > 0;
    if (hasDebit && hasCredit) {
      issues.push({
        id: `${entry.seq}-row${i}-both`,
        type: "both_amounts",
        severity: "error",
        rowIndex: r._rowIndex,
        message: `السطر ${i + 1}: لا يمكن تعبئة خانتي مدين ودائن معاً بنفس السطر`,
      });
    }
    if (!hasDebit && !hasCredit) {
      issues.push({
        id: `${entry.seq}-row${i}-none`,
        type: "no_amount",
        severity: "error",
        rowIndex: r._rowIndex,
        message: `السطر ${i + 1}: لم يتم إدخال أي مبلغ (لا مدين ولا دائن)`,
      });
    }
    if (!r.code || !chartMap[r.code]) {
      issues.push({
        id: `${entry.seq}-row${i}-unknown`,
        type: "unknown_code",
        severity: "error",
        rowIndex: r._rowIndex,
        code: r.code,
        message: `السطر ${i + 1}: رمز الحساب "${r.code || "فارغ"}" غير موجود في شجرة الحسابات المرفقة`,
      });
    } else if (parentInfo.parentCodes.has(r.code)) {
      const children = (parentInfo.childrenByParent[r.code] || []).slice(0, 6);
      const childrenText = children.map((c) => `${c.code} ${c.name}`).join("، ");
      issues.push({
        id: `${entry.seq}-row${i}-parent`,
        type: "parent_account",
        severity: "error",
        rowIndex: r._rowIndex,
        code: r.code,
        message: `السطر ${i + 1}: الحساب "${r.code} — ${chartMap[r.code].name}" حساب رئيسي وله حسابات فرعية، لا يمكن ترحيل قيد عليه مباشرة في قيود.${
          childrenText ? ` اختر أحد الحسابات الفرعية: ${childrenText}` : ""
        }`,
      });
    }
  });

  return issues;
}
