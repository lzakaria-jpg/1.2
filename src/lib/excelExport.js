import ExcelJS from "exceljs";

// Loads the real official blank template bundled with the app (public/qoyod_template.xlsx)
async function loadTemplateWorkbook() {
  const res = await fetch(`${import.meta.env.BASE_URL}qoyod_template.xlsx`);
  if (!res.ok) throw new Error("تعذر تحميل قالب قيود الرسمي المرفق بالتطبيق");
  const buf = await res.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

function findHeaderRowNumber(worksheet, mustInclude) {
  let found = 1;
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (typeof cell.value === "string" && cell.value.trim() === mustInclude) found = rowNumber;
    });
  });
  return found;
}

// Builds the final .xlsx as a Blob, with entries written straight into the real template's
// own cells (header row and column widths are never touched, so they survive untouched).
export async function buildImportFile(entries) {
  const wb = await loadTemplateWorkbook();
  const ws = wb.worksheets[0];
  const headerRowNum = findHeaderRowNumber(ws, "تسلسل القيد");
  let r = headerRowNum + 1;

  entries.forEach((entry, idx) => {
    entry.rows.forEach((row, i) => {
      const vals = [
        entry.seq,
        i === 0 ? entry.date : "",
        i === 0 ? entry.desc : "",
        row.accType || "حسابات دفتر الاستاذ",
        row.code || "",
        row.contact || "",
        row.debit || "",
        row.credit || "",
        row.comment || "",
      ];
      vals.forEach((val, cIdx) => {
        const cell = ws.getCell(r, cIdx + 1);
        cell.value = val === "" || val === null || val === undefined ? null : val;
      });
      r++;
    });
    if (idx < entries.length - 1) r++; // mandatory blank separator row between journal entries
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Same 9 columns, tab-separated, blank line between entries — kept as a fallback/manual option.
export function buildPasteText(entries) {
  const lines = [];
  entries.forEach((entry, idx) => {
    entry.rows.forEach((r, i) => {
      lines.push(
        [
          entry.seq,
          i === 0 ? entry.date : "",
          i === 0 ? entry.desc : "",
          r.accType || "حسابات دفتر الاستاذ",
          r.code || "",
          r.contact || "",
          r.debit ?? "",
          r.credit ?? "",
          r.comment || "",
        ].join("\t")
      );
    });
    if (idx < entries.length - 1) lines.push(new Array(9).fill("").join("\t"));
  });
  return lines.join("\n");
}
