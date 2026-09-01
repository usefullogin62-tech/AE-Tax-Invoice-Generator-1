import * as ExcelJS from "exceljs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

import { CERTIFICATE_LINES, FIRM, computeTotals, inr, type Invoice } from "./invoice-model";

const fileBase = (inv: Invoice) =>
  `Tax_Invoice_WO_${inv.workOrderNo || "unknown"}${inv.reBillNo ? `_RE_${inv.reBillNo}` : ""}`;

interface Row {
  sr: string | number;
  desc: string;
  unit: string;
  qty: number | "";
  rate: number | "";
  amount: number | "";
}

function bodyRows(inv: Invoice): Row[] {
  const rows: Row[] = [];
  rows.push({
    sr: "Material Supplied by Agency",
    desc: "",
    unit: "",
    qty: "",
    rate: "",
    amount: "",
  });
  inv.materials.forEach((m, i) =>
    rows.push({
      sr: i + 1,
      desc: m.description,
      unit: m.unit,
      qty: m.qty,
      rate: m.rate,
      amount: m.qty * m.rate,
    }),
  );
  rows.push({
    sr: `Transportation Charges @ ${inv.transportPct}%`,
    desc: "",
    unit: "",
    qty: "",
    rate: "",
    amount: inv.transportAmount,
  });
  rows.push({
    sr: `Insurance Charges @ ${inv.insurancePct}%`,
    desc: "",
    unit: "",
    qty: "",
    rate: "",
    amount: inv.insuranceAmount,
  });
  rows.push({
    sr: "Services Supplied by Agency",
    desc: "",
    unit: "",
    qty: "",
    rate: "",
    amount: "",
  });
  inv.services.forEach((s, i) =>
    rows.push({
      sr: i + 1,
      desc: s.description,
      unit: s.unit,
      qty: s.qty,
      rate: s.rate,
      amount: s.qty * s.rate,
    }),
  );
  rows.push({
    sr: `Rate quoted by Agency ${inv.ratePct}% Above`,
    desc: "",
    unit: "",
    qty: "",
    rate: inv.ratePct,
    amount: inv.rateAmount,
  });
  return rows;
}

const isSectionRow = (r: Row) => r.desc === "" && typeof r.sr === "string";

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ------------------------------- Excel -------------------------------- */
/* Enterprise-formatted workbook built with ExcelJS (fonts, fills, borders,
 * number formats, freeze panes) with print margins reserved top & bottom
 * so the sheet prints cleanly onto pre-printed letterhead stationery. */

const ACCENT_HEX = "0F4C81";
const BAND_HEX = "F1F5F9";
const ZEBRA_HEX = "F8FAFC";
const LINE_HEX = "94A3B8";
const INK_HEX = "111827";

const thin = { style: "thin" as const, color: { argb: `FF${LINE_HEX}` } };
const boxBorder = { top: thin, left: thin, bottom: thin, right: thin };

type Cell = ExcelJS.Cell;

function styleRange(
  ws: ExcelJS.Worksheet,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  apply: (cell: Cell) => void,
) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      apply(ws.getCell(r, c));
    }
  }
}

export async function exportExcel(inv: Invoice) {
  const wb = new ExcelJS.Workbook();
  wb.creator = FIRM.name;
  wb.created = new Date();
  buildInvoiceSheet(wb, inv, inv.workOrderNo || "Tax Invoice");
  const buf = await wb.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `${fileBase(inv)}.xlsx`,
  );
}

/** Bulk export: one workbook, one sheet per invoice (max 10 POs at a time). */
export async function exportExcelBulk(invoices: Invoice[]) {
  const wb = new ExcelJS.Workbook();
  wb.creator = FIRM.name;
  wb.created = new Date();
  const usedNames = new Set<string>();
  invoices.forEach((inv, i) => {
    let name =
      (inv.reBillNo ? `RE ${inv.reBillNo}` : inv.workOrderNo || `Invoice ${i + 1}`)
        .replace(/[\\/*?:[\]]/g, "-")
        .slice(0, 28)
        .trim() || `Invoice ${i + 1}`;
    while (usedNames.has(name)) name = `${name.slice(0, 25)} (${i + 1})`;
    usedNames.add(name);
    buildInvoiceSheet(wb, inv, name);
  });
  const buf = await wb.xlsx.writeBuffer();
  const first = invoices[0];
  const label = first ? `${first.section || "Bulk"}_${invoices.length}_Invoices` : "Bulk_Invoices";
  downloadBlob(
    new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `Tax_Invoices_${label.replace(/[^a-zA-Z0-9_-]/g, "_")}.xlsx`,
  );
}

function buildInvoiceSheet(wb: ExcelJS.Workbook, inv: Invoice, sheetName: string) {
  const t = computeTotals(inv);
  const ws = wb.addWorksheet(sheetName, {
    pageSetup: {
      paperSize: 9, // A4
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      // Generous top/bottom margins so the printer leaves the pre-printed
      // letterhead's header & footer artwork untouched.
      margins: { top: 1.7, bottom: 1.1, left: 0.45, right: 0.45, header: 0.3, footer: 0.3 },
      horizontalCentered: true,
    },
    headerFooter: {
      oddFooter: "&L&8This is a system-generated tax invoice.&R&8Page &P of &N",
    },
  });

  ws.columns = [
    { width: 9 },
    { width: 52 },
    { width: 9 },
    { width: 9 },
    { width: 16 },
    { width: 16 },
  ];

  let r = 1;

  // ---- Title -------------------------------------------------------
  ws.mergeCells(r, 1, r, 6);
  const title = ws.getCell(r, 1);
  title.value = "TAX INVOICE";
  title.font = { bold: true, size: 16, color: { argb: `FF${INK_HEX}` } };
  title.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(r).height = 24;
  r++;
  ws.mergeCells(r, 1, r, 6);
  const woTag = ws.getCell(r, 1);
  woTag.value = `WO No. ${inv.workOrderNo || "-"}`;
  woTag.font = { size: 8, italic: true, color: { argb: `FF${LINE_HEX}` } };
  woTag.alignment = { horizontal: "right" };
  r++;

  // ---- Info box (bordered, two columns) -----------------------------
  const boxTop = r;
  const leftPairs: [string, string][] = [
    ["To,", "Executive Engineer"],
    ["Section :-", inv.section],
    ["Sub Division :-", inv.subDivision],
    ["Division :-", `M.S.E.D.C.L O & M ${inv.division}`],
    ["Work Order :-", `EE/${inv.division}/LOE :- ${inv.loeNo}  Dt. ${inv.loeDate}`],
  ];
  const rightPairs: [string, string][] = [
    ["RE Bill No :", inv.reBillNo],
    ["RA Bill Date :", inv.raBillDate],
    ["Work Order No :", inv.workOrderNo],
    ["W.O. Date :", inv.workOrderDate],
    ["GSTIN :", FIRM.gstin],
  ];
  leftPairs.forEach(([label, value], i) => {
    const row = boxTop + i;
    ws.mergeCells(row, 1, row, 4);
    const c = ws.getCell(row, 1);
    c.value = {
      richText: [
        { font: { bold: true, size: 9 }, text: `${label} ` },
        { font: { size: 9 }, text: value },
      ],
    };
    c.alignment = { vertical: "middle", wrapText: true };
    ws.getRow(row).height = 15;
  });
  rightPairs.forEach(([label, value], i) => {
    const row = boxTop + i;
    ws.mergeCells(row, 5, row, 6);
    const c = ws.getCell(row, 5);
    c.value = {
      richText: [
        { font: { bold: true, size: 9 }, text: `${label} ` },
        { font: { size: 9 }, text: String(value) },
      ],
    };
    c.alignment = { vertical: "middle", wrapText: true };
  });
  const boxBottom = boxTop + leftPairs.length - 1;
  styleRange(ws, boxTop, 1, boxBottom, 6, (cell) => {
    cell.border = { ...boxBorder };
  });
  // divider between the two columns
  styleRange(ws, boxTop, 4, boxBottom, 4, (cell) => {
    cell.border = { ...cell.border, right: thin };
  });
  r = boxBottom + 2;

  // ---- Table header --------------------------------------------------
  const headerRow = r;
  ["Sr.No", "Description", "Unit", "Qty", "Rate", "Amount (Rs.)"].forEach((h, i) => {
    const c = ws.getCell(headerRow, i + 1);
    c.value = h;
    c.font = { bold: true, size: 9.5, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${ACCENT_HEX}` } };
    c.alignment = { horizontal: i === 1 ? "left" : "center", vertical: "middle" };
    c.border = { ...boxBorder };
  });
  ws.getRow(headerRow).height = 18;
  ws.views = [{ state: "frozen", ySplit: headerRow }];
  r++;

  let zebraIdx = 0;
  for (const row of bodyRows(inv)) {
    if (isSectionRow(row)) {
      const hasAmount = row.amount !== "";
      ws.mergeCells(r, 1, r, hasAmount ? 5 : 6);
      const c = ws.getCell(r, 1);
      c.value = String(row.sr);
      c.font = { bold: true, size: 9, color: { argb: `FF${ACCENT_HEX}` } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${BAND_HEX}` } };
      c.alignment = { vertical: "middle" };
      if (hasAmount) {
        const amt = ws.getCell(r, 6);
        amt.value = Number(row.amount);
        amt.numFmt = "#,##0.00";
        amt.font = { bold: true, size: 9 };
        amt.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${BAND_HEX}` } };
        amt.alignment = { horizontal: "right", vertical: "middle" };
      }
      styleRange(ws, r, 1, r, 6, (cell) => (cell.border = { ...boxBorder }));
    } else {
      const vals = [
        row.sr,
        row.desc,
        row.unit,
        Number(row.qty),
        Number(row.rate),
        Number(row.qty) * Number(row.rate),
      ];
      vals.forEach((v, i) => {
        const c = ws.getCell(r, i + 1);
        c.value = v as string | number;
        c.font = { size: 9 };
        c.border = { ...boxBorder };
        c.alignment = {
          horizontal: i === 1 ? "left" : i >= 3 ? "right" : "center",
          vertical: "middle",
          wrapText: i === 1,
        };
        if (i === 4 || i === 5) c.numFmt = "#,##0.00";
        if (zebraIdx % 2 === 1) {
          c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${ZEBRA_HEX}` } };
        }
      });
      zebraIdx++;
    }
    r++;
  }

  // ---- Totals ----------------------------------------------------------
  const totalLines: [string, number][] = [
    ["TOTAL (Excl. Taxes)", t.totalExclusive],
    ["G.S.T. 18%", t.gst],
    ["GRAND TOTAL (Incl. Taxes)", t.totalInclusive],
  ];
  totalLines.forEach(([label, val], i) => {
    const isGrand = i === totalLines.length - 1;
    ws.mergeCells(r, 1, r, 5);
    const lbl = ws.getCell(r, 1);
    lbl.value = label;
    lbl.font = { bold: true, size: 9.5, color: { argb: isGrand ? "FFFFFFFF" : `FF${INK_HEX}` } };
    lbl.alignment = { horizontal: "right", vertical: "middle" };
    const amt = ws.getCell(r, 6);
    amt.value = val;
    amt.numFmt = "#,##0.00";
    amt.font = { bold: true, size: 9.5, color: { argb: isGrand ? "FFFFFFFF" : `FF${INK_HEX}` } };
    amt.alignment = { horizontal: "right", vertical: "middle" };
    styleRange(ws, r, 1, r, 6, (cell) => {
      cell.border = { ...boxBorder };
      if (isGrand)
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${ACCENT_HEX}` } };
    });
    r++;
  });
  r++;

  // ---- Amount in words + GST/PAN ---------------------------------------
  ws.mergeCells(r, 1, r, 6);
  ws.getCell(r, 1).value = `Amount in words: ${inv.amountInWords}`;
  ws.getCell(r, 1).font = { bold: true, size: 9.5 };
  r++;
  ws.mergeCells(r, 1, r, 6);
  ws.getCell(r, 1).value = `GST No: ${FIRM.gstin}      PAN No: ${FIRM.pan}`;
  ws.getCell(r, 1).font = { size: 9.5 };
  r += 2;

  // ---- Signature block (comes right after totals, as per format) -------
  ws.mergeCells(r, 5, r, 6);
  const forFirm = ws.getCell(r, 5);
  forFirm.value = `For ${FIRM.name}`;
  forFirm.font = { bold: true, size: 9.5 };
  forFirm.alignment = { horizontal: "right" };
  r += 3;
  ws.mergeCells(r, 5, r, 6);
  const sig = ws.getCell(r, 5);
  sig.value = "Authorised Signatory";
  sig.font = { size: 9.5 };
  sig.alignment = { horizontal: "right" };
  r++;

  // ---- 5-6 line gap, then CERTIFICATE below the signature ---------------
  r += 6;

  ws.mergeCells(r, 1, r, 6);
  const certHead = ws.getCell(r, 1);
  certHead.value = "CERTIFICATE";
  certHead.font = { bold: true, size: 10.5, color: { argb: `FF${ACCENT_HEX}` } };
  certHead.border = { bottom: { style: "medium", color: { argb: `FF${ACCENT_HEX}` } } };
  r++;

  CERTIFICATE_LINES.forEach((line, i) => {
    ws.getCell(r, 1).value = `${i + 1})`;
    ws.getCell(r, 1).font = { bold: true, size: 9 };
    ws.getCell(r, 1).alignment = { vertical: "top" };
    ws.mergeCells(r, 2, r, 6);
    const c = ws.getCell(r, 2);
    c.value = line;
    c.font = { size: 9 };
    c.alignment = { wrapText: true, vertical: "top" };
    ws.getRow(r).height = 14;
    r++;
  });

  ws.pageSetup.printArea = `A1:F${r}`;
}

/* -------------------------------- PDF --------------------------------- */

export function exportPdf(inv: Invoice) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  drawInvoicePage(doc, inv);
  stampFooters(doc);
  doc.save(`${fileBase(inv)}.pdf`);
}

/** Bulk export: one merged PDF, each invoice starting on its own page (max 10 POs at a time). */
export function exportPdfBulk(invoices: Invoice[]) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  invoices.forEach((inv, i) => {
    if (i > 0) doc.addPage();
    drawInvoicePage(doc, inv);
  });
  stampFooters(doc);
  const first = invoices[0];
  const label = first ? `${first.section || "Bulk"}_${invoices.length}_Invoices` : "Bulk_Invoices";
  doc.save(`Tax_Invoices_${label.replace(/[^a-zA-Z0-9_-]/g, "_")}.pdf`);
}

function stampFooters(doc: jsPDF) {
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const M = 34;
  const BOTTOM = 70;
  const LINE: [number, number, number] = [148, 163, 184];
  const pageCount = doc.getNumberOfPages();
  for (let p = 1; p <= pageCount; p++) {
    doc.setPage(p);
    doc.setDrawColor(...LINE).setLineWidth(0.4);
    doc.line(M, H - BOTTOM + 12, W - M, H - BOTTOM + 12);
    doc.setFont("helvetica", "italic").setFontSize(7).setTextColor(120, 120, 120);
    doc.text("This is a system-generated tax invoice.", M, H - BOTTOM + 24);
    doc.setFont("helvetica", "normal");
    doc.text(`Page ${p} of ${pageCount}`, W - M, H - BOTTOM + 24, { align: "right" });
  }
}

/** Draws one invoice starting at the current page's top; adds internal pages
 * of its own (via ensureSpace) if a single invoice overflows one page. */
function drawInvoicePage(doc: jsPDF, inv: Invoice) {
  const t = computeTotals(inv);
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const M = 34;
  const TOP = 120; // blank space reserved for pre-printed letterhead
  const BOTTOM = 70; // blank space reserved for footer

  const INK: [number, number, number] = [17, 24, 39];
  const ACCENT: [number, number, number] = [15, 76, 129];
  const LINE: [number, number, number] = [148, 163, 184];
  const BAND: [number, number, number] = [241, 245, 249];
  const ZEBRA: [number, number, number] = [248, 250, 252];

  let y = TOP;

  // ---- Title + invoice ref tag -----------------------------------------
  doc
    .setFont("helvetica", "bold")
    .setFontSize(16)
    .setTextColor(...INK);
  doc.text("TAX INVOICE", W / 2, y, { align: "center" });
  doc.setFillColor(...ACCENT);
  doc.rect(W / 2 - 50, y + 5, 100, 2.4, "F");
  doc
    .setFont("helvetica", "normal")
    .setFontSize(8)
    .setTextColor(...LINE);
  doc.text(`WO No. ${inv.workOrderNo || "-"}`, W - M, y - 4, { align: "right" });
  y += 24;

  // ---- Info box (bold labels, normal values) -----------------------------
  const boxTop = y;
  const leftPairs: [string, string][] = [
    ["To,", "Executive Engineer"],
    ["Section :-", inv.section],
    ["Sub Division :-", inv.subDivision],
    ["Division :-", `M.S.E.D.C.L O & M ${inv.division}`],
    ["Work Order :-", `EE/${inv.division}/LOE :- ${inv.loeNo}  Dt. ${inv.loeDate}`],
  ];
  const rightPairs: [string, string][] = [
    ["RE Bill No :", inv.reBillNo],
    ["RA Bill Date :", inv.raBillDate],
    ["Work Order No :", inv.workOrderNo],
    ["W.O. Date :", inv.workOrderDate],
    ["GSTIN :", FIRM.gstin],
  ];
  const rowH = 14;
  const drawPair = (x: number, align: "left" | "right", pairs: [string, string][]) => {
    pairs.forEach(([label, value], i) => {
      const ry = boxTop + 12 + i * rowH;
      doc
        .setFont("helvetica", "bold")
        .setFontSize(9)
        .setTextColor(...INK);
      if (align === "left") {
        doc.text(label, x, ry);
        doc.setFont("helvetica", "normal").setTextColor(30, 30, 30);
        doc.text(value, x + doc.getTextWidth(label) + 4, ry, { maxWidth: W / 2 - x - 8 });
      } else {
        const full = `${label} ${value}`;
        doc.setFont("helvetica", "normal").setTextColor(30, 30, 30);
        doc.text(full, x, ry, { align: "right" });
      }
    });
  };
  drawPair(M + 8, "left", leftPairs);
  drawPair(W - M - 8, "right", rightPairs);
  const boxH = rowH * leftPairs.length + 16;
  doc.setFillColor(...ACCENT);
  doc.rect(M, boxTop, W - 2 * M, 2.2, "F");
  doc.setDrawColor(...LINE).setLineWidth(0.6);
  doc.rect(M, boxTop, W - 2 * M, boxH);
  doc.line(W / 2, boxTop, W / 2, boxTop + boxH);
  y = boxTop + boxH + 16;

  // ---- Line items table -----------------------------------------------
  const rows = bodyRows(inv).map((rr) =>
    isSectionRow(rr)
      ? [String(rr.sr), "", "", "", "", rr.amount === "" ? "" : inr(Number(rr.amount))]
      : [
          String(rr.sr),
          rr.desc,
          rr.unit,
          String(rr.qty),
          inr(Number(rr.rate)),
          inr(Number(rr.qty) * Number(rr.rate)),
        ],
  );

  autoTable(doc, {
    startY: y,
    head: [["Sr.No", "Description", "Unit", "Qty", "Rate", "Amount (Rs.)"]],
    body: rows,
    styles: {
      fontSize: 8,
      cellPadding: 3,
      lineColor: LINE,
      lineWidth: 0.5,
      textColor: [0, 0, 0],
    },
    headStyles: { fillColor: ACCENT, textColor: 255, halign: "center", fontStyle: "bold" },
    columnStyles: {
      0: { cellWidth: 42, halign: "center" },
      1: { cellWidth: "auto" },
      2: { cellWidth: 36, halign: "center" },
      3: { cellWidth: 36, halign: "right" },
      4: { cellWidth: 58, halign: "right" },
      5: { cellWidth: 72, halign: "right" },
    },
    margin: { left: M, right: M, top: TOP, bottom: BOTTOM },
    didParseCell: (data) => {
      const raw = data.row.raw as string[];
      const isBanner = data.section === "body" && raw[1] === "" && raw[2] === "";
      if (isBanner) {
        if (data.column.index === 0) {
          data.cell.colSpan = raw[5] ? 5 : 6;
          data.cell.styles.fontStyle = "bold";
        }
        data.cell.styles.fillColor = BAND;
        data.cell.styles.textColor = ACCENT;
      } else if (data.section === "body" && data.row.index % 2 === 1) {
        data.cell.styles.fillColor = ZEBRA;
      }
    },
  });

  y = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY;

  // ---- Totals box (right aligned, boxed) --------------------------------
  const totalsW = 220;
  const totalsX = W - M - totalsW;
  const totalLines: [string, string][] = [
    ["TOTAL (Excl. Taxes)", inr(t.totalExclusive)],
    ["G.S.T. 18%", inr(t.gst)],
    ["GRAND TOTAL (Incl. Taxes)", inr(t.totalInclusive)],
  ];
  const totalsRowH = 16;
  doc.setDrawColor(...LINE).setLineWidth(0.6);
  doc.rect(totalsX, y, totalsW, totalsRowH * totalLines.length);
  totalLines.forEach(([label, val], i) => {
    const ry = y + i * totalsRowH;
    if (i === totalLines.length - 1) {
      doc.setFillColor(...ACCENT);
      doc.rect(totalsX, ry, totalsW, totalsRowH, "F");
      doc.setDrawColor(...LINE).rect(totalsX, ry, totalsW, totalsRowH);
      doc.setFont("helvetica", "bold").setTextColor(255, 255, 255);
    } else {
      doc.setFont("helvetica", "normal").setTextColor(0, 0, 0);
    }
    doc.setFontSize(9);
    doc.text(label, totalsX + 8, ry + 11);
    doc.text(val, totalsX + totalsW - 8, ry + 11, { align: "right" });
    if (i > 0) doc.line(totalsX, ry, totalsX + totalsW, ry);
  });
  y += totalsRowH * totalLines.length + 18;

  // ---- Wrapping writer with page-break + reserved footer space ---------
  const ensureSpace = (need: number) => {
    if (y + need > H - BOTTOM) {
      doc.addPage();
      y = TOP;
    }
  };
  const wrap = (text: string, size = 8.5, indent = 0, bold = false) => {
    doc
      .setFontSize(size)
      .setFont("helvetica", bold ? "bold" : "normal")
      .setTextColor(0, 0, 0);
    const lines = doc.splitTextToSize(text, W - M * 2 - indent) as string[];
    for (const line of lines) {
      ensureSpace(size + 4);
      doc.text(line, M + indent, y);
      y += size + 4;
    }
  };

  wrap(`Amount in words: ${inv.amountInWords}`, 9, 0, true);
  wrap(`GST No: ${FIRM.gstin}      PAN No: ${FIRM.pan}`, 9);
  y += 10;

  // ---- Signature block comes right after totals — kept as one unit so it
  // can never be orphaned onto its own page. -----------------------------
  const SIGNATURE_BLOCK_H = 46;
  ensureSpace(SIGNATURE_BLOCK_H);
  doc
    .setFont("helvetica", "bold")
    .setFontSize(9.5)
    .setTextColor(...INK);
  doc.text(`For ${FIRM.name}`, W - M, y, { align: "right" });
  y += 26;
  doc.setFont("helvetica", "normal").setFontSize(8.5).setTextColor(0, 0, 0);
  doc.text("Authorised Signatory", W - M, y, { align: "right" });

  // ---- 5-6 line gap, then the CERTIFICATE block below the signature -----
  y += 8.5 * 6;

  // Keep the certificate heading + all its lines together where possible;
  // if it can't fit on the current page, push the whole block to a new one
  // rather than splitting it right after the heading.
  const certBlockH = 24 + CERTIFICATE_LINES.length * 12.5;
  if (y + Math.min(certBlockH, 140) > H - BOTTOM) {
    doc.addPage();
    y = TOP;
  }

  ensureSpace(16);
  doc.setFillColor(...ACCENT);
  doc.rect(M, y, W - 2 * M, 1.4, "F");
  y += 14;
  wrap("CERTIFICATE", 10, 0, true);
  CERTIFICATE_LINES.forEach((line, i) => wrap(`${i + 1})  ${line}`, 8.5, 8));
}
