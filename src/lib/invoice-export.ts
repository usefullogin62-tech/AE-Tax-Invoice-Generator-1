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

function colAlign(i: number): "left" | "center" | "right" {
  if (i === 1) return "left"; // Description
  if (i >= 3) return "right"; // Qty, Rate, Amount
  return "center"; // Sr.No, Unit
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
      // Measured directly off Aditya Enterprises' printed letterhead: the
      // header artwork fills the top 2.01in and the footer artwork fills
      // the bottom 2.01in of the A4 sheet — these margins keep every cell
      // clear of both so the sheet can be printed straight onto it.
      margins: { top: 2.15, bottom: 2.15, left: 0.45, right: 0.45, header: 0, footer: 0 },
      horizontalCentered: true,
    },
    // No header/footer text of our own — the letterhead already carries
    // the firm's contact details in that reserved band, and Excel's
    // header/footer prints inside the margin, which would land on top of
    // the pre-printed artwork.
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
    c.alignment = { horizontal: colAlign(i), vertical: "middle" };
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
          horizontal: colAlign(i),
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

  // ---- GST/PAN ------------------------------------------------------------
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
  certHead.alignment = { horizontal: "center", vertical: "middle" };
  certHead.border = { bottom: { style: "medium", color: { argb: `FF${ACCENT_HEX}` } } };
  r++;

  CERTIFICATE_LINES.forEach((line, i) => {
    ws.getCell(r, 1).value = `${i + 1})`;
    ws.getCell(r, 1).font = { bold: true, size: 9 };
    ws.getCell(r, 1).alignment = { horizontal: "center", vertical: "top" };
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

/** Tries the invoice at full size first; only shrinks fonts/padding a notch
 * at a time, and only as far as needed, to keep everything within a single
 * front+back sheet (2 pages) instead of spilling onto a fresh letterhead
 * page. Never shrinks below 87% (stays legible). */
function pickScale(inv: Invoice): number {
  const candidates = [1, 0.93, 0.87];
  for (const s of candidates) {
    const probe = new jsPDF({ unit: "pt", format: "a4" });
    drawInvoicePage(probe, inv, s);
    if (probe.getNumberOfPages() <= 2) return s;
  }
  return candidates[candidates.length - 1];
}

export function exportPdf(inv: Invoice) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  drawInvoicePage(doc, inv, pickScale(inv));
  doc.save(`${fileBase(inv)}.pdf`);
}

/** Bulk export: one merged PDF, each invoice starting on its own page (max 10 POs at a time). */
export function exportPdfBulk(invoices: Invoice[]) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  invoices.forEach((inv, i) => {
    if (i > 0) doc.addPage();
    drawInvoicePage(doc, inv, pickScale(inv));
  });
  const first = invoices[0];
  const label = first ? `${first.section || "Bulk"}_${invoices.length}_Invoices` : "Bulk_Invoices";
  doc.save(`Tax_Invoices_${label.replace(/[^a-zA-Z0-9_-]/g, "_")}.pdf`);
}

/** Front pages (1st, 3rd, 5th... of THIS invoice) reserve the letterhead's
 * top/bottom band. Back pages (2nd, 4th...) are printed on the blank
 * reverse of the previous sheet, so they get the full page — this is what
 * lets a long item list spill onto page 2 without wasting a fresh
 * pre-printed sheet, and is why the certificate/signature block fights so
 * hard (via pickScale above) to stay within page 1+2 instead of forcing a
 * page 3. */
function pageMarginsFor(relativePage: number) {
  return relativePage % 2 === 1 ? { top: 154.8, bottom: 154.8 } : { top: 24, bottom: 24 };
}

/** Draws one invoice starting at the current page's top; adds internal pages
 * of its own (via ensureSpace) if a single invoice overflows one page.
 * `scale` (0.87–1) shrinks fonts/padding uniformly so more content fits
 * before a new page is ever added. */
function drawInvoicePage(doc: jsPDF, inv: Invoice, scale = 1) {
  const t = computeTotals(inv);
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const M = 34;

  const invoiceStartPage = doc.getCurrentPageInfo().pageNumber;
  const relPage = () => doc.getCurrentPageInfo().pageNumber - invoiceStartPage + 1;
  const TOP = pageMarginsFor(1).top; // this invoice's page 1 is always a front (letterhead) page

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
  y = boxTop + boxH + 16 * scale;

  // ---- Line items table (fontSize/padding shrink with `scale`) ----------
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
      fontSize: 9 * scale,
      cellPadding: 3 * scale,
      lineColor: LINE,
      lineWidth: 0.5,
      textColor: [0, 0, 0],
    },
    headStyles: {
      fillColor: ACCENT,
      textColor: 255,
      halign: "center",
      fontStyle: "bold",
      fontSize: 9.5 * scale,
    },
    columnStyles: {
      0: { cellWidth: 42, halign: "center" },
      1: { cellWidth: "auto" },
      2: { cellWidth: 36, halign: "center" },
      3: { cellWidth: 36, halign: "right" },
      4: { cellWidth: 58, halign: "right" },
      5: { cellWidth: 72, halign: "right" },
    },
    // Page 1 (front) reserves the letterhead band. didDrawPage flips the
    // margin for whichever page comes next: back pages get almost the full
    // sheet, front pages reserve the band again — see pageMarginsFor above.
    margin: { left: M, right: M, top: TOP, bottom: pageMarginsFor(1).bottom },
    didDrawPage: (data) => {
      const next = pageMarginsFor(relPage() + 1);
      data.settings.margin.top = next.top;
      data.settings.margin.bottom = next.bottom;
    },
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

  // ---- Page-aware helpers for everything drawn after the table ----------
  const ensureSpace = (need: number) => {
    const bm = pageMarginsFor(relPage()).bottom;
    if (y + need > H - bm) {
      doc.addPage();
      y = pageMarginsFor(relPage()).top;
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

  // ---- Totals box (right aligned, boxed) --------------------------------
  const totalsW = 220;
  const totalsX = W - M - totalsW;
  const totalLines: [string, string][] = [
    ["TOTAL (Excl. Taxes)", inr(t.totalExclusive)],
    ["G.S.T. 18%", inr(t.gst)],
    ["GRAND TOTAL (Incl. Taxes)", inr(t.totalInclusive)],
  ];
  const totalsRowH = 16 * scale;
  ensureSpace(totalsRowH * totalLines.length);
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
    doc.setFontSize(9.5 * scale);
    doc.text(label, totalsX + 8, ry + totalsRowH * 0.68);
    doc.text(val, totalsX + totalsW - 8, ry + totalsRowH * 0.68, { align: "right" });
    if (i > 0) doc.line(totalsX, ry, totalsX + totalsW, ry);
  });
  y += totalsRowH * totalLines.length + 16 * scale;

  wrap(`GST No: ${FIRM.gstin}      PAN No: ${FIRM.pan}`, 9.5 * scale);
  y += 8 * scale;

  // ---- CERTIFICATE — bigger, centered heading, kept together where the
  // available room allows (this is what pickScale is fighting to avoid
  // spilling onto a fresh letterhead page). --------------------------------
  const certBlockH = 30 * scale + CERTIFICATE_LINES.length * 13.5 * scale;
  if (y + Math.min(certBlockH, 160) > H - pageMarginsFor(relPage()).bottom) {
    doc.addPage();
    y = pageMarginsFor(relPage()).top;
  }

  ensureSpace(16);
  doc.setFillColor(...ACCENT);
  doc.rect(M, y, W - 2 * M, 1.4, "F");
  y += 16 * scale;
  doc.setFontSize(12 * scale).setFont("helvetica", "bold").setTextColor(0, 0, 0);
  doc.text("CERTIFICATE", W / 2, y, { align: "center" });
  y += 12 * scale + 6 * scale;
  const certLine = (num: number, text: string) => {
    const indent = 16;
    const fs = 9.5 * scale;
    doc.setFontSize(fs).setFont("helvetica", "normal");
    const lines = doc.splitTextToSize(text, W - M * 2 - indent) as string[];
    lines.forEach((line, li) => {
      ensureSpace(fs + 4.5);
      if (li === 0) {
        doc.setFont("helvetica", "bold").setTextColor(0, 0, 0);
        doc.text(`${num})`, M + indent / 2, y, { align: "center" });
      }
      doc.setFont("helvetica", "normal").setTextColor(0, 0, 0);
      doc.text(line, M + indent, y);
      y += fs + 4.5;
    });
  };
  CERTIFICATE_LINES.forEach((line, i) => certLine(i + 1, line));

  // ---- Signature — centered 12pt block, kept together, anchored bottom
  // right beneath the certificate. -----------------------------------------
  const sigCenterX = W - M - 90;
  const SIGNATURE_BLOCK_H = 54;
  ensureSpace(SIGNATURE_BLOCK_H);
  y += 18 * scale;
  doc.setFont("helvetica", "bold").setFontSize(12).setTextColor(...INK);
  doc.text(`For ${FIRM.name}`, sigCenterX, y, { align: "center" });
  y += 30;
  doc.setFont("helvetica", "normal").setFontSize(12).setTextColor(0, 0, 0);
  doc.text("Authorised Signatory", sigCenterX, y, { align: "center" });
}
