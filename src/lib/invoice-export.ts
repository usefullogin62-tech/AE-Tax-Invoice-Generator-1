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

  // Compress the item-row font/height as the item count grows, so a long
  // PO's table doesn't push the CERTIFICATE block onto a 3rd printed page
  // (each new page would need fresh letterhead stationery, which is not
  // allowed just for the certificate).
  const totalItemRows = inv.materials.length + inv.services.length;
  const rowScale = totalItemRows <= 25 ? 1 : totalItemRows <= 40 ? 0.9 : totalItemRows <= 55 ? 0.82 : 0.75;
  const itemFontSize = Math.max(9 * rowScale, 7);
  const itemRowH = Math.max(15 * rowScale, 11);

  let zebraIdx = 0;
  for (const row of bodyRows(inv)) {
    if (isSectionRow(row)) {
      const hasAmount = row.amount !== "";
      ws.mergeCells(r, 1, r, hasAmount ? 5 : 6);
      const c = ws.getCell(r, 1);
      c.value = String(row.sr);
      c.font = { bold: true, size: itemFontSize, color: { argb: `FF${ACCENT_HEX}` } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${BAND_HEX}` } };
      c.alignment = { vertical: "middle" };
      if (hasAmount) {
        const amt = ws.getCell(r, 6);
        amt.value = Number(row.amount);
        amt.numFmt = "#,##0.00";
        amt.font = { bold: true, size: itemFontSize };
        amt.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${BAND_HEX}` } };
        amt.alignment = { horizontal: "right", vertical: "middle" };
      }
      styleRange(ws, r, 1, r, 6, (cell) => (cell.border = { ...boxBorder }));
      ws.getRow(r).height = itemRowH;
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
        c.font = { size: itemFontSize };
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
      ws.getRow(r).height = itemRowH;
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

  // ---- Signature block (comes right after totals, as per format). Font
  // fixed at 12, centered within its own merged cell, anchored bottom-right
  // of the block. ----------------------------------------------------------
  r += 3;
  ws.mergeCells(r, 5, r, 6);
  const forFirm = ws.getCell(r, 5);
  forFirm.value = `For ${FIRM.name}`;
  forFirm.font = { bold: true, size: 12 };
  forFirm.alignment = { horizontal: "center", vertical: "middle" };
  r += 3;
  ws.mergeCells(r, 5, r, 6);
  const sig = ws.getCell(r, 5);
  sig.value = "Authorised Signatory";
  sig.font = { size: 12 };
  sig.alignment = { horizontal: "center", vertical: "middle" };
  r++;

  // ---- Gap, then CERTIFICATE below the signature. Larger, centered
  // heading; larger body text — this block's size is fixed, never shrunk,
  // so it always reads clearly regardless of how the table above was
  // compressed to make room for it. ----------------------------------------
  r += 6;

  ws.mergeCells(r, 1, r, 6);
  const certHead = ws.getCell(r, 1);
  certHead.value = "CERTIFICATE";
  certHead.font = { bold: true, size: 12, color: { argb: `FF${ACCENT_HEX}` } };
  certHead.alignment = { horizontal: "center", vertical: "middle" };
  certHead.border = { bottom: { style: "medium", color: { argb: `FF${ACCENT_HEX}` } } };
  ws.getRow(r).height = 20;
  r++;

  CERTIFICATE_LINES.forEach((line, i) => {
    ws.getCell(r, 1).value = `${i + 1})`;
    ws.getCell(r, 1).font = { bold: true, size: 10 };
    ws.getCell(r, 1).alignment = { horizontal: "center", vertical: "top" };
    ws.mergeCells(r, 2, r, 6);
    const c = ws.getCell(r, 2);
    c.value = line;
    c.font = { size: 10 };
    c.alignment = { wrapText: true, vertical: "top" };
    ws.getRow(r).height = 16;
    r++;
  });

  ws.pageSetup.printArea = `A1:F${r}`;
}

/* -------------------------------- PDF --------------------------------- */
/*
 * Print-back-to-back strategy: pages are printed double-sided onto
 * pre-printed letterhead stationery. Physical sheet 1's FRONT is PDF
 * page 1 (needs the reserved letterhead top/bottom band); sheet 1's
 * BACK is PDF page 2 (plain paper visually, so it can use the full
 * page — no reserved band, saves space); sheet 2's FRONT is PDF page
 * 3 (letterhead band again); and so on. So ODD pages get the wide
 * letterhead margins, EVEN pages get slim margins.
 *
 * Certificate placement rule: the certificate block must never be the
 * reason a 3rd physical sheet (page 3) is needed. If normal-size
 * content would spill past 2 pages, the table/spacing is compressed
 * (smaller table font/padding/gaps) — tried at decreasing scales —
 * until everything (including the certificate) fits within 2 pages.
 * Signature and certificate text sizes are fixed, not shrunk.
 */

const LETTERHEAD_MARGIN = 154.8; // 2.15in
const PLAIN_MARGIN = 30; // slim margin for the back-of-sheet page

function marginFor(pageInInvoice: number) {
  const isOdd = pageInInvoice % 2 === 1;
  return isOdd
    ? { top: LETTERHEAD_MARGIN, bottom: LETTERHEAD_MARGIN }
    : { top: PLAIN_MARGIN, bottom: PLAIN_MARGIN };
}

/** Renders one invoice at a given compression scale into a fresh doc, just
 * to measure how many pages it takes — used to pick the smallest scale
 * (closest to 1 = least compressed) that still fits in 2 pages. */
function pagesNeededAtScale(inv: Invoice, scale: number): number {
  const probe = new jsPDF({ unit: "pt", format: "a4" });
  drawInvoicePage(probe, inv, scale);
  return probe.getNumberOfPages();
}

function bestScaleFor(inv: Invoice): number {
  const candidates = [1, 0.92, 0.85, 0.78, 0.72, 0.66];
  for (const s of candidates) {
    if (pagesNeededAtScale(inv, s) <= 2) return s;
  }
  return candidates.at(-1) ?? 0.66;
}

export function exportPdf(inv: Invoice) {
  const scale = bestScaleFor(inv);
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  drawInvoicePage(doc, inv, scale);
  doc.save(`${fileBase(inv)}.pdf`);
}

/** Moves the cursor so the next invoice starts on an ODD (letterhead) page —
 * inserting one blank back-of-sheet page first if needed — then returns the
 * new total page count. */
function advanceToOddStart(doc: jsPDF, currentPages: number): number {
  if (currentPages % 2 === 0) {
    doc.addPage();
    return currentPages + 1;
  }
  doc.addPage(); // blank filler: unused back of the previous letterhead sheet
  doc.addPage(); // fresh odd page to draw the next invoice on
  return currentPages + 2;
}

/** Bulk export: one merged PDF; every invoice starts fresh on an odd
 * (letterhead) page so its own front page always lands on new stationery
 * (max 10 POs at a time). */
export function exportPdfBulk(invoices: Invoice[]) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  invoices.forEach((inv, i) => {
    if (i > 0) advanceToOddStart(doc, doc.getNumberOfPages());
    const scale = bestScaleFor(inv);
    drawInvoicePage(doc, inv, scale);
  });
  const first = invoices[0];
  const label = first ? `${first.section || "Bulk"}_${invoices.length}_Invoices` : "Bulk_Invoices";
  doc.save(`Tax_Invoices_${label.replace(/[^a-zA-Z0-9_-]/g, "_")}.pdf`);
}

/** Draws one invoice starting at the current page's top; adds internal pages
 * of its own (via ensureSpace) if a single invoice overflows one page.
 * `scale` compresses only the line-item table + inter-block spacing —
 * signature and certificate text stay fixed size. */
function drawInvoicePage(doc: jsPDF, inv: Invoice, scale = 1) {
  const t = computeTotals(inv);
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const M = 34;
  let pageInInvoice = 1;
  let { top: TOP, bottom: BOTTOM } = marginFor(pageInInvoice);
  const sc = (v: number) => v * scale;

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

  const pagesBeforeTable = doc.getNumberOfPages();
  const pageInInvoiceAtTableStart = pageInInvoice;
  autoTable(doc, {
    startY: y,
    head: [["Sr.No", "Description", "Unit", "Qty", "Rate", "Amount (Rs.)"]],
    body: rows,
    styles: {
      fontSize: Math.max(sc(9), 6.5),
      cellPadding: Math.max(sc(3), 1.4),
      lineColor: LINE,
      lineWidth: 0.5,
      textColor: [0, 0, 0],
    },
    headStyles: {
      fillColor: ACCENT,
      textColor: 255,
      halign: "center",
      fontStyle: "bold",
      fontSize: Math.max(sc(9.5), 7),
    },
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
    didDrawPage: (data) => {
      // autoTable's own internal page breaks also flip between the wide
      // letterhead margin and the slim back-of-sheet margin. Derive the
      // relative page number from the doc's absolute page count so this
      // stays correct even though didDrawPage also fires once for the
      // very first (non-break) page. Updating data.settings.margin here
      // is what makes autoTable actually use the new margin on the page
      // it's about to continue rendering onto.
      const extraPages = doc.getNumberOfPages() - pagesBeforeTable;
      pageInInvoice = pageInInvoiceAtTableStart + extraPages;
      const m = marginFor(pageInInvoice);
      TOP = m.top;
      BOTTOM = m.bottom;
      data.settings.margin.top = TOP;
      data.settings.margin.bottom = BOTTOM;
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
    doc.setFontSize(9.5);
    doc.text(label, totalsX + 8, ry + 11);
    doc.text(val, totalsX + totalsW - 8, ry + 11, { align: "right" });
    if (i > 0) doc.line(totalsX, ry, totalsX + totalsW, ry);
  });
  y += totalsRowH * totalLines.length + sc(18);

  // ---- Wrapping writer with page-break + reserved footer space; each
  // manual page break here also flips the margin (odd letterhead / even
  // plain), same as the table above. -------------------------------------
  const ensureSpace = (need: number) => {
    if (y + need > H - BOTTOM) {
      doc.addPage();
      pageInInvoice++;
      const m = marginFor(pageInInvoice);
      TOP = m.top;
      BOTTOM = m.bottom;
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

  wrap(`GST No: ${FIRM.gstin}      PAN No: ${FIRM.pan}`, Math.max(sc(9.5), 7.5));
  y += sc(10);

  // ---- Signature block comes right after totals — kept as one unit so it
  // can never be orphaned onto its own page. Font size and spacing here
  // are FIXED (not compressed) — this block should always read clearly. -
  const SIG_SIZE = 12;
  const SIGNATURE_BLOCK_H = 52;
  ensureSpace(SIGNATURE_BLOCK_H);
  doc
    .setFont("helvetica", "bold")
    .setFontSize(SIG_SIZE)
    .setTextColor(...INK);
  doc.text(`For ${FIRM.name}`, W - M, y, { align: "right" });
  y += 30;
  doc.setFont("helvetica", "normal").setFontSize(SIG_SIZE).setTextColor(0, 0, 0);
  doc.text("Authorised Signatory", W - M, y, { align: "right" });

  // ---- Gap, then the CERTIFICATE block below the signature. Its size is
  // also FIXED — only the space above it (table) was compressed to make
  // room, so the certificate never has to be shrunk or pushed to a 3rd
  // (fresh-letterhead) page. -----------------------------------------------
  y += sc(8.5 * 6);

  const CERT_HEAD_SIZE = 12;
  const CERT_LINE_SIZE = 10;
  const certBlockH = 26 + CERTIFICATE_LINES.length * 14;
  if (y + Math.min(certBlockH, 160) > H - BOTTOM) {
    doc.addPage();
    pageInInvoice++;
    const m = marginFor(pageInInvoice);
    TOP = m.top;
    BOTTOM = m.bottom;
    y = TOP;
  }

  ensureSpace(18);
  doc.setFillColor(...ACCENT);
  doc.rect(M, y, W - 2 * M, 1.4, "F");
  y += 16;
  doc.setFontSize(CERT_HEAD_SIZE).setFont("helvetica", "bold").setTextColor(0, 0, 0);
  doc.text("CERTIFICATE", W / 2, y, { align: "center" });
  y += CERT_HEAD_SIZE + 6;
  const certLine = (num: number, text: string) => {
    const indent = 16;
    doc.setFontSize(CERT_LINE_SIZE).setFont("helvetica", "normal");
    const lines = doc.splitTextToSize(text, W - M * 2 - indent) as string[];
    lines.forEach((line, li) => {
      ensureSpace(CERT_LINE_SIZE + 4);
      if (li === 0) {
        doc.setFont("helvetica", "bold").setTextColor(0, 0, 0);
        doc.text(`${num})`, M + indent / 2, y, { align: "center" });
      }
      doc.setFont("helvetica", "normal").setTextColor(0, 0, 0);
      doc.text(line, M + indent, y);
      y += CERT_LINE_SIZE + 4;
    });
  };
  CERTIFICATE_LINES.forEach((line, i) => certLine(i + 1, line));
}
