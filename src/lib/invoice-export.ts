import * as XLSX from "xlsx";
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
  rows.push({ sr: "Material Supplied by Agency", desc: "", unit: "", qty: "", rate: "", amount: "" });
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
  rows.push({ sr: "Services Supplied by Agency", desc: "", unit: "", qty: "", rate: "", amount: "" });
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

/* ------------------------------- Excel -------------------------------- */

export function exportExcel(inv: Invoice) {
  const t = computeTotals(inv);
  const aoa: (string | number | null)[][] = [];
  const push = (...cells: (string | number | null)[]) => aoa.push(cells);

  push("TAX INVOICE", null, null, null, null, null);
  push("To, Executive Engineer", null, null, null, `RE Bill No : ${inv.reBillNo}`, null);
  push(`Section :- ${inv.section}`, null, null, null, `RA Bill Date : ${inv.raBillDate}`, null);
  push(`Sub Division :- ${inv.subDivision}`, null, null, null, `Work Order No : ${inv.workOrderNo}`, null);
  push(
    `M.S.E.D.C.L O & M Division :- ${inv.division}`,
    null,
    null,
    null,
    `W.O. Date : ${inv.workOrderDate}`,
    null,
  );
  push(
    `Work Order: EE/${inv.division}/LOE :- ${inv.loeNo}   Dt. ${inv.loeDate}`,
    null,
    null,
    null,
    `GSTIN : ${FIRM.gstin}`,
    null,
  );
  push(null);
  push("Sr.No", "Description", "Unit", "Qty", "Rate", "Amount (Rs.)");

  const headerRows = aoa.length;
  for (const r of bodyRows(inv)) {
    if (isSectionRow(r) && r.amount === "") push(String(r.sr), null, null, null, null, null);
    else if (isSectionRow(r)) push(String(r.sr), null, null, null, null, Number(r.amount));
    else push(r.sr, r.desc, r.unit, Number(r.qty), Number(r.rate), Number(r.qty) * Number(r.rate));
  }

  push("TOTAL (Excl. Taxes)", null, null, null, null, t.totalExclusive);
  push("G.S.T. 18%", null, null, null, null, t.gst);
  push("GRAND TOTAL (Incl. Taxes)", null, null, null, null, t.totalInclusive);
  push(`Amount in words: ${inv.amountInWords}`, null, null, null, null, null);
  push(`GST No: ${FIRM.gstin}      PAN No: ${FIRM.pan}`, null, null, null, null, null);
  push(null);
  push("CERTIFICATE", null, null, null, null, null);
  CERTIFICATE_LINES.forEach((line, i) => push(`${i + 1})`, line, null, null, null, null));
  push(null);
  push(null, null, null, null, `For ${FIRM.name}`, null);
  push(null);
  push(null, null, null, null, "Authorised Signatory", null);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 8 }, { wch: 55 }, { wch: 8 }, { wch: 8 }, { wch: 18 }, { wch: 16 }];
  const merges: XLSX.Range[] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
    { s: { r: 4, c: 0 }, e: { r: 4, c: 3 } },
    { s: { r: 5, c: 0 }, e: { r: 5, c: 3 } },
  ];
  bodyRows(inv).forEach((r, i) => {
    if (isSectionRow(r)) {
      const row = headerRows + i;
      merges.push({ s: { r: row, c: 0 }, e: { r: row, c: r.amount === "" ? 5 : 4 } });
    }
  });
  const totalsStart = headerRows + bodyRows(inv).length;
  for (let i = 0; i < 3; i++)
    merges.push({ s: { r: totalsStart + i, c: 0 }, e: { r: totalsStart + i, c: 4 } });
  for (let i = 3; i < 5; i++)
    merges.push({ s: { r: totalsStart + i, c: 0 }, e: { r: totalsStart + i, c: 5 } });
  const certStart = totalsStart + 6;
  merges.push({ s: { r: certStart, c: 0 }, e: { r: certStart, c: 5 } });
  CERTIFICATE_LINES.forEach((_, i) =>
    merges.push({ s: { r: certStart + 1 + i, c: 1 }, e: { r: certStart + 1 + i, c: 5 } }),
  );
  const sigForRow = certStart + CERTIFICATE_LINES.length + 2;
  merges.push({ s: { r: sigForRow, c: 4 }, e: { r: sigForRow, c: 5 } });
  merges.push({ s: { r: sigForRow + 2, c: 4 }, e: { r: sigForRow + 2, c: 5 } });
  ws["!merges"] = merges;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, inv.workOrderNo || "Tax Invoice");
  XLSX.writeFile(wb, `${fileBase(inv)}.xlsx`);
}

/* -------------------------------- PDF --------------------------------- */

export function exportPdf(inv: Invoice) {
  const t = computeTotals(inv);
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const M = 34;
  const TOP = 120; // blank space reserved for pre-printed letterhead
  const BOTTOM = 70; // blank space reserved for footer

  const INK: [number, number, number] = [30, 41, 59];
  const LINE: [number, number, number] = [100, 116, 139];
  const BAND: [number, number, number] = [241, 245, 249];

  let y = TOP;

  // ---- Title -------------------------------------------------------
  doc.setFont("helvetica", "bold").setFontSize(15).setTextColor(...INK);
  doc.text("TAX INVOICE", W / 2, y, { align: "center" });
  doc.setDrawColor(...INK).setLineWidth(0.8);
  doc.line(W / 2 - 44, y + 4, W / 2 + 44, y + 4);
  y += 22;

  // ---- Info box ------------------------------------------------------
  const boxTop = y;
  const left = [
    "To, Executive Engineer",
    `Section :- ${inv.section}`,
    `Sub Division :- ${inv.subDivision}`,
    `M.S.E.D.C.L O & M Division :- ${inv.division}`,
    `Work Order: EE/${inv.division}/LOE :- ${inv.loeNo}   Dt. ${inv.loeDate}`,
  ];
  const right = [
    `RE Bill No : ${inv.reBillNo}`,
    `RA Bill Date : ${inv.raBillDate}`,
    `Work Order No : ${inv.workOrderNo}`,
    `W.O. Date : ${inv.workOrderDate}`,
    `GSTIN : ${FIRM.gstin}`,
  ];
  doc.setFontSize(9).setFont("helvetica", "normal").setTextColor(0, 0, 0);
  const rowH = 14;
  left.forEach((l, i) => doc.text(l, M + 8, boxTop + 12 + i * rowH, { maxWidth: W / 2 - M - 16 }));
  right.forEach((l, i) => doc.text(l, W - M - 8, boxTop + 12 + i * rowH, { align: "right" }));
  const boxH = rowH * left.length + 16;
  doc.setDrawColor(...LINE).setLineWidth(0.6);
  doc.rect(M, boxTop, W - 2 * M, boxH);
  doc.line(W / 2, boxTop, W / 2, boxTop + boxH);
  y = boxTop + boxH + 16;

  // ---- Line items table -----------------------------------------------
  const rows = bodyRows(inv).map((r) =>
    isSectionRow(r)
      ? [String(r.sr), "", "", "", "", r.amount === "" ? "" : inr(Number(r.amount))]
      : [
          String(r.sr),
          r.desc,
          r.unit,
          String(r.qty),
          inr(Number(r.rate)),
          inr(Number(r.qty) * Number(r.rate)),
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
    headStyles: { fillColor: INK, textColor: 255, halign: "center", fontStyle: "bold" },
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
      doc.setFillColor(...BAND);
      doc.rect(totalsX, ry, totalsW, totalsRowH, "F");
      doc.setDrawColor(...LINE).rect(totalsX, ry, totalsW, totalsRowH);
      doc.setFont("helvetica", "bold");
    } else {
      doc.setFont("helvetica", "normal");
    }
    doc.setFontSize(9).setTextColor(0, 0, 0);
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
    doc.setFontSize(size).setFont("helvetica", bold ? "bold" : "normal").setTextColor(0, 0, 0);
    const lines = doc.splitTextToSize(text, W - M * 2 - indent) as string[];
    for (const line of lines) {
      ensureSpace(size + 4);
      doc.text(line, M + indent, y);
      y += size + 4;
    }
  };

  wrap(`Amount in words: ${inv.amountInWords}`, 9, 0, true);
  wrap(`GST No: ${FIRM.gstin}      PAN No: ${FIRM.pan}`, 9);
  y += 6;

  ensureSpace(16);
  doc.setDrawColor(...LINE).setLineWidth(0.6);
  doc.line(M, y, W - M, y);
  y += 12;
  wrap("CERTIFICATE", 10, 0, true);
  CERTIFICATE_LINES.forEach((line, i) => wrap(`${i + 1})  ${line}`, 8.5, 8));

  // ---- Signature block (kept together, never orphaned alone) -----------
  ensureSpace(50);
  y += 14;
  doc.setFont("helvetica", "bold").setFontSize(9.5).setTextColor(0, 0, 0);
  doc.text(`For ${FIRM.name}`, W - M, y, { align: "right" });
  y += 26;
  doc.setFont("helvetica", "normal").setFontSize(8.5);
  doc.text("Authorised Signatory", W - M, y, { align: "right" });

  // ---- Footer: page numbers + reserved blank strip on every page -------
  const pageCount = doc.getNumberOfPages();
  for (let p = 1; p <= pageCount; p++) {
    doc.setPage(p);
    doc.setFont("helvetica", "normal").setFontSize(7.5).setTextColor(120, 120, 120);
    doc.text(`Page ${p} of ${pageCount}`, W / 2, H - BOTTOM + 24, { align: "center" });
  }

  doc.save(`${fileBase(inv)}.pdf`);
}
