/**
 * Coordinate-aware parser for MSEDCL Work Order PDFs (text-based, no OCR).
 *
 * Strategy: pdf.js gives every text chunk with an (x, y) position. Chunks are
 * grouped into visual lines; a line that matches the "row" signature starts a
 * new line item and also fixes the x boundary between the Description column
 * and the Location column. Wrapped continuation lines are then attributed to
 * the description only when they sit to the left of that boundary, which makes
 * the parser resilient to page breaks and multi-line cells.
 */

export interface PdfChunk {
  str: string;
  x: number;
  y: number;
  w: number;
}

export interface WoLineItem {
  id: string;
  srNo: number;
  description: string;
  location: string;
  unit: string;
  qty: number;
  rate: number;
  amount: number;
}

export interface ParsedWorkOrder {
  workOrderNo: string;
  workOrderDate: string;
  loeNo: string;
  loeDate: string;
  moNos: string[];
  section: string;
  subDivision: string;
  division: string;
  vendorName: string;
  materials: WoLineItem[];
  services: WoLineItem[];
  transportPct: number;
  transportAmount: number;
  insurancePct: number;
  insuranceAmount: number;
  ratePct: number;
  rateAmount: number;
  totalExclusive: number;
  totalInclusive: number;
  amountInWords: string;
  warnings: string[];
}

interface Line {
  text: string;
  chunks: PdfChunk[];
  y: number;
  page: number;
}

const num = (s: string) => Number(String(s).replace(/,/g, "")) + 0 || 0;

/** Serial | description | location | unit | qty | rate | amount */
const ROW_RE =
  /^(\d{1,3})\s+(.+?)\s+(\d{3,8}\s*\|\s*\d+\s*\|\s*[\w-]+)\s+([A-Za-z][A-Za-z0-9./]{0,5})\s+(-?[\d,]+(?:\.\d+)?)\s+(-?[\d,]+(?:\.\d+)?)\s+(-?[\d,]+(?:\.\d+)?)\s*$/;

/** Fallback for work-order layouts (e.g. substation PM work orders) whose
 * Location column has no "code | code | name" delimiter — it's free text
 * glued onto the description instead, so this only captures Sr/Unit/Qty/
 * Rate/Amount and leaves description+location combined. */
const SIMPLE_ROW_RE =
  /^(\d{1,3})\s+(.+?)\s+([A-Za-z][A-Za-z0-9./]{0,5})\s+(-?[\d,]+(?:\.\d+)?)\s+(-?[\d,]+(?:\.\d+)?)\s+(-?[\d,]+(?:\.\d+)?)\s*$/;

interface RowMatch {
  sr: string;
  desc: string;
  loc: string;
  unit: string;
  qty: string;
  rate: string;
  amount: string;
}

function matchRow(t: string): RowMatch | null {
  const full = t.match(ROW_RE);
  if (full) {
    return {
      sr: full[1] ?? "0",
      desc: full[2] ?? "",
      loc: full[3] ?? "",
      unit: full[4] ?? "",
      qty: full[5] ?? "0",
      rate: full[6] ?? "0",
      amount: full[7] ?? "0",
    };
  }
  const simple = t.match(SIMPLE_ROW_RE);
  if (simple) {
    return {
      sr: simple[1] ?? "0",
      desc: simple[2] ?? "",
      loc: "",
      unit: simple[3] ?? "",
      qty: simple[4] ?? "0",
      rate: simple[5] ?? "0",
      amount: simple[6] ?? "0",
    };
  }
  return null;
}

export function groupLines(chunks: PdfChunk[], page: number): Line[] {
  const sorted = [...chunks]
    .filter((c) => c.str.trim() !== "")
    .sort((a, b) => b.y - a.y || a.x - b.x);
  const lines: Line[] = [];
  for (const c of sorted) {
    const last = lines[lines.length - 1];
    if (last && Math.abs(last.y - c.y) <= 3) {
      last.chunks.push(c);
    } else {
      lines.push({ text: "", chunks: [c], y: c.y, page });
    }
  }
  for (const l of lines) {
    l.chunks.sort((a, b) => a.x - b.x);
    l.text = l.chunks
      .map((c) => c.str)
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
  }
  return lines;
}

function descriptionFromChunks(chunks: PdfChunk[], boundaryX: number): string {
  return chunks
    .filter((c) => c.x < boundaryX - 2)
    .map((c) => c.str)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function locationBoundaryX(line: Line, locationText: string): number {
  const first = locationText.trim().split(/\s+/)[0] ?? "";
  const hit = first ? line.chunks.find((c) => c.str.includes(first.slice(0, 8))) : undefined;
  if (hit) return hit.x;
  // Fallback: assume the location column starts after 45% of the line width.
  const xs = line.chunks.map((c) => c.x);
  return Math.min(...xs) + (Math.max(...xs) - Math.min(...xs)) * 0.45;
}

export function parseWorkOrder(pages: PdfChunk[][]): ParsedWorkOrder {
  const lines: Line[] = pages.flatMap((p, i) => groupLines(p, i + 1));
  const fullText = lines.map((l) => l.text).join("\n");
  const warnings: string[] = [];

  const pick = (re: RegExp, group = 1) => {
    const m = fullText.match(re);
    return m ? (m[group] ?? "").trim() : "";
  };

  const workOrderNo =
    pick(/Work Order No\s*:\s*(\d+)/i) || pick(/WO\s*\/\s*(\d{6,})/i) || pick(/\bWO\/(\d{6,})/i);
  const workOrderDate = pick(/\bDate\s*:\s*(\d{1,2}[-/][\w]{2,3}[-/]\d{2,4})/i);
  const loeNo = pick(/Award No\.?\s*([\w/-]+)/i);
  const loeDate = pick(/Award No\.?\s*[\w/-]+\s*Dt\.?\s*(\d{1,2}[-/][\w]{2,3}[-/]\d{2,4})/i);
  const moNos = Array.from(fullText.matchAll(/'(\d{9,})'/g)).map((m) => m[1] ?? "");
  const subDivisionRaw = pick(/in\s+([A-Za-z0-9 .&-]+?)\s*O&M\s*S\/DN/i);
  const division = pick(/([A-Z][A-Z .&-]+)\s+DIVISION/);
  const vendorName = pick(/To,\s*\n?([^\n(]+)/) || pick(/^([A-Za-z][\w .&-]+)\s*\(\d+\)/m);

  // ---- line items -------------------------------------------------------
  const materials: WoLineItem[] = [];
  const services: WoLineItem[] = [];
  let bucket: WoLineItem[] | null = null;
  let current: WoLineItem | null = null;
  let boundaryX = Number.POSITIVE_INFINITY;

  let transportPct = 0;
  let transportAmount = 0;
  let insurancePct = 0;
  let insuranceAmount = 0;
  let ratePct = 0;
  let rateAmount = 0;
  let pendingNumber: number | null = null;
  let pendingRateAmount: number | null = null;
  let pendingRateSign = 1;

  const finish = () => {
    current = null;
    boundaryX = Number.POSITIVE_INFINITY;
  };

  for (const line of lines) {
    const t = line.text;

    if (/Material\s+Supplied\s+by\s+Agency/i.test(t)) {
      finish();
      bucket = materials;
      continue;
    }
    if (/Services?\s+Supplied\s+by\s+Agency/i.test(t)) {
      finish();
      bucket = services;
      continue;
    }
    if (/Total\s+(ERP\s+)?Amount\s*\(Rs/i.test(t) || /TERMS AND CONDITIONS/i.test(t)) {
      finish();
      bucket = null;
      continue;
    }

    const transport = t.match(/Transport\w*\s+Charges\s*@\s*(-?[\d.]+)\s*%/i);
    if (transport) {
      finish();
      transportPct = num(transport[1] ?? "0");
      const inline = t.match(/(-?[\d,]+\.\d+)\s*$/);
      transportAmount = inline ? num(inline[1] ?? "0") : (pendingNumber ?? 0);
      pendingNumber = null;
      continue;
    }
    const insurance = t.match(/Insurance\s+Charges\s*@\s*(-?[\d.]+)\s*%/i);
    if (insurance) {
      finish();
      insurancePct = num(insurance[1] ?? "0");
      const inline = t.match(/(-?[\d,]+\.\d+)\s*$/);
      insuranceAmount = inline ? num(inline[1] ?? "0") : (pendingNumber ?? 0);
      pendingNumber = null;
      continue;
    }
    // A row's numeric cells sometimes render on a slightly different visual
    // line than its (wrapped) "Rate quoted by Agency ...% Above/Below" label
    // — e.g. "Below -5.0 -1921.50" appearing just before the label line.
    // Capture that amount so the label line below can pick it up.
    const rateNumbers = t.match(/^(Above|Below)\s+(-?[\d,]+\.?\d*)\s+(-?[\d,]+\.\d+)\s*$/i);
    if (rateNumbers) {
      pendingRateAmount = num(rateNumbers[3] ?? "0");
      pendingRateSign = /below/i.test(rateNumbers[1] ?? "") ? -1 : 1;
      continue;
    }

    const quoted = t.match(/Rate\s+quoted\s+by\s+Agency\s*(-?[\d.]+)\s*%/i);
    if (quoted) {
      finish();
      ratePct = num(quoted[1] ?? "0");
      if (pendingRateAmount !== null) {
        rateAmount = pendingRateSign < 0 ? -Math.abs(pendingRateAmount) : Math.abs(pendingRateAmount);
      } else {
        const nums = Array.from(t.matchAll(/(-?[\d,]+\.\d+)/g)).map((m) => num(m[1] ?? "0"));
        rateAmount = nums.length ? (nums[nums.length - 1] ?? 0) : 0;
      }
      pendingRateAmount = null;
      bucket = null;
      continue;
    }

    if (/^-?[\d,]+\.\d+$/.test(t)) {
      pendingNumber = num(t);
      continue;
    }

    if (!bucket) continue;

    const row = matchRow(t);
    if (row) {
      const { sr, desc, loc, unit, qty, rate, amount } = row;
      const item: WoLineItem = {
        id: `${bucket === materials ? "M" : "S"}-${sr}-${Math.random().toString(36).slice(2, 7)}`,
        srNo: Number(sr),
        description: desc.replace(/\s+/g, " ").trim(),
        location: loc.replace(/\s+/g, ""),
        unit,
        qty: num(qty),
        rate: num(rate),
        amount: num(amount),
      };
      current = item;
      boundaryX = locationBoundaryX(line, loc);
      bucket.push(item);
      continue;
    }

    // Wrapped continuation of the description cell.
    if (current) {
      const extra = descriptionFromChunks(line.chunks, boundaryX);
      if (extra && !/^\d+$/.test(extra)) {
        current.description = `${current.description} ${extra}`.replace(/\s+/g, " ").trim();
      }
    }
  }

  for (const item of [...materials, ...services]) {
    if (Math.abs(item.qty * item.rate - item.amount) > 0.5) {
      warnings.push(
        `Sr. ${item.srNo} "${item.description.slice(0, 40)}": qty x rate does not match the work-order amount.`,
      );
    }
  }

  const totalExclusive = num(pick(/Total Amount \(Rs INR\)\s*([\d,]+\.?\d*)/i) || "0");
  const totalInclusive = num(pick(/Total ERP Amount \(Rs INR\)\s*([\d,]+\.?\d*)/i) || "0");
  const amountInWords = pick(/Total Amount In words:\s*([\s\S]*?)\(Excluding/i).replace(/\s+/g, " ");

  if (!materials.length && !services.length) {
    warnings.push("No line items were detected. The PDF may be scanned rather than text-based.");
  }
  if (!totalInclusive) warnings.push("Total ERP (inclusive) amount not found in the work order.");

  const subDivision = subDivisionRaw ? `${subDivisionRaw} O&M Sub Division` : "";

  return {
    workOrderNo,
    workOrderDate,
    loeNo,
    loeDate,
    moNos,
    section: subDivisionRaw,
    subDivision,
    division,
    vendorName,
    materials,
    services,
    transportPct,
    transportAmount,
    insurancePct,
    insuranceAmount,
    ratePct,
    rateAmount,
    totalExclusive,
    totalInclusive,
    amountInWords,
    warnings,
  };
}

/** Reads a text-based PDF in the browser and returns positioned text chunks. */
export async function extractPdfChunks(data: ArrayBuffer): Promise<PdfChunk[][]> {
  const pdfjs = await import("pdfjs-dist");
  const workerSrc = (await import("pdfjs-dist/build/pdf.worker.min.mjs?url")).default;
  pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;

  const doc = await pdfjs.getDocument({ data: new Uint8Array(data) }).promise;
  const pages: PdfChunk[][] = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    pages.push(
      content.items
        .filter((it): it is import("pdfjs-dist/types/src/display/api").TextItem => "str" in it)
        .map((it) => ({
          str: it.str,
          x: it.transform[4] as number,
          y: it.transform[5] as number,
          w: it.width,
        })),
    );
  }
  return pages;
}
