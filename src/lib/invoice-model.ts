import type { ParsedWorkOrder, WoLineItem } from "./wo-parser";

export const FIRM = {
  name: "Aditya Enterprises",
  gstin: "27ANIPM4768J2ZD",
  pan: "ANIPM4768J",
};

export const CERTIFICATE_LINES = [
  "All Photos before work and after work uploaded by agency are verified.",
  "All the formalities (JMC, Service SES, Material (MIGO), Centages SES) regarding this invoice are completed in Maintenance portal by me.",
  "Original bill copy of contractor bills duly certified and verified by me.",
  "This Invoice is not recorded previously.",
  "Service Location and services are not repeated and recorded at the first time.",
  "Work Completion report for period of work carried out is attached.",
  "Bill is recorded for passing",
];

/** "54" -> "55", "RE-054" -> "RE-055". Falls back to the input unchanged if
 * it has no trailing digits to increment. */
export function nextReBillNo(prev: string): string {
  const m = /^(.*?)(\d+)(\D*)$/.exec(prev.trim());
  if (!m) return prev;
  const prefix = m[1] ?? "";
  const digits = m[2] ?? "";
  const suffix = m[3] ?? "";
  if (!digits) return prev;
  const next = String(parseInt(digits, 10) + 1).padStart(digits.length, "0");
  return `${prefix}${next}${suffix}`;
}

/** For bulk uploads: carries RE Bill No (auto-incremented), RA Bill Date,
 * Section & Sub Division forward from the previous PO in the same batch —
 * every field stays manually editable afterwards. Only fills gaps; never
 * overwrites values the work-order PDF itself parsed correctly. */
export function chainInvoice(prev: Invoice | undefined, inv: Invoice): Invoice {
  if (!prev) return inv;
  return {
    ...inv,
    reBillNo: inv.reBillNo || nextReBillNo(prev.reBillNo),
    raBillDate: inv.raBillDate || prev.raBillDate,
    section: inv.section || prev.section,
    subDivision: inv.subDivision || prev.subDivision,
  };
}

export interface Invoice {
  id: string;
  savedAt: string;
  /** Manual inputs — the only two fields the user is expected to type. */
  reBillNo: string;
  raBillDate: string;
  /** Parsed metadata (editable in review). */
  workOrderNo: string;
  workOrderDate: string;
  loeNo: string;
  loeDate: string;
  moNos: string[];
  section: string;
  subDivision: string;
  division: string;
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
}

export function invoiceFromWorkOrder(wo: ParsedWorkOrder): Invoice {
  return {
    id: crypto.randomUUID(),
    savedAt: new Date().toISOString(),
    reBillNo: "",
    raBillDate: "",
    workOrderNo: wo.workOrderNo,
    workOrderDate: wo.workOrderDate,
    loeNo: wo.loeNo,
    loeDate: wo.loeDate,
    moNos: wo.moNos,
    section: wo.section,
    subDivision: wo.subDivision,
    division: wo.division,
    materials: wo.materials,
    services: wo.services,
    transportPct: wo.transportPct,
    transportAmount: wo.transportAmount,
    insurancePct: wo.insurancePct,
    insuranceAmount: wo.insuranceAmount,
    ratePct: wo.ratePct,
    rateAmount: wo.rateAmount,
    totalExclusive: wo.totalExclusive,
    totalInclusive: wo.totalInclusive,
    amountInWords: wo.amountInWords,
  };
}

export interface InvoiceTotals {
  materialTotal: number;
  serviceTotal: number;
  totalExclusive: number;
  gst: number;
  totalInclusive: number;
}

const lineTotal = (items: WoLineItem[]) => items.reduce((s, i) => s + i.qty * i.rate, 0);

export function computeTotals(inv: Invoice): InvoiceTotals {
  const materialTotal = lineTotal(inv.materials);
  const serviceTotal = lineTotal(inv.services);
  const totalExclusive =
    materialTotal + inv.transportAmount + inv.insuranceAmount + serviceTotal + inv.rateAmount;
  const totalInclusive = inv.totalInclusive;
  // GST is never recalculated: it is strictly inclusive ERP total - exclusive total.
  return {
    materialTotal,
    serviceTotal,
    totalExclusive,
    gst: round2(totalInclusive - totalExclusive),
    totalInclusive,
  };
}

export const round2 = (n: number) => Math.round((n + Number.EPSILON) * 100) / 100;

export const inr = (n: number) =>
  n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
