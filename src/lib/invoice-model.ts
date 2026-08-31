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
  "Bill is recoded for passing",
];

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
