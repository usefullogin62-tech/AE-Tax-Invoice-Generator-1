import type { Invoice } from "./invoice-model";

const KEY = "aditya-invoice-history-v1";

export function loadHistory(): Invoice[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(KEY);
    const list = raw ? (JSON.parse(raw) as Invoice[]) : [];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

function persist(list: Invoice[]) {
  try {
    window.localStorage.setItem(KEY, JSON.stringify(list));
  } catch {
    // Quota or private-mode storage — keep working in memory for this session.
  }
}

export function saveInvoice(inv: Invoice): Invoice[] {
  const list = loadHistory();
  const next = { ...inv, savedAt: new Date().toISOString() };
  const idx = list.findIndex((i) => i.id === inv.id);
  if (idx >= 0) list[idx] = next;
  else list.unshift(next);
  persist(list);
  return list;
}

export function deleteInvoice(id: string): Invoice[] {
  const list = loadHistory().filter((i) => i.id !== id);
  persist(list);
  return list;
}

export function duplicateInvoice(id: string): Invoice[] {
  const list = loadHistory();
  const src = list.find((i) => i.id === id);
  if (!src) return list;
  const copy: Invoice = {
    ...structuredClone(src),
    id: crypto.randomUUID(),
    savedAt: new Date().toISOString(),
    reBillNo: src.reBillNo ? `${src.reBillNo}-COPY` : "",
  };
  list.unshift(copy);
  persist(list);
  return list;
}

export type OfficePreset = {
  section: string;
  subDivision: string;
};

const OFFICE_KEY = "aditya-office-defaults-v1";

export function loadOfficePreset(): OfficePreset | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = window.localStorage.getItem(OFFICE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Partial<OfficePreset>;
    const section = String(parsed.section ?? "").trim();
    const subDivision = String(parsed.subDivision ?? "").trim();
    if (!section && !subDivision) return null;
    return { section, subDivision };
  } catch {
    return null;
  }
}

export function saveOfficePreset(preset: OfficePreset): OfficePreset {
  const next = {
    section: preset.section.trim(),
    subDivision: preset.subDivision.trim(),
  };
  try {
    window.localStorage.setItem(OFFICE_KEY, JSON.stringify(next));
  } catch {
    // Quota or private-mode storage — still return the in-memory preset.
  }
  return next;
}

/** Unique section + sub-division pairs from newest saved invoices first. */
export function uniqueOfficesFromHistory(list: Invoice[], limit = 6): OfficePreset[] {
  const seen = new Set<string>();
  const out: OfficePreset[] = [];
  for (const inv of list) {
    const section = (inv.section ?? "").trim();
    const subDivision = (inv.subDivision ?? "").trim();
    if (!section && !subDivision) continue;
    const key = `${section.toLowerCase()}|${subDivision.toLowerCase()}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ section, subDivision });
    if (out.length >= limit) break;
  }
  return out;
}

export function searchHistory(list: Invoice[], q: string): Invoice[] {
  const term = q.trim().toLowerCase();
  if (!term) return list;
  return list.filter((i) =>
    [i.reBillNo, i.workOrderNo, i.loeNo, i.subDivision, i.division]
      .filter(Boolean)
      .some((v) => String(v).toLowerCase().includes(term)),
  );
}
