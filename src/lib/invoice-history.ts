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
  window.localStorage.setItem(KEY, JSON.stringify(list));
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

export function searchHistory(list: Invoice[], q: string): Invoice[] {
  const term = q.trim().toLowerCase();
  if (!term) return list;
  return list.filter((i) =>
    [i.reBillNo, i.workOrderNo, i.loeNo, i.subDivision, i.division]
      .filter(Boolean)
      .some((v) => String(v).toLowerCase().includes(term)),
  );
}
