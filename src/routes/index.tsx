import { createFileRoute } from "@tanstack/react-router";
import {
  AlertTriangle,
  Copy,
  Download,
  FileSpreadsheet,
  FileText,
  Loader2,
  Save,
  Search,
  Trash2,
  Upload,
} from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import { toast } from "sonner";

import { InvoiceItemsTable } from "@/components/InvoiceItemsTable";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Separator } from "@/components/ui/separator";
import { Toaster } from "@/components/ui/sonner";
import { exportExcel, exportExcelBulk, exportPdf, exportPdfBulk } from "@/lib/invoice-export";
import {
  CERTIFICATE_LINES,
  FIRM,
  chainInvoice,
  computeTotals,
  inr,
  invoiceFromWorkOrder,
  nextReBillNo,
  type Invoice,
} from "@/lib/invoice-model";
import {
  deleteInvoice,
  duplicateInvoice,
  loadHistory,
  saveInvoice,
  searchHistory,
} from "@/lib/invoice-history";
import { extractPdfChunks, parseWorkOrder } from "@/lib/wo-parser";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "MSEDCL Tax Invoice Generator | Aditya Enterprises" },
      {
        name: "description",
        content:
          "Upload an MSEDCL work order PDF and generate a ready-to-file tax invoice in Excel and PDF, with editable line items and saved invoice history.",
      },
      { property: "og:title", content: "MSEDCL Tax Invoice Generator | Aditya Enterprises" },
      {
        property: "og:description",
        content:
          "Parse MSEDCL work orders and produce Excel and PDF tax invoices entirely in your browser.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

function Field({
  label,
  value,
  onChange,
  highlight,
  placeholder,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  highlight?: boolean;
  placeholder?: string;
}) {
  return (
    <div className="space-y-1.5">
      <Label className="text-xs text-muted-foreground">{label}</Label>
      <Input
        value={value}
        placeholder={placeholder}
        onChange={(e) => onChange(e.target.value)}
        className={highlight ? "border-primary/60 ring-2 ring-primary/15" : undefined}
      />
    </div>
  );
}

const MAX_BULK = 10;

function Index() {
  const [invoices, setInvoices] = useState<Invoice[]>([]);
  const [activeIndex, setActiveIndex] = useState(0);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const [history, setHistory] = useState<Invoice[]>([]);
  const [query, setQuery] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => setHistory(loadHistory()), []);

  const invoice = invoices[activeIndex] ?? null;
  const totals = useMemo(() => (invoice ? computeTotals(invoice) : null), [invoice]);
  const filtered = useMemo(() => searchHistory(history, query), [history, query]);
  const mixedBatch = useMemo(() => {
    if (invoices.length < 2) return false;
    const sections = new Set(invoices.map((i) => i.section));
    const divisions = new Set(invoices.map((i) => i.division));
    return sections.size > 1 || divisions.size > 1;
  }, [invoices]);

  /** Editing RE Bill No, RA Bill Date, Section, or Sub Division on any PO in
   * the batch re-chains every PO after it (RE Bill No auto-increments, the
   * rest copy forward) — still fully editable afterwards on any tab. */
  const CASCADE_KEYS = ["reBillNo", "raBillDate", "section", "subDivision"] as const;
  const patch = (p: Partial<Invoice>) =>
    setInvoices((cur) => {
      const next = cur.map((inv, i) => (i === activeIndex ? { ...inv, ...p } : inv));
      const touched = CASCADE_KEYS.filter((k) => k in p);
      if (touched.length) {
        for (let i = activeIndex + 1; i < next.length; i++) {
          const prev = next[i - 1]!;
          const updates: Partial<Invoice> = {};
          if (touched.includes("reBillNo")) updates.reBillNo = nextReBillNo(prev.reBillNo);
          if (touched.includes("raBillDate")) updates.raBillDate = prev.raBillDate;
          if (touched.includes("section")) updates.section = prev.section;
          if (touched.includes("subDivision")) updates.subDivision = prev.subDivision;
          next[i] = { ...next[i]!, ...updates };
        }
      }
      return next;
    });

  async function handleFiles(files: File[]) {
    if (invoices.length >= MAX_BULK) {
      toast.error(`You can batch a maximum of ${MAX_BULK} POs at a time.`);
      return;
    }
    const room = MAX_BULK - invoices.length;
    const toParse = files.slice(0, room);
    if (files.length > room) {
      toast.error(`Only added ${room} of ${files.length} files — ${MAX_BULK} PO limit per batch.`);
    }
    setBusy(true);
    try {
      const parsedInvoices: Invoice[] = [];
      const allWarnings: string[] = [];
      for (const file of toParse) {
        try {
          const chunks = await extractPdfChunks(await file.arrayBuffer());
          const parsed = parseWorkOrder(chunks);
          if (parsed.warnings.length)
            allWarnings.push(...parsed.warnings.map((w) => `${file.name}: ${w}`));
          parsedInvoices.push(invoiceFromWorkOrder(parsed));
        } catch (err) {
          console.error(err);
          toast.error(
            `Could not read "${file.name}" — make sure it's a text-based MSEDCL work order.`,
          );
        }
      }
      if (parsedInvoices.length) {
        setInvoices((cur) => {
          const merged = [...cur];
          for (const inv of parsedInvoices) {
            merged.push(chainInvoice(merged[merged.length - 1], inv));
          }
          return merged;
        });
        setActiveIndex(invoices.length); // jump to first newly added
        setWarnings(allWarnings);
        toast.success(
          parsedInvoices.length === 1
            ? `Parsed ${parsedInvoices[0]!.materials.length} material and ${parsedInvoices[0]!.services.length} service line items.`
            : `Parsed ${parsedInvoices.length} POs — RE Bill No, date, section & sub division carried forward from the previous PO.`,
        );
      }
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="min-h-screen bg-muted/30">
      <Toaster position="top-right" />
      <header className="border-b border-border bg-card">
        <div className="mx-auto flex max-w-6xl flex-wrap items-center justify-between gap-3 px-6 py-5">
          <div>
            <h1 className="text-xl font-semibold tracking-tight">MSEDCL Tax Invoice Generator</h1>
            <p className="text-sm text-muted-foreground">
              {FIRM.name} · GSTIN {FIRM.gstin} · PAN {FIRM.pan}
            </p>
          </div>
          <div className="flex gap-2">
            <input
              ref={fileRef}
              type="file"
              accept="application/pdf"
              multiple
              className="hidden"
              onChange={(e) => {
                const files = Array.from(e.target.files ?? []);
                if (files.length) void handleFiles(files);
                e.target.value = "";
              }}
            />
            <Button onClick={() => fileRef.current?.click()} disabled={busy}>
              {busy ? <Loader2 className="size-4 animate-spin" /> : <Upload className="size-4" />}
              Upload PO PDF{invoices.length > 0 ? "s" : ""} ({invoices.length}/{MAX_BULK})
            </Button>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-6xl space-y-6 px-6 py-6">
        {!invoice && (
          <section className="rounded-xl border border-dashed border-border bg-card p-12 text-center">
            <FileText className="mx-auto size-10 text-muted-foreground" />
            <h2 className="mt-4 text-lg font-medium">Start with a work order</h2>
            <p className="mx-auto mt-1 max-w-md text-sm text-muted-foreground">
              Upload up to {MAX_BULK} text-based MSEDCL PO PDFs from the same section &amp; division
              at once. Material and service line items, work order metadata, transportation/
              insurance charges and ERP totals are parsed automatically — everything stays in your
              browser.
            </p>
            <Button className="mt-5" onClick={() => fileRef.current?.click()} disabled={busy}>
              <Upload className="size-4" /> Choose PDF(s)
            </Button>
          </section>
        )}

        {invoices.length > 1 && (
          <section className="rounded-xl border border-border bg-card p-4">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap gap-1.5">
                {invoices.map((inv, i) => (
                  <button
                    key={inv.id}
                    onClick={() => setActiveIndex(i)}
                    className={`rounded-md border px-3 py-1.5 text-xs font-medium transition-colors ${
                      i === activeIndex
                        ? "border-primary bg-primary/10 text-primary"
                        : "border-border text-muted-foreground hover:bg-muted"
                    }`}
                  >
                    PO {i + 1} · {inv.reBillNo || inv.workOrderNo || "untitled"}
                    <span
                      role="button"
                      className="ml-2 text-muted-foreground/70 hover:text-destructive"
                      onClick={(e) => {
                        e.stopPropagation();
                        setInvoices((cur) => cur.filter((_, idx) => idx !== i));
                        setActiveIndex((cur) => Math.max(0, cur >= i ? cur - 1 : cur));
                      }}
                    >
                      ✕
                    </span>
                  </button>
                ))}
              </div>
              <div className="flex flex-wrap gap-2">
                <Button size="sm" onClick={() => void exportExcelBulk(invoices)}>
                  <FileSpreadsheet className="size-4" /> Export all → Excel ({invoices.length}{" "}
                  sheets)
                </Button>
                <Button size="sm" variant="secondary" onClick={() => exportPdfBulk(invoices)}>
                  <Download className="size-4" /> Export all → PDF (merged)
                </Button>
              </div>
            </div>
            {mixedBatch && (
              <p className="mt-3 flex items-center gap-2 text-xs text-destructive">
                <AlertTriangle className="size-3.5" /> These POs are not all from the same section
                &amp; division — double check before batch-exporting.
              </p>
            )}
          </section>
        )}

        {invoice && totals && (
          <>
            {warnings.length > 0 && (
              <div className="rounded-lg border border-destructive/40 bg-destructive/5 p-4 text-sm">
                <p className="flex items-center gap-2 font-medium text-destructive">
                  <AlertTriangle className="size-4" /> Review these parsed values
                </p>
                <ul className="mt-2 list-disc space-y-1 pl-6 text-muted-foreground">
                  {warnings.map((w) => (
                    <li key={w}>{w}</li>
                  ))}
                </ul>
              </div>
            )}

            <section className="rounded-xl border border-border bg-card p-5">
              <h2 className="text-sm font-semibold">Invoice header</h2>
              <p className="mt-1 text-xs text-muted-foreground">
                RE Bill No and RA Bill Date are the manual inputs; everything else is parsed and
                editable.
              </p>
              <div className="mt-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
                <Field
                  label="RE Bill No (manual)"
                  highlight
                  value={invoice.reBillNo}
                  placeholder="27"
                  onChange={(v) => patch({ reBillNo: v })}
                />
                <Field
                  label="RA Bill Date (manual)"
                  highlight
                  value={invoice.raBillDate}
                  placeholder="06/07/2026"
                  onChange={(v) => patch({ raBillDate: v })}
                />
                <Field
                  label="Work Order No"
                  value={invoice.workOrderNo}
                  onChange={(v) => patch({ workOrderNo: v })}
                />
                <Field
                  label="W.O. Date"
                  value={invoice.workOrderDate}
                  onChange={(v) => patch({ workOrderDate: v })}
                />
                <Field
                  label="LOE / Award No"
                  value={invoice.loeNo}
                  onChange={(v) => patch({ loeNo: v })}
                />
                <Field
                  label="LOE Date"
                  value={invoice.loeDate}
                  onChange={(v) => patch({ loeDate: v })}
                />
                <Field
                  label="Section"
                  value={invoice.section}
                  onChange={(v) => patch({ section: v })}
                />
                <Field
                  label="Sub division"
                  value={invoice.subDivision}
                  onChange={(v) => patch({ subDivision: v })}
                />
                <Field
                  label="O & M Division"
                  value={invoice.division}
                  onChange={(v) => patch({ division: v })}
                />
                <div className="sm:col-span-2 lg:col-span-3">
                  <Field
                    label={`MO numbers (${invoice.moNos.length})`}
                    value={invoice.moNos.join(", ")}
                    onChange={(v) =>
                      patch({
                        moNos: v
                          .split(",")
                          .map((s) => s.trim())
                          .filter(Boolean),
                      })
                    }
                  />
                </div>
              </div>
            </section>

            <InvoiceItemsTable
              title="Material Supplied by Agency"
              items={invoice.materials}
              onChange={(materials) => patch({ materials })}
            />
            <InvoiceItemsTable
              title="Services Supplied by Agency"
              items={invoice.services}
              onChange={(services) => patch({ services })}
            />

            <section className="grid gap-6 rounded-xl border border-border bg-card p-5 lg:grid-cols-2">
              <div className="space-y-4">
                <h2 className="text-sm font-semibold">Charges &amp; rate quoted</h2>
                <div className="grid gap-4 sm:grid-cols-2">
                  <Field
                    label={`Transportation Charges @ ${invoice.transportPct}%`}
                    value={String(invoice.transportAmount)}
                    onChange={(v) => patch({ transportAmount: Number(v) || 0 })}
                  />
                  <Field
                    label={`Insurance Charges @ ${invoice.insurancePct}%`}
                    value={String(invoice.insuranceAmount)}
                    onChange={(v) => patch({ insuranceAmount: Number(v) || 0 })}
                  />
                  <Field
                    label="Rate quoted by Agency (%)"
                    value={String(invoice.ratePct)}
                    onChange={(v) => patch({ ratePct: Number(v) || 0 })}
                  />
                  <Field
                    label="Rate quoted amount"
                    value={String(invoice.rateAmount)}
                    onChange={(v) => patch({ rateAmount: Number(v) || 0 })}
                  />
                </div>
                <Field
                  label="Total ERP Amount (Incl. Taxes)"
                  value={String(invoice.totalInclusive)}
                  onChange={(v) => patch({ totalInclusive: Number(v) || 0 })}
                />
                <Field
                  label="Amount in words"
                  value={invoice.amountInWords}
                  onChange={(v) => patch({ amountInWords: v })}
                />
              </div>

              <div className="space-y-2 rounded-lg bg-muted/60 p-4 text-sm">
                <Row label="Material total" value={totals.materialTotal} />
                <Row
                  label={`Transportation @ ${invoice.transportPct}%`}
                  value={invoice.transportAmount}
                />
                <Row
                  label={`Insurance @ ${invoice.insurancePct}%`}
                  value={invoice.insuranceAmount}
                />
                <Row label="Service total" value={totals.serviceTotal} />
                <Row label={`Rate quoted ${invoice.ratePct}% Above`} value={invoice.rateAmount} />
                <Separator />
                <Row label="TOTAL (Excl. Taxes)" value={totals.totalExclusive} bold />
                <Row label="G.S.T 18% (inclusive − exclusive)" value={totals.gst} bold />
                <Row label="G. TOTAL (Incl. Taxes)" value={totals.totalInclusive} bold />
                {Math.abs(totals.totalExclusive - invoice.totalExclusive) > 0.5 &&
                  invoice.totalExclusive > 0 && (
                    <p className="pt-1 text-xs text-destructive">
                      Computed exclusive total differs from the work order value (
                      {inr(invoice.totalExclusive)}).
                    </p>
                  )}
                <div className="flex flex-wrap gap-2 pt-3">
                  <Button onClick={() => exportExcel(invoice)}>
                    <FileSpreadsheet className="size-4" /> Excel
                  </Button>
                  <Button variant="secondary" onClick={() => exportPdf(invoice)}>
                    <Download className="size-4" /> PDF
                  </Button>
                  <Button
                    variant="outline"
                    onClick={() => {
                      setHistory(saveInvoice(invoice));
                      toast.success("Invoice saved to history.");
                    }}
                  >
                    <Save className="size-4" /> Save
                  </Button>
                </div>
              </div>
            </section>

            <section className="rounded-xl border border-border bg-card p-5">
              <h2 className="text-sm font-semibold">Certificate (printed on every invoice)</h2>
              <ol className="mt-2 list-decimal space-y-1 pl-6 text-sm text-muted-foreground">
                {CERTIFICATE_LINES.map((c) => (
                  <li key={c}>{c}</li>
                ))}
              </ol>
            </section>
          </>
        )}

        <section className="rounded-xl border border-border bg-card p-5">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <h2 className="text-sm font-semibold">Invoice history ({history.length})</h2>
            <div className="relative">
              <Search className="absolute left-2.5 top-2.5 size-4 text-muted-foreground" />
              <Input
                className="w-64 pl-8"
                placeholder="Search by invoice or WO number"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
              />
            </div>
          </div>
          <div className="mt-4 divide-y divide-border">
            {filtered.map((h) => (
              <div key={h.id} className="flex flex-wrap items-center justify-between gap-3 py-2.5">
                <div className="text-sm">
                  <p className="font-medium">
                    RE Bill {h.reBillNo || "—"} · WO {h.workOrderNo || "—"}
                  </p>
                  <p className="text-xs text-muted-foreground">
                    {h.subDivision || "—"} · {inr(computeTotals(h).totalInclusive)} · saved{" "}
                    {new Date(h.savedAt).toLocaleString("en-IN")}
                  </p>
                </div>
                <div className="flex gap-1.5">
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={() => {
                      if (invoices.length >= MAX_BULK) {
                        toast.error(`You can batch a maximum of ${MAX_BULK} POs at a time.`);
                        return;
                      }
                      setInvoices((cur) => [...cur, structuredClone(h)]);
                      setActiveIndex(invoices.length);
                      setWarnings([]);
                      window.scrollTo({ top: 0, behavior: "smooth" });
                    }}
                  >
                    Load
                  </Button>
                  <Button
                    size="sm"
                    variant="ghost"
                    onClick={() => {
                      setHistory(duplicateInvoice(h.id));
                      toast.success("Invoice duplicated.");
                    }}
                  >
                    <Copy className="size-3.5" /> Duplicate
                  </Button>
                  <Button
                    size="sm"
                    variant="ghost"
                    className="text-destructive"
                    onClick={() => setHistory(deleteInvoice(h.id))}
                  >
                    <Trash2 className="size-3.5" />
                  </Button>
                </div>
              </div>
            ))}
            {filtered.length === 0 && (
              <p className="py-6 text-center text-sm text-muted-foreground">
                {history.length ? "No invoices match that search." : "No saved invoices yet."}
              </p>
            )}
          </div>
        </section>
      </div>
    </main>
  );
}

function Row({ label, value, bold }: { label: string; value: number; bold?: boolean }) {
  return (
    <div className={`flex justify-between ${bold ? "font-semibold" : "text-muted-foreground"}`}>
      <span>{label}</span>
      <span className="tabular-nums">{inr(value)}</span>
    </div>
  );
}
