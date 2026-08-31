import { Trash2, Plus } from "lucide-react";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { inr } from "@/lib/invoice-model";
import type { WoLineItem } from "@/lib/wo-parser";

interface Props {
  title: string;
  items: WoLineItem[];
  onChange: (items: WoLineItem[]) => void;
}

export function InvoiceItemsTable({ title, items, onChange }: Props) {
  const update = (id: string, patch: Partial<WoLineItem>) =>
    onChange(items.map((i) => (i.id === id ? { ...i, ...patch } : i)));

  const total = items.reduce((s, i) => s + i.qty * i.rate, 0);

  return (
    <div className="rounded-lg border border-border">
      <div className="flex items-center justify-between border-b border-border bg-muted/60 px-3 py-2">
        <h3 className="text-sm font-semibold">
          {title} <span className="text-muted-foreground">({items.length})</span>
        </h3>
        <div className="flex items-center gap-3">
          <span className="text-sm font-medium tabular-nums">{inr(total)}</span>
          <Button
            size="sm"
            variant="outline"
            onClick={() =>
              onChange([
                ...items,
                {
                  id: crypto.randomUUID(),
                  srNo: items.length + 1,
                  description: "",
                  location: "",
                  unit: "NO",
                  qty: 0,
                  rate: 0,
                  amount: 0,
                },
              ])
            }
          >
            <Plus className="size-3.5" /> Row
          </Button>
        </div>
      </div>
      <div className="max-h-[420px] overflow-auto">
        <table className="w-full text-sm">
          <thead className="sticky top-0 bg-card text-xs text-muted-foreground">
            <tr className="border-b border-border">
              <th className="w-10 px-2 py-1.5 text-left">#</th>
              <th className="px-2 py-1.5 text-left">Description</th>
              <th className="w-20 px-2 py-1.5 text-left">Unit</th>
              <th className="w-20 px-2 py-1.5 text-right">Qty</th>
              <th className="w-24 px-2 py-1.5 text-right">Rate</th>
              <th className="w-28 px-2 py-1.5 text-right">Amount</th>
              <th className="w-10" />
            </tr>
          </thead>
          <tbody>
            {items.map((item, idx) => (
              <tr key={item.id} className="border-b border-border/60 last:border-0">
                <td className="px-2 py-1 text-muted-foreground tabular-nums">{idx + 1}</td>
                <td className="px-1 py-1">
                  <Input
                    className="h-8 border-transparent bg-transparent hover:border-input focus-visible:border-input"
                    value={item.description}
                    onChange={(e) => update(item.id, { description: e.target.value })}
                  />
                </td>
                <td className="px-1 py-1">
                  <Input
                    className="h-8 border-transparent bg-transparent hover:border-input focus-visible:border-input"
                    value={item.unit}
                    onChange={(e) => update(item.id, { unit: e.target.value })}
                  />
                </td>
                <td className="px-1 py-1">
                  <Input
                    type="number"
                    className="h-8 border-transparent bg-transparent text-right hover:border-input focus-visible:border-input"
                    value={item.qty}
                    onChange={(e) => update(item.id, { qty: Number(e.target.value) })}
                  />
                </td>
                <td className="px-1 py-1">
                  <Input
                    type="number"
                    className="h-8 border-transparent bg-transparent text-right hover:border-input focus-visible:border-input"
                    value={item.rate}
                    onChange={(e) => update(item.id, { rate: Number(e.target.value) })}
                  />
                </td>
                <td className="px-2 py-1 text-right tabular-nums">{inr(item.qty * item.rate)}</td>
                <td className="px-1 py-1">
                  <Button
                    size="icon"
                    variant="ghost"
                    className="size-7 text-muted-foreground"
                    onClick={() => onChange(items.filter((i) => i.id !== item.id))}
                    aria-label="Delete row"
                  >
                    <Trash2 className="size-3.5" />
                  </Button>
                </td>
              </tr>
            ))}
            {items.length === 0 && (
              <tr>
                <td colSpan={7} className="px-3 py-6 text-center text-muted-foreground">
                  No items parsed for this section.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
