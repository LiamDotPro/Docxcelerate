// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Turns a subscription's mid-period changes into credit and charge lines, and
 * decides what the net means. Installed by `dxcl add plan-change-proration`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "planChangeProration",
  run: (subscription: unknown) => {
   
    const sub = subscription as any;
    const empty = () => ({ description: "", period: "", amount: "", amountMinor: 0 });

    // Currency first: the exponent decides how every amount is read and printed.
    const currency = sub?.currency;
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency });
    } catch {
      throw new Error("subscription.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmtMoney = (minor: number) => money.format(minor / 10 ** exponent);

    // UTC day indexes, so a day does not shift with the engine's time zone.
    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    const dayIndex = (iso: string) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d) / 86_400_000;
    };
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const fmt = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));

    const startIso = isoDay(sub?.period?.start, "subscription.period.start");
    const endIso = isoDay(sub?.period?.end, "subscription.period.end");
    if (endIso < startIso) throw new Error("subscription.period.end is before subscription.period.start");
    const start = dayIndex(startIso);
    const end = dayIndex(endIso);
    const days = end - start + 1;

    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
        throw new Error(field + " must be a non-negative integer of minor units");
      }
      return value;
    };
    const labelOf = (side: any, field: string) => {
      if (typeof side?.label !== "string" || side.label.trim() === "") {
        throw new Error(field + ".label is required");
      }
      const q = side.quantity;
      if (q !== undefined && (typeof q !== "number" || !Number.isFinite(q) || q < 0)) {
        throw new Error(field + ".quantity must be a finite number of 0 or more");
      }
      return { label: side.label as string, quantity: q as number | undefined };
    };
    const nameOf = (l: { label: string; quantity: number | undefined }) =>
      l.quantity === undefined || l.quantity === 1
        ? l.label
        : new Intl.NumberFormat("en-GB").format(l.quantity) + " × " + l.label;

    // Half away from zero on exact integers; no daily rate is ever formed.
    const roundHalfAway = (numerator: number, denominator: number) =>
      Math.sign(numerator) * Math.floor((2 * Math.abs(numerator) + denominator) / (2 * denominator));

    const raw: unknown[] = Array.isArray(sub.changes) ? sub.changes : [];
    const kept: Array<{ i: number; at: string; day: number; item: string; c: any }> = [];
    raw.forEach((c: any, i) => {
      if (c?.metered === true) return;
      const at = isoDay(c?.at, "subscription.changes[" + i + "].at");
      const day = dayIndex(at);
      if (day <= start || day > end) return;
      const f = "subscription.changes[" + i + "]";
      labelOf(c.from, f + ".from");
      labelOf(c.to, f + ".to");
      minorOf(c.to?.periodAmountMinor, f + ".to.periodAmountMinor");
      kept.push({ i, at, day, item: c.item ?? "", c });
    });
    // Stable by construction: ties fall back to input order.
    kept.sort((a, b) => a.day - b.day || a.i - b.i);

    // One group per item and day: the first `from`, the last `to`.
    const groups: Array<{ item: string; at: string; day: number; i: number; first: any; last: any }> = [];
    const lastOf = new Map<string, (typeof groups)[number]>();
    for (const k of kept) {
      const g = lastOf.get(k.item);
      if (g && g.day === k.day) g.last = k.c;
      else {
        const created = { item: k.item, at: k.at, day: k.day, i: k.i, first: k.c, last: k.c };
        groups.push(created);
        lastOf.set(k.item, created);
      }
    }

    const prevOf = new Map<string, any>();
    const lines: Array<{ description: string; period: string; amount: string; amountMinor: number }> = [];
    for (const g of groups) {
      const remaining = end - g.day + 1;
      const prev = prevOf.get(g.item);
      const credited = prev
        ? { label: labelOf(prev, "to"), basis: prev.periodAmountMinor as number }
        : {
          label: labelOf(g.first.from, "from"),
          basis: minorOf(g.first.from.paidMinor, "subscription.changes[" + g.i + "].from.paidMinor"),
        };
      const charged = labelOf(g.last.to, "to");
      const creditMinor = -roundHalfAway(credited.basis * remaining, days);
      const chargeMinor = roundHalfAway(g.last.to.periodAmountMinor * remaining, days);
      prevOf.set(g.item, g.last.to);

      if (
        credited.label.label === charged.label && credited.label.quantity === charged.quantity &&
        creditMinor + chargeMinor === 0
      ) continue;

      const period = fmt(g.at) + " to " + fmt(endIso);
      lines.push(
        { description: "Unused time on " + nameOf(credited.label) + " from " + fmt(g.at), period, amount: fmtMoney(creditMinor), amountMinor: creditMinor },
        { description: "Remaining time on " + nameOf(charged) + " from " + fmt(g.at), period, amount: fmtMoney(chargeMinor), amountMinor: chargeMinor },
      );
    }
    // Groups were walked in date order but chained per item, so the lines
    // already come out ordered by date, then by first input index.
    const printed = lines.filter((l) => l.amountMinor !== 0);
    if (printed.length > 8) {
      throw new Error(
        "planChangeProration: " + printed.length + " proration lines in one period; this component prints at most 8",
      );
    }

    const netMinor = printed.reduce((sum, l) => sum + l.amountMinor, 0);
    const outcome = printed.length === 0
      ? "none"
      : netMinor > 0
      ? "upgrade"
      : netMinor < 0
      ? (sub.refundPolicy === "refund" ? "credit-refund" : "credit-balance")
      : "lateral";
    const showLines = outcome === "upgrade" || outcome === "credit-balance" || outcome === "credit-refund";
    const slots = showLines ? printed : [];
    while (slots.length < 8) slots.push(empty());

    return {
      outcome,
      showLines,
      periodLabel: outcome === "none" ? "" : fmt(startIso) + " to " + fmt(endIso),
      days: outcome === "none" ? 0 : days,
      lines: slots,
      net: outcome === "none" ? "" : fmtMoney(netMinor),
      netAbs: outcome === "none" ? "" : fmtMoney(Math.abs(netMinor)),
      netMinor,
    };
  },
});
