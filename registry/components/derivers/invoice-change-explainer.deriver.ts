// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Explains why one invoice differs from the one before it. Installed by
 * `dxcl add invoice-change-explainer`.
 *
 * Self-contained on purpose: it travels to the engine as source text, so a
 * module-level helper or constant would be a ReferenceError there. Every figure
 * is an integer in the currency's minor units from the moment it is computed,
 * and remainders are taken by subtraction, never rounded on their own, which is
 * what makes the causes add up to the difference exactly.
 */
export default deriver({
  name: "invoiceChangeExplainer",
  run: (change: unknown, locale: unknown) => {
   
    const c = change as any;
    const loc = typeof locale === "string" && locale !== "" ? locale : "en-GB";
    const KINDS = ["recurring", "usage", "one-off", "discount", "tax"];

    const silent = (silentReason: string, periodLabel: string) => ({
      speak: false,
      silentReason,
      direction: null,
      difference: "",
      periodLabel,
      causes: [] as never[],
      omitted: 0,
      sentence: "",
      facts: "",
    });

    if (typeof c?.currency !== "string" || c.currency === "") {
      throw new Error("invoiceChange.currency is required");
    }
    const currency: string = c.currency;
    let digits: number;
    try {
      digits = new Intl.NumberFormat("en", { style: "currency", currency }).resolvedOptions().maximumFractionDigits ?? 2;
    } catch {
      throw new Error("invoiceChange.currency is not a valid ISO 4217 code: " + currency);
    }
    const scale = 10 ** digits;
    const toMinor = (x: number) => Math.round(x * scale);

    const label = typeof c.periodLabel === "string" && c.periodLabel.trim() !== ""
      ? c.periodLabel.trim()
      : "your last invoice";

    const th = c.threshold ?? {};
    const absolute = th.absolute ?? 1;
    const pct = th.pct ?? 0.02;
    if (typeof absolute !== "number" || !Number.isFinite(absolute) || absolute < 0) {
      throw new Error("invoiceChange.threshold.absolute must be a number of 0 or more");
    }
    if (typeof pct !== "number" || !Number.isFinite(pct) || pct < 0 || pct >= 1) {
      throw new Error("invoiceChange.threshold.pct must be at least 0 and below 1");
    }

    const optNumber = (v: unknown, field: string) => {
      if (v !== undefined && v !== null && (typeof v !== "number" || !Number.isFinite(v))) {
        throw new Error(field + " must be a finite number");
      }
    };

    // Checks one invoice and returns its total net of the balance brought forward.
    const check = (snap: any, name: string): number => {
      if (snap === null || typeof snap !== "object") throw new Error(name + " is required");
      if (typeof snap.total !== "number" || !Number.isFinite(snap.total)) {
        throw new Error(name + ".total must be a finite number");
      }
      optNumber(snap.broughtForward, name + ".broughtForward");
      const pd = snap.periodDays;
      if (pd !== undefined && pd !== null && (!Number.isInteger(pd) || pd <= 0)) {
        throw new Error(name + ".periodDays must be a positive integer");
      }
      if (!Array.isArray(snap.lines)) throw new Error(name + ".lines must be an array");
      const seen = new Set<string>();
      let sum = 0;
      for (const [i, l] of snap.lines.entries()) {
        const at = name + ".lines[" + i + "]";
        if (typeof l?.key !== "string" || l.key === "") throw new Error(at + ".key is required");
        if (seen.has(l.key)) throw new Error(at + ".key is duplicated: " + l.key);
        seen.add(l.key);
        if (typeof l.description !== "string" || l.description === "") {
          throw new Error(at + ".description is required");
        }
        if (!KINDS.includes(l.kind)) throw new Error(at + ".kind must be one of " + KINDS.join(", "));
        if (typeof l.amount !== "number" || !Number.isFinite(l.amount)) {
          throw new Error(at + ".amount must be a finite number");
        }
        optNumber(l.quantity, at + ".quantity");
        optNumber(l.unitPrice, at + ".unitPrice");
        optNumber(l.taxRate, at + ".taxRate");
        sum += toMinor(l.amount);
      }
      const net = toMinor(snap.total) - toMinor(snap.broughtForward ?? 0);
      if (net !== sum) throw new Error(name + ": lines do not add up to total less broughtForward");
      return net;
    };

    const cur = c.current;
    const curNet = check(cur, "invoiceChange.current");

    const prev = c.previous;
    if (prev === undefined || prev === null) return silent("first-invoice", label);
    if (cur.documentType === "credit-note" || prev.documentType === "credit-note") {
      return silent("credit-note", label);
    }
    if (typeof prev.currency === "string" && prev.currency !== "" && prev.currency !== currency) {
      return silent("currency-change", label);
    }
    const prevNet = check(prev, "invoiceChange.previous");

    const delta = curNet - prevNet;
    const floor = Math.max(Math.round(absolute * scale), Math.round(pct * Math.abs(prevNet)));
    if (delta === 0 || Math.abs(delta) < floor) return silent("below-threshold", label);

    // Causes, in the fixed order that breaks ties.
    type Cause = { kind: string; key: string; description: string | null; effect: number; clause: string };
    type Raw = { kind: string; key: string; description: string | null; effect: number; clause: (e: string) => string };
    const raws: Raw[] = [];
    const add = (kind: string, key: string, description: string | null, effect: number, clause: (e: string) => string) => {
      if (effect !== 0) raws.push({ kind, key, description, effect, clause });
    };

    const money = new Intl.NumberFormat(loc, { style: "currency", currency });
    const signed = new Intl.NumberFormat(loc, { style: "currency", currency, signDisplay: "always" });
    const unit = new Intl.NumberFormat(loc, {
      style: "currency",
      currency,
      minimumFractionDigits: digits,
      maximumFractionDigits: Math.max(digits, 6),
    });
    const qty = new Intl.NumberFormat(loc, { maximumFractionDigits: 6 });
    const rate = new Intl.NumberFormat(loc, { style: "percent", maximumFractionDigits: 2 });
    const m = (minor: number) => money.format(minor / scale);
    const s = (minor: number) => signed.format(minor / scale);
    const move = (up: boolean) => (up ? "rose" : "fell");

    const before = new Map<string, any>();
    for (const l of prev.lines) before.set(l.key, l);
    const after = new Set<string>();
    const d0: number | undefined = prev.periodDays ?? undefined;
    const d1: number | undefined = cur.periodDays ?? undefined;
    let periodEffect = 0;

    for (const l of cur.lines) {
      after.add(l.key);
      const o = before.get(l.key);
      const d: string = l.description;
      const a1 = toMinor(l.amount);
      if (!o) {
        const kind = l.kind === "one-off" ? "one-off" : l.kind === "discount" ? "discount-started" : "new";
        add(kind, l.key, d, a1, (e) =>
          kind === "one-off"
            ? d + " is a one-off charge on this invoice (" + e + ")"
            : kind === "discount-started"
            ? d + " was applied (" + e + ")"
            : d + " was added (" + e + ")");
        continue;
      }
      const a0 = toMinor(o.amount);
      const delta2 = a1 - a0;
      const changed = () =>
        add("changed", l.key, d, delta2, (e) => d + " " + move(delta2 > 0) + " from " + m(a0) + " to " + m(a1) + " (" + e + ")");
      if (o.kind !== l.kind) {
        changed();
      } else if (l.kind === "tax") {
        const r0 = o.taxRate, r1 = l.taxRate;
        if (!(r0 > 0 && r1 > 0)) {
          changed();
        } else if (r0 === r1) {
          add("tax-base", l.key, d, delta2, (e) => d + " " + move(delta2 > 0) + " with the charges it is worked out on (" + e + ")");
        } else {
          const base = Math.round((l.amount / r1 - o.amount / r0) * r0 * scale);
          const byRate = delta2 - base;
          add("tax-base", l.key, d, base, (e) => d + " " + move(base > 0) + " with the charges it is worked out on (" + e + ")");
          add("tax-rate", l.key, d, byRate, (e) => d + " changed from " + rate.format(r0) + " to " + rate.format(r1) + " (" + e + ")");
        }
      } else if (
        typeof l.quantity === "number" && typeof l.unitPrice === "number" &&
        typeof o.quantity === "number" && typeof o.unitPrice === "number"
      ) {
        const q0: number = o.quantity, q1: number = l.quantity, p0: number = o.unitPrice, p1: number = l.unitPrice;
        const volume = p0 === p1 ? delta2 : q0 === q1 ? 0 : Math.round((q1 - q0) * p0 * scale);
        const price = delta2 - volume;
        add("price", l.key, d, price, (e) => "the price of " + d + " went from " + unit.format(p0) + " to " + unit.format(p1) + " (" + e + ")");
        let period = 0;
        if (l.kind === "usage" && d0 !== undefined && d1 !== undefined && d0 !== d1 && q0 > 0) {
          const perDay0 = q0 / d0, perDay1 = q1 / d1;
          period = Math.abs(perDay1 - perDay0) <= 0.01 * perDay0
            ? volume
            : Math.round((d1 - d0) * perDay0 * p0 * scale);
        }
        periodEffect += period;
        add("quantity", l.key, d, volume - period, (e) => d + " " + move(q1 > q0) + " from " + qty.format(q0) + " to " + qty.format(q1) + " (" + e + ")");
      } else {
        changed();
      }
    }

    for (const o of prev.lines) {
      if (after.has(o.key)) continue;
      const d: string = o.description;
      const kind = o.kind === "one-off" ? "one-off-ended" : o.kind === "discount" ? "discount-ended" : "removed";
      add(kind, o.key, d, -toMinor(o.amount), (e) =>
        kind === "one-off-ended"
          ? d + " from " + label + " is not repeated (" + e + ")"
          : kind === "discount-ended"
          ? d + " has ended (" + e + ")"
          : d + " is no longer charged (" + e + ")");
    }
    add("period-length", "period-length", null, periodEffect, (e) =>
      "this period had " + d1 + " days against " + d0 + " (" + e + ")");

    const sorted = raws.slice().sort((x, y) => Math.abs(y.effect) - Math.abs(x.effect));
    const gross = sorted.reduce((t, r) => t + Math.abs(r.effect), 0);
    const chosen: Cause[] = [];
    let cumulative = 0;
    for (const r of sorted) {
      chosen.push({ kind: r.kind, key: r.key, description: r.description, effect: r.effect, clause: r.clause(s(r.effect)) });
      cumulative += Math.abs(r.effect);
      if (10 * cumulative >= 9 * gross || chosen.length === 3) break;
    }
    const omitted = sorted.length - chosen.length;

    const difference = m(Math.abs(delta));
    const direction = delta > 0 ? "higher" : "lower";
    const clauses = chosen.map((x) => x.clause);
    const list = new Intl.ListFormat(loc, { type: "conjunction" }).format(clauses);
    const rest = omitted > 0 ? ", with " + omitted + " smaller change" + (omitted === 1 ? "" : "s") + " making up the rest" : "";
    const sentence = "Your total is " + difference + " " + direction + " than " + label + ", " +
      (omitted === 0 ? "because " : "mostly because ") + list + rest + ".";
    const facts = "The total is " + difference + " " + direction + " than " + label +
      ". What changed, largest effect first: " + clauses.join("; ") + "." +
      (omitted > 0
        ? " " + omitted + " smaller change" + (omitted === 1 ? " accounts" : "s account") + " for the rest; do not describe them."
        : "");

    return { speak: true, silentReason: null, direction, difference, periodLabel: label, causes: chosen, omitted, sentence, facts };
  },
});
