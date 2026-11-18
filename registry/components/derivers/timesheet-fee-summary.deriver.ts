import { deriver } from "docxcelerate/template";

/**
 * Turns raw time entries into the rows, adjustments and subtotal of a fee
 * summary. Installed by `dxcl add timesheet-fee-summary`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level helper or constant would be a ReferenceError there. Every
 * helper lives inside `run`.
 */
export default deriver({
  name: "timesheetFeeSummary",
  run: (time: unknown) => {
    // deno-lint-ignore no-explicit-any
    const t = time as any;
    const fail = (message: string): never => {
      throw new Error(message);
    };
    const finite = (n: unknown): n is number => typeof n === "number" && Number.isFinite(n);
    const text = (s: unknown): s is string => typeof s === "string" && s.trim() !== "";
    const isDay = (s: unknown): boolean => {
      if (typeof s !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(s)) return false;
      const d = new Date(s.slice(0, 10) + "T00:00:00Z");
      return !Number.isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s.slice(0, 10);
    };
    const safe = (n: number): number => Number.isSafeInteger(n) ? n : fail("A figure is too large to count exactly");

    // Validate everything before pricing, so bad data throws even with no entries.
    if (t === null || typeof t !== "object") fail("time must be an object");
    const code: string = String(t.currency ?? "GBP").toUpperCase();
    if (!/^[A-Z]{3}$/.test(code)) fail("time.currency must be a three-letter ISO 4217 code");
    const digits: number = new Intl.NumberFormat("en", { style: "currency", currency: code })
      .resolvedOptions().maximumFractionDigits as number;
    const scale = 10 ** digits;
    const toMinor = (value: number, field: string): number => {
      const minor = Math.round(value * scale);
      if (Math.abs(value * scale - minor) > 1e-6) {
        fail(`${field} ${value} has more decimals than ${code} has`);
      }
      return minor;
    };
    const money = (minor: number): string =>
      new Intl.NumberFormat("en-GB", { minimumFractionDigits: digits, maximumFractionDigits: digits })
        .format(minor / scale);

    if (!Array.isArray(t.entries)) fail("time.entries must be an array");
    if (!Array.isArray(t.timekeepers)) fail("time.timekeepers must be an array");
    const billing = t.billing ?? {};
    const inc: number = billing.incrementMinutes ?? 6;
    if (!Number.isInteger(inc) || inc <= 0 || inc % 3 !== 0) {
      fail("time.billing.incrementMinutes must be a positive whole number of minutes that is a multiple of 3");
    }
    const rounding: string = billing.rounding ?? "up";
    if (rounding !== "up" && rounding !== "nearest") fail('time.billing.rounding must be "up" or "nearest"');
    const travelFactor: number = billing.travelFactor ?? 1;
    if (!finite(travelFactor) || travelFactor <= 0 || travelFactor > 1) {
      fail("time.billing.travelFactor must be greater than 0 and at most 1");
    }
    let writeDownMinor = 0;
    let writeDownReason = "";
    if (billing.writeDown !== undefined) {
      const w = billing.writeDown;
      if (!finite(w?.amount) || w.amount < 0) fail("time.billing.writeDown.amount must be a finite number, not negative");
      if (!text(w.reason)) fail("time.billing.writeDown.reason is required");
      writeDownMinor = toMinor(w.amount, "time.billing.writeDown.amount");
      writeDownReason = w.reason.trim();
    }

    interface Rate {
      from: string;
      minor: number;
    }
    interface Keeper {
      name: string;
      role: string;
      rates: Rate[];
      approved: number;
    }
    const keepers = new Map<string, Keeper & { index: number }>();
    // deno-lint-ignore no-explicit-any
    t.timekeepers.forEach((k: any, i: number) => {
      const at = `time.timekeepers[${i}]`;
      if (!text(k?.id)) fail(`${at}.id is required`);
      if (keepers.has(k.id)) fail(`${at}.id "${k.id}" is a duplicate`);
      if (!text(k.name)) fail(`${at}.name is required`);
      if (typeof k.role !== "string") fail(`${at}.role is required`);
      if (!Array.isArray(k.rates) || k.rates.length === 0) fail(`${at}.rates must be a non-empty array`);
      // deno-lint-ignore no-explicit-any
      const rates: Rate[] = k.rates.map((r: any, j: number) => {
        if (!isDay(r?.from)) fail(`${at}.rates[${j}].from must be YYYY-MM-DD`);
        if (!finite(r.rate) || r.rate < 0) fail(`${at}.rates[${j}].rate must be a finite number, not negative`);
        return { from: r.from.slice(0, 10), minor: toMinor(r.rate, `${at}.rates[${j}].rate`) };
      });
      if (new Set(rates.map((r) => r.from)).size !== rates.length) fail(`${at}.rates has a duplicate from date`);
      let approved = Infinity;
      if (k.approvedRate !== undefined) {
        if (!finite(k.approvedRate) || k.approvedRate < 0) fail(`${at}.approvedRate must be a finite number, not negative`);
        approved = toMinor(k.approvedRate, `${at}.approvedRate`);
      }
      keepers.set(k.id, { name: k.name, role: k.role, rates, approved, index: i });
    });

    const kinds = ["work", "travel", "no-charge"];
    const groups = new Map<string, { keeper: number; from: string; kind: number; minutes: number }>();
    // deno-lint-ignore no-explicit-any
    t.entries.forEach((e: any, i: number) => {
      const at = `time.entries[${i}]`;
      if (!isDay(e?.date)) fail(`${at}.date must be YYYY-MM-DD`);
      const keeper = keepers.get(e.timekeeperId);
      if (!keeper) fail(`${at}.timekeeperId "${e.timekeeperId}" is not in time.timekeepers`);
      if (!finite(e.minutes) || e.minutes <= 0) fail(`${at}.minutes must be a finite number greater than 0`);
      const kind = kinds.indexOf(e.kind ?? "work");
      if (kind < 0) fail(`${at}.kind "${e.kind}" must be work, travel or no-charge`);

      // Per entry, never per day or per total: they bill different hours.
      let billed = rounding === "up" ? Math.ceil(e.minutes / inc) * inc : Math.floor(e.minutes / inc + 0.5) * inc;
      billed = Math.max(billed, inc);

      const day = e.date.slice(0, 10);
      let rate: Rate | undefined;
      for (const r of keeper!.rates) if (r.from <= day && (!rate || r.from > rate.from)) rate = r;
      if (!rate) fail(`${at}.date ${day} is before the earliest rate of ${e.timekeeperId}`);
      const key = `${keeper!.index}|${rate!.from}|${kind}`;
      const group = groups.get(key) ?? { keeper: keeper!.index, from: rate!.from, kind, minutes: 0 };
      group.minutes += billed;
      groups.set(key, group);
    });

    const ordered = [...groups.values()].sort((a, b) =>
      a.keeper - b.keeper || (a.from < b.from ? -1 : a.from > b.from ? 1 : 0) || a.kind - b.kind
    );
    const periods = new Map<number, Set<string>>();
    for (const g of ordered) periods.set(g.keeper, (periods.get(g.keeper) ?? new Set()).add(g.from));
    const keeperList = [...keepers.values()];

    // Hours are exact: minutes are a multiple of 3, so hundredths are whole.
    const hours = (minutes: number): string => {
      const hundredths = minutes * 5 / 3;
      return new Intl.NumberFormat("en-GB").format(Math.floor(hundredths / 100)) + "." +
        String(hundredths % 100).padStart(2, "0");
    };
    const pct = (f: number): string => String(Number((f * 100).toFixed(2)));
    const day = (iso: string): string =>
      new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" }).format(new Date(iso + "T00:00:00Z"));
    // minutes * rate / 60, half up, in integers.
    const amountFor = (minutes: number, rate: number): number => safe(Math.floor(safe(minutes / 3 * rate + 10) / 20));

    let rowsMinor = 0;
    let reductionsMinor = 0;
    const adjustments: Array<{ label: string; amount: string }> = [];
    const rows = ordered.map((g) => {
      const k = keeperList[g.keeper];
      const rate = k.rates.find((r) => r.from === g.from)!;
      const f = g.kind === 0 ? 1 : g.kind === 1 ? travelFactor : 0;
      const printed = Math.round(rate.minor * f);
      const allowed = Math.round(Math.min(rate.minor, k.approved) * f);
      const qualifiers = [
        periods.get(g.keeper)!.size > 1 ? `rate from ${day(g.from)}` : "",
        g.kind === 1 ? `travel at ${pct(f)}%` : "",
        g.kind === 2 ? "no charge" : "",
      ].filter(Boolean);
      const label = qualifiers.length ? `${k.name} (${qualifiers.join(", ")})` : k.name;
      const amount = amountFor(g.minutes, printed);
      rowsMinor += amount;
      if (allowed < printed) {
        const cut = amount - amountFor(g.minutes, allowed);
        reductionsMinor += cut;
        adjustments.push({
          label: `Reduction per billing guidelines: ${label}, ${hours(g.minutes)} hours at ${money(allowed)}`,
          amount: money(-cut),
        });
      }
      return { label, role: k.role, hours: hours(g.minutes), rate: money(printed), amount: money(amount) };
    });

    if (writeDownMinor > 0) {
      if (writeDownMinor > rowsMinor - reductionsMinor) fail("time.billing.writeDown.amount is larger than the fees");
      adjustments.push({ label: `Write-down: ${writeDownReason}`, amount: money(-writeDownMinor) });
    }
    const subtotalMinor = rowsMinor - reductionsMinor - writeDownMinor;

    const hasRows = ordered.length > 0;
    const basis = hasRows
      ? `Each time entry is rounded ${rounding === "up" ? "up to the next" : "to the nearest"} ${inc} minutes (${
        hours(inc)
      } hours), with a minimum of ${inc} minutes per entry.` +
        (travelFactor < 1 && ordered.some((g) => g.kind === 1)
          ? ` Travel is billed at ${pct(travelFactor)}% of the hourly rate.`
          : "")
      : "";
    return { currency: code, hasRows, basis, rows, adjustments, subtotal: money(subtotalMinor), subtotalMinor };
  },
});
