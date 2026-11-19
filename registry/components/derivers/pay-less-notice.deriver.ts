// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Settles a UK pay less notice: the deadline, whether service was in time, the
 * itemised schedule, the sum due and the facts the generated basis may use.
 * Installed by `dxcl add pay-less-notice`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "payLessNotice",
  run: (payLess: unknown) => {
    const pl = payLess as any;
    const fail = (message: string): never => {
      throw new Error("payLessNotice: " + message);
    };
    const gbp = new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" });
    const money = (minor: number) => gbp.format(minor / 100);
    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        return fail(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    // UTC day indexes, so a day does not shift with the engine's time zone.
    const dayIndex = (iso: string) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d) / 86_400_000;
    };
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const fmt = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));
    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0) {
        return fail(field + " must be a non-negative integer of pence");
      }
      return value;
    };
    const nonBlank = (value: unknown) => typeof value === "string" && value.trim() !== "";

    if (pl === null || typeof pl !== "object") fail("payLess is required");

    // Notified sum and its source.
    const notifiedMinor = minorOf(pl.notifiedSum?.amountMinor, "payLess.notifiedSum.amountMinor");
    const sources: Record<string, string> = {
      "payment-notice": "the payment notice",
      "application": "the payee's payment application, which stands as the notice of the sum due",
      "default-notice": "the payee's default payment notice",
    };
    const source = pl.notifiedSum?.source;
    if (typeof source !== "string" || !Object.prototype.hasOwnProperty.call(sources, source)) {
      fail("payLess.notifiedSum.source must be payment-notice, application or default-notice");
    }
    const reference = pl.notifiedSum.reference;
    const notifiedSource = sources[source] + (nonBlank(reference) ? " (" + reference.trim() + ")" : "");

    // Dates.
    const finalIso = isoDay(pl.finalDateForPayment, "payLess.finalDateForPayment");
    const servedIso = isoDay(pl.servedDate, "payLess.servedDate");
    let deadlineIso: string;
    if (pl.payLessDeadline !== undefined && pl.payLessDeadline !== null) {
      deadlineIso = isoDay(pl.payLessDeadline, "payLess.payLessDeadline");
      if (deadlineIso > finalIso) fail("payLess.payLessDeadline is after payLess.finalDateForPayment");
    } else {
      const n = pl.prescribedPeriodDays;
      if (typeof n !== "number" || !Number.isSafeInteger(n) || n < 1) {
        fail("one of payLess.payLessDeadline or payLess.prescribedPeriodDays (a positive integer) is required");
      }
      const excludedRaw: unknown[] = pl.excludedDays === undefined || pl.excludedDays === null
        ? []
        : Array.isArray(pl.excludedDays)
        ? pl.excludedDays
        : fail("payLess.excludedDays must be an array of ISO 8601 dates");
      const excluded = new Set<number>(
        excludedRaw.map((e, i) => dayIndex(isoDay(e, "payLess.excludedDays[" + i + "]"))),
      );
      let d = dayIndex(finalIso);
      let count = 0;
      while (count < n) {
        if (!excluded.has(d)) count++;
        d -= 1;
      }
      deadlineIso = new Date(d * 86_400_000).toISOString().slice(0, 10);
    }
    const timing = servedIso <= deadlineIso ? "in-time" : "late";

    // Lines: LDs first, then each kind in input order.
    const lines: Array<{ kind: string; description: string; amount: string; minor: number }> = [];
    const ld = pl.ld;
    if (ld !== undefined && ld !== null) {
      const rate = ld.rateMinor;
      if (typeof rate !== "number" || !Number.isSafeInteger(rate) || rate < 1) {
        fail("payLess.ld.rateMinor must be a positive integer of pence");
      }
      if (ld.per !== "day" && ld.per !== "week") fail("payLess.ld.per must be day or week");
      const delay = ld.delayDays;
      if (typeof delay !== "number" || !Number.isSafeInteger(delay) || delay < 0) {
        fail("payLess.ld.delayDays must be a non-negative integer");
      }
      const cap = ld.capMinor === undefined || ld.capMinor === null
        ? undefined
        : minorOf(ld.capMinor, "payLess.ld.capMinor");
      const previous = ld.previouslyDeductedMinor === undefined || ld.previouslyDeductedMinor === null
        ? 0
        : minorOf(ld.previouslyDeductedMinor, "payLess.ld.previouslyDeductedMinor");
      if (delay > 0) {
        const product = rate * delay;
        if (!Number.isSafeInteger(product)) fail("payLess.ld rate times delayDays is not a safe integer");
        // The only rounding in the component: pro rata by the day, once, in pence.
        const computed = ld.per === "day" ? product : Math.floor((2 * product + 7) / 14);
        const headroom = cap === undefined ? Infinity : Math.max(0, cap - previous);
        const applied = Math.min(computed, headroom);
        let description = delay + (delay === 1 ? " day" : " days") + " of delay at " + money(rate) +
          (ld.per === "day" ? " per day" : " per week, pro rata by the day");
        if (cap !== undefined && headroom === 0) {
          description += "; nothing deducted, because earlier notices have already deducted " + money(previous) +
            " against the " + money(cap) + " cap";
        } else if (applied < computed) {
          description += "; " + money(computed) + " capped at " + money(applied) + ", the " + money(cap as number) +
            " cap less " + money(previous) + " already deducted in earlier notices";
        }
        lines.push({ kind: "Liquidated damages", description, amount: money(applied), minor: applied });
      }
    }

    const raw: unknown[] = pl.deductions === undefined || pl.deductions === null
      ? []
      : Array.isArray(pl.deductions)
      ? pl.deductions
      : fail("payLess.deductions must be an array");
    const labels: Record<string, string> = { contra: "Contra-charge", abatement: "Abatement", setoff: "Set-off" };
    const items = raw.map((entry: any, i) => {
      const f = "payLess.deductions[" + i + "]";
      if (entry?.kind === "cis") {
        fail(
          f + " is a CIS deduction. CIS withholding is not a ground for a pay less notice; " +
            "show it on the CIS payment and deduction statement instead.",
        );
      }
      if (typeof entry?.kind !== "string" || !Object.prototype.hasOwnProperty.call(labels, entry.kind)) {
        fail(f + ".kind must be contra, abatement or setoff");
      }
      if (!nonBlank(entry.description)) fail(f + ".description is required");
      const minor = minorOf(entry.amountMinor, f + ".amountMinor");
      return { entry, minor };
    });
    for (const kind of ["contra", "abatement", "setoff"]) {
      for (const { entry, minor } of items) {
        if (entry.kind !== kind || minor === 0) continue;
        lines.push({
          kind: labels[kind],
          description: entry.description.trim() + (nonBlank(entry.reference) ? " (ref: " + entry.reference.trim() + ")" : ""),
          amount: money(minor),
          minor,
        });
      }
    }
    if (lines.length > 12) fail(lines.length + " deduction lines; this component prints at most 12");

    const total = lines.reduce((sum, l) => sum + l.minor, 0);
    const net = notifiedMinor - total;
    const sumDue = net > 0 ? money(net) : "nil (£0.00)";
    const excess = net < 0 ? money(-net) : "";
    const notified = money(notifiedMinor);
    const totalDeductions = money(total);

    const basisFacts = "Notified sum: " + notified + " (" + notifiedSource + "). " +
      (lines.length === 0
        ? "No deductions. "
        : lines.map((l) => l.kind + ": " + l.amount + " — " + l.description + ". ").join("")) +
      "Total deductions: " + totalDeductions + ". Sum due on the date of service: " + sumDue + "." +
      (excess
        ? " The deductions exceed the notified sum by " + excess + "; the excess is not recovered through this notice."
        : "");

    const slots = lines.map((l) => ({ kind: l.kind, description: l.description, amount: l.amount }));
    while (slots.length < 12) slots.push({ kind: "", description: "", amount: "" });

    return {
      timing,
      deadline: fmt(deadlineIso),
      served: fmt(servedIso),
      finalDate: fmt(finalIso),
      notifiedSource,
      notified,
      lines: slots,
      totalDeductions,
      sumDue,
      excess,
      basisFacts,
    };
  },
});
