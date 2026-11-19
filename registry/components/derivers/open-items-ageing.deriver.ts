// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Every figure of an open-item statement: which invoices are still open, how
 * old each is, the six ageing buckets, the shares, and the sentences around
 * them. Installed by `dxcl add open-items-ageing`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "openItemsAgeing",
  run: (accountInput: unknown) => {
    const account = accountInput as any;

    // Currency first: its exponent decides how every amount is read and printed.
    const currency = account?.currency;
    let formatter: Intl.NumberFormat;
    try {
      formatter = new Intl.NumberFormat("en-GB", { style: "currency", currency });
    } catch {
      throw new Error("account.currency must be an ISO 4217 code");
    }
    const exponent = formatter.resolvedOptions().maximumFractionDigits as number;
    // Rounded once on the way in; everything after is integer minor units, so
    // the buckets sum exactly.
    const toMinor = (x: unknown, field: string): number => {
      if (typeof x !== "number" || !Number.isFinite(x)) throw new Error(field + " must be a finite number");
      return Math.round(x * 10 ** exponent);
    };
    const money = (minor: number) => formatter.format(minor / 10 ** exponent);

    // UTC day indexes, so a date does not shift with the engine's time zone.
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
    const shortDates = new Intl.DateTimeFormat("en-GB", { dateStyle: "medium", timeZone: "UTC" });
    const longDates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const short = (day: number) => shortDates.format(new Date(day * 86_400_000));
    const long = (iso: string) => longDates.format(new Date(iso + "T00:00:00Z"));

    const statementIso = isoDay(account?.statementDate, "account.statementDate");
    const statementDay = dayIndex(statementIso);

    const terms = account.defaultTermsDays;
    if (terms !== undefined && (typeof terms !== "number" || !Number.isInteger(terms) || terms < 0)) {
      throw new Error("account.defaultTermsDays must be a whole number of 0 or more");
    }

    const invoices: any[] = Array.isArray(account.invoices) ? account.invoices : [];
    const credits: any[] = Array.isArray(account.credits) ? account.credits : [];

    const byNumber = new Map<string, { i: number; amountMinor: number; paidMinor: number }>();
    const parsed = invoices.map((inv: any, i: number) => {
      const f = "account.invoices[" + i + "]";
      if (typeof inv?.number !== "string" || inv.number.trim() === "") throw new Error(f + ".number is required");
      if (byNumber.has(inv.number)) throw new Error("account.invoices has invoice " + inv.number + " twice");
      const amountMinor = toMinor(inv.amount, f + ".amount");
      if (amountMinor < 0) throw new Error(f + ".amount must not be negative; a credit note belongs in account.credits");
      const date = isoDay(inv.date, f + ".date");
      const dueDate = inv.dueDate === undefined || inv.dueDate === null ? undefined : isoDay(inv.dueDate, f + ".dueDate");
      const entry = { i, amountMinor, paidMinor: 0 };
      byNumber.set(inv.number, entry);
      return { i, number: inv.number as string, date, dueDate, entry, disputed: inv.disputed === true };
    });

    const parsedCredits = credits.map((c: any, i: number) => {
      const f = "account.credits[" + i + "]";
      if (c?.type !== "payment" && c?.type !== "credit-note") throw new Error(f + ".type must be payment or credit-note");
      const date = isoDay(c.date, f + ".date");
      if (typeof c.number !== "string" || c.number.trim() === "") throw new Error(f + ".number is required");
      const amountMinor = toMinor(c.amount, f + ".amount");
      if (amountMinor <= 0) throw new Error(f + ".amount must be above zero");
      const allocations: any[] = Array.isArray(c.allocations) ? c.allocations : [];
      let allocatedMinor = 0;
      allocations.forEach((a: any, j: number) => {
        const g = f + ".allocations[" + j + "]";
        const target = byNumber.get(a?.invoice);
        if (!target) throw new Error(g + " names invoice " + a?.invoice + ", which is not in account.invoices");
        const minor = toMinor(a.amount, g + ".amount");
        if (minor <= 0) throw new Error(g + ".amount must be above zero");
        allocatedMinor += minor;
        target.paidMinor += minor;
      });
      if (allocatedMinor > amountMinor) throw new Error("credit " + c.number + " allocates more than its amount");
      return { i, type: c.type as string, date, number: c.number as string, unallocatedMinor: amountMinor - allocatedMinor };
    });

    const open = parsed.flatMap((p) => {
      const openMinor = p.entry.amountMinor - p.entry.paidMinor;
      if (openMinor < 0) {
        throw new Error(
          "invoice " + p.number + " is over-allocated: " + money(p.entry.paidMinor) + " allocated against " +
            money(p.entry.amountMinor),
        );
      }
      return openMinor === 0 ? [] : [{ ...p, openMinor }];
    });

    let usedDefault = false;
    const aged = open.map((p) => {
      let dueDay: number;
      if (p.dueDate !== undefined) dueDay = dayIndex(p.dueDate);
      else {
        if (terms === undefined) {
          throw new Error("invoice " + p.number + " has no dueDate and account.defaultTermsDays is missing");
        }
        usedDefault = true;
        dueDay = dayIndex(p.date) + terms;
      }
      const age = statementDay - dueDay;
      // First match wins; each bucket includes its top edge.
      const bucket = p.disputed ? 5 : age <= 0 ? 0 : age <= 30 ? 1 : age <= 60 ? 2 : age <= 90 ? 3 : 4;
      const ageLabel = p.disputed ? "In dispute" : age <= 0 ? "Not yet due" : age === 1 ? "1 day" : age + " days";
      return { ...p, dueDay, age, bucket, ageLabel };
    });
    // Oldest debt first; the index tiebreak makes the order explicit.
    aged.sort((a, b) => a.dueDay - b.dueDay || dayIndex(a.date) - dayIndex(b.date) || a.i - b.i);

    const labels = ["Not yet due", "1–30 days", "31–60 days", "61–90 days", "Over 90 days", "In dispute"];
    const sums = [0, 0, 0, 0, 0, 0];
    for (const p of aged) sums[p.bucket] += p.openMinor;
    const totalOpen = sums.reduce((a, b) => a + b, 0);
    const overdueMinor = sums[1] + sums[2] + sums[3] + sums[4];
    const unallocatedList = parsedCredits.filter((c) => c.unallocatedMinor > 0)
      .sort((a, b) => dayIndex(a.date) - dayIndex(b.date) || a.i - b.i);
    const unallocatedMinor = unallocatedList.reduce((a, c) => a + c.unallocatedMinor, 0);
    const balanceMinor = totalOpen - unallocatedMinor;

    // Largest remainder, in integers, so the shares sum to exactly 100.
    const percents = sums.map((b) => (totalOpen > 0 ? Math.floor(b * 100 / totalOpen) : 0));
    if (totalOpen > 0) {
      const left = 100 - percents.reduce((a, b) => a + b, 0);
      sums.map((b, index) => ({ index, rem: (b * 100) % totalOpen, b }))
        .filter((x) => x.b > 0)
        .sort((a, b) => b.rem - a.rem || a.index - b.index)
        .slice(0, left)
        .forEach((x) => percents[x.index]++);
    }

    const overdue = overdueMinor > 0;
    const empty = aged.length === 0 && unallocatedMinor === 0;
    // An empty account prints its one line and nothing else.
    const nothingOverdue = overdue || empty
      ? ""
      : sums[5] > 0
      ? "Nothing on this account is overdue, leaving aside " + money(sums[5]) + " in dispute."
      : "Nothing on this account is overdue.";
    const basis = "Aged by days past each invoice's due date, at " + long(statementIso) + "." +
      (usedDefault
        ? " An invoice with no due date is counted as due " + terms + " " + (terms === 1 ? "day" : "days") +
          " after its invoice date."
        : "");
    const hasUnallocated = unallocatedMinor > 0;
    const detail = new Intl.ListFormat("en-GB", { type: "conjunction" }).format(
      unallocatedList.map((c) =>
        (c.type === "payment" ? "payment" : "credit note") + " " + c.number + " of " + long(c.date) + " (" +
        money(c.unallocatedMinor) + ")"
      ),
    );

    return {
      empty,
      emptyLine: empty ? "There are no open items on this account at " + long(statementIso) + "." : "",
      basis,
      hasItems: aged.length > 0,
      items: aged.map((p) => ({
        number: p.number,
        date: short(dayIndex(p.date)),
        due: short(p.dueDay),
        age: p.ageLabel,
        amount: money(p.entry.amountMinor),
        open: money(p.openMinor),
        disputed: p.disputed,
        ageDays: p.age,
        openMinor: p.openMinor,
      })),
      overdue,
      nothingOverdue,
      buckets: sums.map((amountMinor, i) => ({
        label: labels[i],
        amount: money(amountMinor),
        share: percents[i] + "%",
        amountMinor,
        percent: percents[i],
      })),
      totalOpen: money(totalOpen),
      overdueTotal: money(overdueMinor),
      hasUnallocated,
      unallocated: money(-unallocatedMinor),
      unallocatedDetail: hasUnallocated ? "Not yet set against an invoice: " + detail + "." : "",
      balanceLabel: balanceMinor > 0 ? "Balance due" : balanceMinor < 0 ? "In credit" : "Balance",
      balance: money(Math.abs(balanceMinor)),
      totalOpenMinor: totalOpen,
      overdueMinor,
      unallocatedMinor,
      balanceMinor,
    };
  },
});
