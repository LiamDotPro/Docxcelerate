// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Turns one period's ledger into printed rows, totals and the closing outcome.
 * Installed by `dxcl add statement-of-account`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "statementOfAccount",
  run: (statement: unknown) => {
    const st = statement as any;

    // Currency first: the exponent decides how every figure is rounded and printed.
    let fmt: Intl.NumberFormat;
    try {
      fmt = new Intl.NumberFormat("en-GB", { style: "currency", currency: st?.currency });
    } catch {
      throw new Error("statement.currency must be an ISO 4217 code");
    }
    const exponent = fmt.resolvedOptions().maximumFractionDigits as number;
    const f = 10 ** exponent;

    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    const finite = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isFinite(value)) throw new Error(field + " must be a finite number");
      return value;
    };
    const periodStart = isoDay(st?.periodStart, "statement.periodStart");
    const statementDate = isoDay(st?.statementDate, "statement.statementDate");
    if (statementDate < periodStart) throw new Error("statement.statementDate is before statement.periodStart");
    const opening = finite(st?.openingBalance, "statement.openingBalance");
    const given = st.closingBalance === undefined || st.closingBalance === null
      ? undefined
      : finite(st.closingBalance, "statement.closingBalance");

    const rank: Record<string, number> = { invoice: 0, "credit-note": 1, payment: 2, refund: 3, adjustment: 4 };
    const labels: Record<string, string> = {
      invoice: "Invoice",
      "credit-note": "Credit note",
      payment: "Payment",
      refund: "Refund",
      adjustment: "Adjustment",
    };
    const raw: any[] = Array.isArray(st.entries) ? st.entries : [];
    // Every entry is validated before any is filtered out: a bad one is a bad
    // ledger whether or not it falls in this period.
    const all = raw.map((e: any, i) => {
      const at = "statement.entries[" + i + "]";
      if (typeof e?.type !== "string" || !Object.hasOwn(rank, e.type)) throw new Error(at + ".type is not a known type");
      const date = isoDay(e.date, at + ".date");
      if (typeof e.number !== "string" || e.number.trim() === "") throw new Error(at + ".number is required");
      const amount = finite(e.amount, at + ".amount");
      if (amount < 0 && e.type !== "adjustment") {
        throw new Error(at + ".amount is negative; use a credit-note, payment or a credit adjustment");
      }
      if (e.currency !== undefined && e.currency !== null && e.currency !== st.currency) {
        throw new Error(at + ".currency differs from statement.currency; one statement, one currency");
      }
      return { i, e, type: e.type as string, date, number: e.number as string, amount };
    });
    const kept = all
      .filter((x) => x.date >= periodStart && x.date <= statementDate)
      .sort((a, b) =>
        a.date < b.date ? -1 : a.date > b.date ? 1 : rank[a.type] - rank[b.type] ||
          a.number.localeCompare(b.number, "en", { numeric: true }) || a.i - b.i
      );

    // Half away from zero to the minor unit; toPrecision strips binary noise (1.005 * 100).
    const roundMinor = (x: number) => Math.sign(x) * Math.round(Number((Math.abs(x) * f).toPrecision(15)));
    const money = (minor: number) => fmt.format(minor / f);
    // A credit balance prints as an amount with CR, never a minus sign.
    const bal = (x: number) => {
      const m = roundMinor(x);
      return m < 0 ? money(-m) + " CR" : money(m);
    };
    const medium = new Intl.DateTimeFormat("en-GB", { dateStyle: "medium", timeZone: "UTC" });
    const long = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const at = (d: Intl.DateTimeFormat, iso: string) => d.format(new Date(iso + "T00:00:00Z"));

    // The balance is carried unrounded so the last row always equals the closing line.
    let balance = opening;
    let totalDebits = 0;
    let totalCredits = 0;
    const rows = kept.map((x) => {
      const credit = x.type === "credit-note" || x.type === "payment" || (x.type === "adjustment" && x.amount < 0);
      const amount = Math.abs(x.amount);
      if (credit) totalCredits += amount;
      else totalDebits += amount;
      balance += credit ? -amount : amount;
      const cell = money(roundMinor(amount));
      let details = labels[x.type] + " " + x.number;
      if (x.e.reference) details += " (" + x.e.reference + ")";
      if (x.e.description) details += " – " + x.e.description;
      return {
        date: at(medium, x.date),
        details,
        debit: credit ? "" : cell,
        credit: credit ? cell : "",
        balance: bal(balance),
      };
    });
    const computedClosing = opening + totalDebits - totalCredits;
    if (given !== undefined && Math.abs(given - computedClosing) * f > 0.5 + 1e-9) {
      throw new Error(
        "statementOfAccount: closing balance " + given + " does not match opening + debits − credits = " + computedClosing,
      );
    }

    // The outcome comes from the rounded figure, so the sentence agrees with what is printed.
    const c = roundMinor(computedClosing);
    const outcome = c > 0 ? "owing" : c === 0 ? "clear" : "credit";
    const dueBy = st.dueBy === undefined || st.dueBy === null ? undefined : isoDay(st.dueBy, "statement.dueBy");

    return {
      hasEntries: kept.length > 0,
      outcome,
      periodLabel: at(long, periodStart) + " to " + at(long, statementDate),
      openingDate: at(medium, periodStart),
      openingDateLong: at(long, periodStart),
      openingBalance: bal(opening),
      rows,
      totalDebits: money(roundMinor(totalDebits)),
      totalCredits: money(roundMinor(totalCredits)),
      closingDate: at(medium, statementDate),
      closingBalance: bal(computedClosing),
      closingAbs: money(Math.abs(c)),
      dueSuffix: outcome === "owing" && dueBy ? ", payable by " + at(long, dueBy) : "",
      closingMinor: c,
    };
  },
});
