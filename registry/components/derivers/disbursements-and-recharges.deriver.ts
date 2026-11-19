// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Splits an invoice's expenses into recharges and disbursements, prices each
 * block, and writes the wording. Installed by `dxcl add disbursements-and-recharges`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper and list is declared inside `run`.
 */
export default deriver({
  name: "disbursementsAndRecharges",
  run: (invoice: unknown, strict: unknown) => {
    const inv = invoice as any;
    const strictMode = strict === true;
    const SLOTS_A = 8;
    const SLOTS_B = 6;
    const empty = () => ({ date: "", description: "", amount: "", amountMinor: 0 });
    const fill = (rows: any[], n: number) => {
      const slots = rows.slice();
      while (slots.length < n) slots.push(empty());
      return slots;
    };
    const nothing = () => ({
      any: false, hasRecharges: false, hasDisbursements: false, showVat: false,
      rechargeTitle: "", rechargeStatement: "", rechargeAmountHeading: "",
      recharges: fill([], SLOTS_A),
      rechargeNet: "", vatLabel: "", vat: "", rechargeTotal: "", rechargeTotalLabel: "",
      disbursementTitle: "", disbursementStatement: "",
      disbursements: fill([], SLOTS_B),
      disbursementTotal: "", disbursementTotalLabel: "",
      notes: "",
      rechargeNetMinor: 0, vatMinor: 0, rechargeTotalMinor: 0, disbursementTotalMinor: 0,
    });

    // Nothing to print means nothing else on `invoice` is read.
    const expenses = inv?.expenses;
    if (!Array.isArray(expenses) || expenses.length === 0) return nothing();

    const big = (n: number) => BigInt(n);
    // Half away from zero, symmetric for credits, on exact integers (d > 0).
    const roundHalfAway = (n: bigint, d: bigint): bigint => {
      const abs = n < big(0) ? -n : n;
      const q = (big(2) * abs + d) / (big(2) * d);
      return n < big(0) ? -q : q;
    };
    // A decimal becomes an integer at `dp` places, or throws: no float rounding survives.
    const scaled = (v: unknown, dp: number, field: string, what: string): number => {
      if (typeof v !== "number" || !Number.isFinite(v)) throw new Error(field + " must be a finite number");
      const s = v * 10 ** dp;
      const m = Math.round(s);
      if (Math.abs(s - m) > 1e-6) throw new Error(field + " has more decimals than " + what + " allows");
      return m;
    };

    const jurisdictions = ["UK", "EU", "DE", "FR", "US"];
    if (!jurisdictions.includes(inv.jurisdiction)) {
      throw new Error("invoice.jurisdiction must be one of " + jurisdictions.join(", "));
    }
    const us = inv.jurisdiction === "US";
    const locale = us ? "en-US" : "en-GB";

    const moneyOf = (code: unknown, field: string) => {
      if (typeof code !== "string") throw new Error(field + " must be an ISO 4217 code");
      const cur = code.toUpperCase();
      try {
        if (!Intl.supportedValuesOf("currency").includes(cur)) throw new Error("unknown");
        const nf = new Intl.NumberFormat(locale, { style: "currency", currency: cur });
        return { cur, nf, e: nf.resolvedOptions().maximumFractionDigits as number };
      } catch {
        throw new Error(field + " must be an ISO 4217 code");
      }
    };
    const home = moneyOf(inv.currency, "invoice.currency");
    const fmt = (m: { nf: Intl.NumberFormat; e: number }, minor: bigint | number) =>
      m.nf.format(Number(minor) / 10 ** m.e);

    const rate = inv.mainSupplyVatRate;
    let rateHundredths = 0;
    if (!us) {
      if (rate !== null) {
        if (typeof rate !== "number" || !Number.isFinite(rate) || rate < 0) {
          throw new Error("invoice.mainSupplyVatRate must be null or a number of 0 or more");
        }
        rateHundredths = scaled(rate, 2, "invoice.mainSupplyVatRate", "two places");
      }
    }

    // The firm's own costs: an item marked paid-as-agent in one of these is a recharge.
    // The electronic-search entry is Brabners (UK FTT, 2017). This is the list an owner edits.
    const ownCosts: Record<string, string> = {
      "travel": "travel",
      "subsistence": "subsistence",
      "accommodation": "accommodation",
      "postage": "postage",
      "telephone": "telephone calls",
      "photocopying": "photocopying",
      "printing": "printing",
      "electronic-search": "electronic search fees",
    };

    const dates = new Intl.DateTimeFormat(locale, { dateStyle: "medium", timeZone: "UTC" });
    type Item = { i: number; date: string; description: string; amount: bigint; note: string; block: "A" | "B" };
    const items: Item[] = [];

    expenses.forEach((x: any, i: number) => {
      const f = "invoice.expenses[" + i + "]";
      const date = typeof x?.date === "string" ? x.date.slice(0, 10) : "";
      const d = /^\d{4}-\d{2}-\d{2}$/.test(date) ? new Date(date + "T00:00:00Z") : null;
      if (!d || Number.isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== date) {
        throw new Error(f + ".date must be an ISO 8601 date");
      }
      if (typeof x.description !== "string" || x.description.trim() === "") {
        throw new Error(f + ".description is required");
      }
      const src = x.currency === undefined ? home : moneyOf(x.currency, f + ".currency");
      const nonNeg = (v: unknown, field: string) => {
        if (v === undefined) return 0;
        if (typeof v === "number" && v < 0) throw new Error(field + " must not be negative");
        return scaled(v, src.e, field, src.cur);
      };
      const net = scaled(x.net, src.e, f + ".net", src.cur);
      const reclaimed = nonNeg(x.inputVatReclaimed, f + ".inputVatReclaimed");
      const irrecoverable = nonNeg(x.irrecoverableTax, f + ".irrecoverableTax");
      let markup = 0;
      if (x.markupPercent !== undefined) {
        if (typeof x.markupPercent === "number" && x.markupPercent < 0) {
          throw new Error(f + ".markupPercent must not be negative");
        }
        markup = scaled(x.markupPercent, 2, f + ".markupPercent", "two places");
      }

      const costSrc = net + irrecoverable;
      const foreign = src.cur !== home.cur;
      let converted = big(costSrc);
      let fxText = "";
      if (foreign) {
        const fx = x.fxRate;
        if (typeof fx !== "number" || !Number.isFinite(fx) || fx <= 0) {
          throw new Error(f + ".fxRate must be a number above 0 when the currency differs");
        }
        const micros = scaled(fx, 6, f + ".fxRate", "six places");
        converted = roundHalfAway(
          big(costSrc) * big(micros) * big(10) ** big(home.e),
          big(10) ** big(6) * big(10) ** big(src.e),
        );
        fxText = ": " + fmt(src, costSrc) + " at " + String(fx) + " = " + fmt(home, converted);
      }
      const amount = roundHalfAway(converted * big(10000 + markup), big(10000));
      if (amount === big(0)) return;

      let block: "A" | "B";
      let note = "";
      const name = x.description as string;
      if (us) {
        block = x.taxable === true ? "A" : "B";
        if (x.taxable === undefined) {
          if (strictMode) {
            throw new Error(
              "disbursementsAndRecharges: " + f + " has no taxable flag; strict mode does not reclassify",
            );
          }
          note = name + ": listed as not taxable, because no taxability was given for it.";
        }
      } else {
        block = "A";
        if (x.paidAsAgent === true) {
          const own = typeof x.category === "string" && Object.hasOwn(ownCosts, x.category)
            ? ownCosts[x.category]
            : "";
          const reason = reclaimed > 0
            ? "we reclaimed the VAT on it"
            : markup > 0
            ? "it includes a handling charge"
            : own
            ? own + " is a cost of providing our services, not a payment made on your behalf"
            : "";
          if (reason) {
            if (strictMode) {
              throw new Error(
                "disbursementsAndRecharges: " + f + " is marked paidAsAgent but " + reason +
                  "; strict mode does not reclassify",
              );
            }
            note = name + ": recharged with VAT, not passed on as a disbursement, because " + reason + ".";
          } else block = "B";
        }
      }
      const supplier = typeof x.supplierName === "string" ? x.supplierName.trim() : "";
      items.push({
        i, date, amount, note, block,
        description: name + (supplier ? " (" + supplier + ")" : "") + fxText,
      });
    });

    // Date, then input order: the same key orders the lines and the notes.
    items.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.i - b.i));
    const row = (it: Item) => ({
      date: dates.format(new Date(it.date + "T00:00:00Z")),
      description: it.description,
      amount: fmt(home, it.amount),
      amountMinor: Number(it.amount),
    });
    const a = items.filter((it) => it.block === "A");
    const b = items.filter((it) => it.block === "B");
    if (a.length > SLOTS_A) {
      throw new Error(
        "disbursementsAndRecharges: " + a.length + " recharged expenses; this component prints at most " + SLOTS_A,
      );
    }
    if (b.length > SLOTS_B) {
      throw new Error(
        "disbursementsAndRecharges: " + b.length + " disbursements; this component prints at most " + SLOTS_B,
      );
    }
    if (items.length === 0) return nothing();

    const sum = (list: Item[]) => list.reduce((s, it) => s + it.amount, big(0));
    const netA = sum(a);
    const showVat = !us && rate !== null;
    const vat = showVat ? roundHalfAway(netA * big(rateHundredths), big(10000)) : big(0);
    const totalA = netA + vat;
    const totalB = sum(b);
    const hasRecharges = a.length > 0;
    const hasDisbursements = b.length > 0;

    const rateText = String(rate);
    const statementA = us
      ? "Sales tax applies to these expenses as part of our taxable services."
      : showVat
      ? "Costs we incurred in providing our services. VAT is charged on them at " + rateText +
        "%, the rate that applies to our fees."
      : "Costs we incurred in providing our services. They take the same VAT treatment as our fees, so no VAT is charged on them here.";

    return {
      any: hasRecharges || hasDisbursements,
      hasRecharges,
      hasDisbursements,
      showVat,
      rechargeTitle: us ? "Taxable reimbursed expenses" : "Recharged expenses",
      rechargeStatement: statementA,
      rechargeAmountHeading: us ? "Amount" : "Net",
      recharges: fill(a.map(row), SLOTS_A),
      rechargeNet: fmt(home, netA),
      vatLabel: showVat ? "VAT at " + rateText + "%" : "",
      vat: showVat ? fmt(home, vat) : "",
      rechargeTotal: fmt(home, totalA),
      rechargeTotalLabel: us ? "Total taxable expenses" : "Total recharged expenses",
      disbursementTitle: us ? "Non-taxable reimbursed expenses" : "Disbursements",
      disbursementStatement: us
        ? "No sales tax is charged on these expenses."
        : "Paid on your behalf as your agent and passed on at exact cost. These are outside the scope of VAT, so no VAT is charged on them.",
      disbursements: fill(b.map(row), SLOTS_B),
      disbursementTotal: fmt(home, totalB),
      disbursementTotalLabel: us ? "Total non-taxable expenses" : "Total disbursements",
      notes: items.filter((it) => it.note).map((it) => it.note).join("\n"),
      rechargeNetMinor: Number(netA),
      vatMinor: Number(vat),
      rechargeTotalMinor: Number(totalA),
      disbursementTotalMinor: Number(totalB),
    };
  },
});
