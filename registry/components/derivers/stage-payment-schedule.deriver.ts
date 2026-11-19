// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Settles a contract's stage payments: the gross, deposit deduction, net, VAT
 * and tax point of every stage, and the totals. Installed by
 * `dxcl add stage-payment-schedule`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "stagePaymentSchedule",
  run: (contract: unknown) => {
    const c = contract as any;
    const SLOTS = 8;

    const currency = c?.currency ?? "GBP";
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency });
    } catch {
      throw new Error("contract.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmtMoney = (minor: number) => money.format(minor / 10 ** exponent);

    // UTC so a date does not shift with the engine's time zone.
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "medium", timeZone: "UTC" });
    const fmtDate = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));
    const isoDate = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      const iso = value.slice(0, 10);
      const d = new Date(iso + "T00:00:00Z");
      if (Number.isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== iso) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return iso;
    };
    // Basis points: a percentage to two decimal places, as an exact integer.
    const basisPoints = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isFinite(value)) throw new Error(field + " must be a number");
      const bp = Math.round(value * 100);
      if (Math.abs(bp - value * 100) > 1e-9) throw new Error(field + " must have at most two decimal places");
      return bp;
    };
    // Half up, once per line. BigInt because minor units times basis points can pass 2^53.
    const vatOf = (net: number, rateBp: number) => Number((BigInt(net) * BigInt(rateBp) + 5000n) / 10000n);

    const S = c?.contractSumMinor;
    if (typeof S !== "number" || !Number.isSafeInteger(S) || S <= 0) {
      throw new Error("contract.contractSumMinor must be a positive integer of minor units");
    }

    let rateBp = 0;
    const hasVat = c.vat !== undefined && c.vat !== null;
    if (hasVat) {
      if (c.vat.jurisdiction !== "UK") throw new Error('contract.vat.jurisdiction must be "UK"');
      rateBp = basisPoints(c.vat.ratePercent, "contract.vat.ratePercent");
      if (rateBp < 0 || rateBp > 10000) throw new Error("contract.vat.ratePercent must be between 0 and 100");
    }
    const single = c.singlePayment === true;
    const taxPoints = hasVat && !single;

    const raw = c.stages;
    if (!Array.isArray(raw) || raw.length === 0) throw new Error("contract.stages must have at least one stage");
    if (raw.length > SLOTS) {
      throw new Error(
        "contract.stages has more than " + SLOTS + " stages; the node prints " + SLOTS +
          ". Add row slots to the node and raise the limit in the deriver.",
      );
    }
    if (single && raw.length !== 1) throw new Error("contract.singlePayment requires exactly one stage");
    const n = raw.length;

    let D = 0;
    let receivedDate = "";
    let application = "";
    if (c.deposit !== undefined && c.deposit !== null) {
      D = c.deposit.amountMinor;
      if (typeof D !== "number" || !Number.isSafeInteger(D) || D < 0) {
        throw new Error("contract.deposit.amountMinor must be a non-negative integer of minor units");
      }
      if (D > S) throw new Error("contract.deposit.amountMinor is more than the contract sum");
      if (D > 0) {
        receivedDate = isoDate(c.deposit.receivedDate, "contract.deposit.receivedDate");
        application = c.deposit.application;
        if (application !== "first" && application !== "final" && application !== "proRata") {
          throw new Error('contract.deposit.application must be "first", "final" or "proRata"');
        }
      }
    }

    // Validate every stage and total what is stated.
    const states = ["pending", "reached", "invoiced", "paid"];
    let fixedTotal = 0;
    let bpTotal = 0;
    let balanceAt = -1;
    const stages = raw.map((s: any, i: number) => {
      const f = "contract.stages[" + i + "]";
      if (typeof s?.label !== "string" || s.label.trim() === "") throw new Error(f + ".label is required");
      if (!states.includes(s.state)) throw new Error(f + ".state must be pending, reached, invoiced or paid");
      const given = [s.percent !== undefined, s.amountMinor !== undefined, s.balance !== undefined].filter(Boolean).length;
      if (given !== 1) throw new Error(f + " must give exactly one of percent, amountMinor or balance");
      let bp = 0;
      let fixed = -1;
      if (s.percent !== undefined) {
        bp = basisPoints(s.percent, f + ".percent");
        if (bp <= 0 || bp > 10000) throw new Error(f + ".percent must be above 0 and at most 100");
        bpTotal += bp;
      } else if (s.amountMinor !== undefined) {
        fixed = s.amountMinor;
        if (typeof fixed !== "number" || !Number.isSafeInteger(fixed) || fixed <= 0) {
          throw new Error(f + ".amountMinor must be a positive integer of minor units");
        }
        fixedTotal += fixed;
      } else {
        if (s.balance !== true) throw new Error(f + ".balance must be true");
        if (i !== n - 1) throw new Error(f + ".balance is allowed on the last stage only");
        balanceAt = i;
      }
      const reached = s.reachedDate === undefined ? "" : isoDate(s.reachedDate, f + ".reachedDate");
      const invoiced = s.invoiceDate === undefined ? "" : isoDate(s.invoiceDate, f + ".invoiceDate");
      const paid = s.paidDate === undefined ? "" : isoDate(s.paidDate, f + ".paidDate");
      if (s.state === "invoiced" && !invoiced) throw new Error(f + ".invoiceDate is required when the stage is invoiced");
      if (s.state === "paid" && !paid) throw new Error(f + ".paidDate is required when the stage is paid");
      return { label: s.label as string, state: s.state as string, bp, fixed, reached, invoiced, paid };
    });

    // Allocation: every message states the excess in money.
    if (bpTotal > 10000) throw new Error("contract.stages: percentages total more than 100");
    if (fixedTotal > S) throw new Error("contract.stages: fixed amounts are " + fmtMoney(fixedTotal - S) + " more than the contract sum");
    const floorBp = (bp: number) => Number((BigInt(S) * BigInt(bp)) / 10000n);
    const g: number[] = stages.map((s: any) => (s.fixed >= 0 ? s.fixed : s.bp > 0 ? floorBp(s.bp) : 0));
    const stated = g.reduce((a, b) => a + b, 0);
    if (stated > S) throw new Error("contract.stages come to " + fmtMoney(stated - S) + " more than the contract sum");

    if (balanceAt >= 0) {
      g[balanceAt] = S - stated;
      if (g[balanceAt] <= 0) throw new Error("contract.stages: the earlier stages already come to the whole contract sum, so there is no balance");
    } else {
      const exact = BigInt(fixedTotal) * 10000n + BigInt(S) * BigInt(bpTotal);
      const whole = BigInt(S) * 10000n;
      if (exact < whole) {
        throw new Error(
          "contract.stages come to " + fmtMoney(S - fixedTotal - floorBp(bpTotal)) +
            " short of the contract sum; make the last stage `balance: true`",
        );
      }
      if (exact > whole) {
        // The percentages floor but the sum is checked exactly, so a stage can round to nothing and still tip it over.
        throw new Error(
          "contract.stages come to " + fmtMoney(Number((exact - whole + 9999n) / 10000n)) +
            " more than the contract sum (a percentage stage rounds to less than one penny)",
        );
      }
      // Lost pennies go on the last percentage stage; a fixed stage is paid as stated.
      let last = -1;
      stages.forEach((s: any, i: number) => {
        if (s.bp > 0) last = i;
      });
      if (last >= 0) g[last] += S - stated;
    }

    // Deposit netting.
    const d: number[] = g.map(() => 0);
    if (D > 0) {
      if (application === "first" || application === "final") {
        let rem = D;
        const order = g.map((_, i) => i);
        if (application === "final") order.reverse();
        for (const i of order) {
          d[i] = Math.min(rem, g[i]);
          rem -= d[i];
        }
      } else {
        let sum = 0;
        for (let i = 0; i < n - 1; i++) {
          d[i] = Number((BigInt(D) * BigInt(g[i])) / BigInt(S));
          sum += d[i];
        }
        d[n - 1] = D - sum;
        let excess = d[n - 1] - g[n - 1];
        if (excess > 0) {
          d[n - 1] = g[n - 1];
          for (let i = n - 2; i >= 0 && excess > 0; i--) {
            const take = Math.min(excess, g[i] - d[i]);
            d[i] += take;
            excess -= take;
          }
        }
      }
    }
    const net = g.map((x, i) => x - d[i]);

    const empty = () => ({
      description: "", status: "", amount: "", vat: "", taxPoint: "",
      grossMinor: 0, deductionMinor: 0, netMinor: 0, vatMinor: null as number | null, taxPointDate: null as string | null,
    });
    const rows = stages.map((s: any, i: number) => {
      let taxPointDate: string | null = null;
      if (taxPoints) {
        if (s.state === "invoiced") taxPointDate = s.invoiced;
        else if (s.state === "paid") taxPointDate = s.invoiced && s.invoiced < s.paid ? s.invoiced : s.paid;
      }
      const vatMinor = taxPointDate ? vatOf(net[i], rateBp) : null;
      const status = s.state === "pending"
        ? "Not yet reached"
        : s.state === "reached"
        ? (s.reached ? "Reached " + fmtDate(s.reached) : "Reached")
        : s.state === "invoiced"
        ? "Invoiced " + fmtDate(s.invoiced)
        : "Paid " + fmtDate(s.paid);
      return {
        description: d[i] > 0 ? s.label + " (" + fmtMoney(g[i]) + " less " + fmtMoney(d[i]) + " of the deposit)" : s.label,
        status,
        amount: fmtMoney(net[i]),
        vat: !taxPoints ? "" : vatMinor === null ? "Not yet due" : fmtMoney(vatMinor),
        taxPoint: !taxPoints ? "" : taxPointDate ? fmtDate(taxPointDate) : "Not yet",
        grossMinor: g[i], deductionMinor: d[i], netMinor: net[i], vatMinor, taxPointDate,
      };
    });

    let deposit = empty();
    let depositVat = 0;
    if (D > 0) {
      if (taxPoints) depositVat = vatOf(D, rateBp);
      deposit = {
        description: "Deposit",
        status: "Received " + fmtDate(receivedDate),
        amount: fmtMoney(D),
        vat: taxPoints ? fmtMoney(depositVat) : "",
        taxPoint: taxPoints ? fmtDate(receivedDate) : "",
        grossMinor: D, deductionMinor: 0, netMinor: D,
        vatMinor: taxPoints ? depositVat : null,
        taxPointDate: taxPoints ? receivedDate : null,
      };
    }

    // Totals.
    let paidMinor = D;
    let paidVatMinor = depositVat;
    let invoicedMinor = 0;
    let invoicedVatMinor = 0;
    let notInvoicedMinor = 0;
    stages.forEach((s: any, i: number) => {
      if (s.state === "paid") {
        paidMinor += net[i];
        paidVatMinor += rows[i].vatMinor ?? 0;
      } else if (s.state === "invoiced") {
        invoicedMinor += net[i];
        invoicedVatMinor += rows[i].vatMinor ?? 0;
      } else notInvoicedMinor += net[i];
    });
    if (paidMinor + invoicedMinor + notInvoicedMinor !== S) {
      throw new Error("stagePaymentSchedule: the totals do not add up to the contract sum");
    }

    const after = D > 0 ? ", after a deposit of " + fmtMoney(D) : "";
    const intro = single
      ? "The contract sum of " + fmtMoney(S) + " is paid in a single sum on completion" + after + "."
      : "The contract sum of " + fmtMoney(S) + " is paid in " + (n === 1 ? "one stage" : n + " stages") + after + ".";

    let taxNote = "";
    if (taxPoints) {
      taxNote =
        "VAT on each stage is due at its tax point: the date it is invoiced or paid, whichever comes first. " +
        "Reaching a stage is not a tax point, so a stage that is reached but not yet invoiced has no VAT due." +
        (D > 0
          ? " VAT on the deposit was due when it was received, so each stage's VAT is worked on what is left of it after the deposit."
          : "");
    } else if (single && hasVat) {
      taxNote =
        "This contract is paid in a single sum, so it is not a stage-payment contract under VAT regulation 93. " +
        "Its tax point is the date the work is completed, unless it is invoiced or paid for before then. VAT of " +
        fmtMoney(vatOf(net[0], rateBp)) + " is due at that point" +
        (D > 0
          ? " on the " + fmtMoney(net[0]) + " left after the deposit, which had its own tax point when it was received on " +
            fmtDate(receivedDate) + "."
          : ".");
    }

    const slots = rows.slice();
    while (slots.length < SLOTS) slots.push(empty());

    return {
      intro,
      taxPoints,
      taxNote,
      deposit,
      stages: slots,
      totals: {
        paid: fmtMoney(paidMinor),
        paidVat: taxPoints ? fmtMoney(paidVatMinor) : "",
        invoiced: fmtMoney(invoicedMinor),
        invoicedVat: taxPoints ? fmtMoney(invoicedVatMinor) : "",
        notInvoiced: fmtMoney(notInvoicedMinor),
        notInvoicedVat: taxPoints ? "Not yet due" : "",
        contractSum: fmtMoney(S),
        paidMinor,
        invoicedMinor,
        notInvoicedMinor,
        paidVatMinor,
        invoicedVatMinor,
      },
    };
  },
});
