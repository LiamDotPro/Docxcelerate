import { deriver } from "docxcelerate/template";

/**
 * Decides what an invoice under a fee limit may charge, and what it has to say
 * about it. Installed by `dxcl add fee-cap-status`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "feeCapStatus",
  run: (feeCap: unknown, fees: unknown) => {
    // deno-lint-ignore no-explicit-any
    const fc = feeCap as any;
    // deno-lint-ignore no-explicit-any
    const fe = fees as any;

    const num = (v: unknown): v is number => typeof v === "number" && Number.isFinite(v);

    const kinds = ["hard", "estimate", "ceiling"];
    if (!kinds.includes(fc?.kind)) {
      throw new Error("feeCap.kind must be one of " + kinds.join(", "));
    }
    const kind: string = fc.kind;

    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency: fc.currency ?? "GBP", currencyDisplay: "narrowSymbol" });
    } catch {
      throw new Error("feeCap.currency must be an ISO 4217 code");
    }
    if (!num(fc.amount) || fc.amount <= 0) throw new Error("feeCap.amount must be a number above 0");
    if (!num(fc.billedToDate) || fc.billedToDate < 0) throw new Error("feeCap.billedToDate must be a number, 0 or more");
    if (!num(fe?.thisPeriod) || fe.thisPeriod < 0) throw new Error("fees.thisPeriod must be a number, 0 or more");
    const nt = fc.noticeThreshold;
    if (nt !== undefined && nt !== null && (!num(nt) || nt <= 0 || nt >= 100)) {
      throw new Error("feeCap.noticeThreshold must be a number above 0 and below 100");
    }
    const isDate = (v: unknown) => typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v);
    const rev = kind === "estimate" && fc.revisedEstimate ? fc.revisedEstimate : null;
    if (rev) {
      if (!num(rev.amount) || rev.amount <= 0) throw new Error("feeCap.revisedEstimate.amount must be a number above 0");
      if (!isDate(rev.notifiedOn)) throw new Error("feeCap.revisedEstimate.notifiedOn must be an ISO 8601 date");
    }
    const wh = fc.withholding ? fc.withholding : null;
    if (wh) {
      if (!num(wh.percent) || wh.percent <= 0 || wh.percent > 100) {
        throw new Error("feeCap.withholding.percent must be a number above 0 and at most 100");
      }
      if (!num(wh.limit) || wh.limit <= 0) throw new Error("feeCap.withholding.limit must be a number above 0");
      if (!num(wh.withheldToDate) || wh.withheldToDate < 0) {
        throw new Error("feeCap.withholding.withheldToDate must be a number, 0 or more");
      }
    }

    // Integer minor units, so 0.1 + 0.2 against a cap of 0.3 is "reached".
    const exp = money.resolvedOptions().maximumFractionDigits ?? 2;
    const minor = (v: number) => Math.round(v * 10 ** exp);
    const fmt = (m: number) => money.format(m / 10 ** exp);
    const major = (m: number) => m / 10 ** exp;
    const pctFmt = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 2 });
    // UTC, so the date does not shift by a day on the engine.
    const dateFmt = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const day = (iso: string) => dateFmt.format(new Date(iso.slice(0, 10) + "T00:00:00Z"));

    const estimate = kind === "estimate";
    const L = minor(rev ? rev.amount : fc.amount);
    const B = minor(fc.billedToDate);
    const P = minor(fe.thisPeriod);
    const T: number | null = nt === undefined || nt === null ? (kind === "ceiling" ? 85 : null) : nt;

    const phrase = kind === "hard"
      ? "the agreed cap"
      : kind === "ceiling"
      ? "the agreed ceiling"
      : rev
      ? "our revised estimate"
      : "our estimate";
    const Phrase = phrase.charAt(0).toUpperCase() + phrase.slice(1);
    const noun = kind === "hard" ? "cap" : kind === "ceiling" ? "ceiling" : rev ? "revised estimate" : "estimate";
    const heading = kind === "hard" ? "Fee cap" : kind === "ceiling" ? "Fee ceiling" : "Fee estimate";

    const charged = estimate ? P : Math.min(P, Math.max(0, L - B));
    const writtenOff = P - charged;
    const used = B + charged;
    const upFront = B >= L;

    let state: "under" | "approaching" | "reached" | "exceeded";
    if (estimate ? used > L : writtenOff > 0) state = "exceeded";
    else if (estimate ? used === L : used >= L) state = "reached";
    else if (T !== null && used * 100 >= T * L) state = "approaching";
    else state = "under";

    // Floored, so a figure just under the threshold or 100% never prints as having reached it.
    const pct = Math.floor(used * 100 / L);

    let status: string;
    if (state === "under") {
      status = fmt(used) + " of " + phrase + " of " + fmt(L) + " has been billed, including this invoice.";
    } else if (state === "approaching") {
      status = fmt(used) + " of " + phrase + " of " + fmt(L) + " has been billed, including this invoice: " + pct +
        "% of it.";
    } else if (state === "reached") {
      status = upFront
        ? Phrase + " of " + fmt(L) + " was reached on an earlier invoice. Nothing further is charged."
        : !estimate
        ? "This invoice brings fees billed to " + phrase + " of " + fmt(L) + ". Work beyond it will not be charged."
        : "This invoice brings fees billed to " + phrase + " of " + fmt(L) +
          ". Work beyond it will be charged only once a revised estimate has been given.";
    } else if (estimate) {
      status = fmt(used) + " has been billed, including this invoice, against " + phrase + " of " + fmt(L) + ".";
    } else {
      status = upFront
        ? Phrase + " of " + fmt(L) + " was reached on an earlier invoice. Fees of " + fmt(writtenOff) +
          " for this period are above it and are not charged."
        : "This invoice reaches " + phrase + " of " + fmt(L) + ". Fees of " + fmt(writtenOff) +
          " above it are not charged.";
    }

    let incomplete: string | null = null;
    if (estimate && state === "exceeded") {
      const against = rev
        ? "our revised estimate of " + fmt(L) + ", notified on " + day(rev.notifiedOn) + ", and no later estimate"
        : "our estimate of " + fmt(L) + ", and no revised estimate";
      incomplete = "Fee estimate incomplete: fees billed to date, including this invoice, are " + fmt(used) +
        ", above " + against +
        " is recorded. An overrun the client was not told of can be challenged. Record the revised estimate and the date the client was told before this invoice is issued.";
    }

    const revised = rev
      ? "Our estimate of " + fmt(minor(fc.amount)) + " was revised to " + fmt(L) + ", and you were told on " +
        day(rev.notifiedOn) + "."
      : null;

    const notice = kind === "ceiling" && state === "approaching" && T !== null
      ? "This is our notice that fees billed have reached " + pctFmt.format(T) + "% of the ceiling."
      : null;

    let withheld = 0;
    let retention: string | null = null;
    let withheldLabel: string | null = null;
    if (wh) {
      const Lw = minor(wh.limit);
      const W0 = minor(wh.withheldToDate);
      const room = Math.max(0, Lw - W0);
      const byPercent = Math.round(charged * wh.percent / 100);
      withheld = Math.min(byPercent, room);
      if (room === 0) {
        retention = "Retention has reached its agreed limit of " + fmt(Lw) +
          ", so nothing is retained from this invoice.";
      } else if (withheld === 0) {
        retention = null;
      } else if (W0 + withheld === Lw) {
        retention = "We have retained " + fmt(withheld) + " of this invoice under the agreement. That brings retention to its agreed limit of " +
          fmt(Lw) + ", and nothing more will be retained.";
      } else {
        retention = "We have retained " + fmt(withheld) + " (" + pctFmt.format(wh.percent) +
          "%) of this invoice under the agreement: " + fmt(W0 + withheld) + " retained to date, against a limit of " +
          fmt(Lw) + ".";
      }
      withheldLabel = withheld === byPercent
        ? "Retained (" + pctFmt.format(wh.percent) + "%)"
        : "Retained (to the agreed limit)";
    }
    const payable = charged - withheld;
    const showRemaining = state === "under" || state === "approaching";

    return {
      heading,
      state,
      status,
      incomplete,
      revised,
      notice,
      retention,
      capLabel: kind === "hard" ? "Agreed cap" : kind === "ceiling" ? "Ceiling" : rev ? "Revised estimate" : "Estimate",
      capValue: fmt(L),
      billedToDate: fmt(B),
      thisPeriod: fmt(P),
      writtenOffLabel: writtenOff > 0 ? "Above the " + noun + ", not charged" : null,
      writtenOff: writtenOff > 0 ? fmt(writtenOff) : null,
      charged: writtenOff > 0 ? fmt(charged) : null,
      withheldLabel: withheld > 0 ? withheldLabel : null,
      withheld: withheld > 0 ? fmt(withheld) : null,
      payable: withheld > 0 ? fmt(payable) : null,
      remainingLabel: showRemaining ? "Remaining under the " + noun : null,
      remaining: showRemaining ? fmt(Math.max(0, L - used)) : null,
      amounts: {
        charged: major(charged),
        writtenOff: major(writtenOff),
        withheld: major(withheld),
        payable: major(payable),
      },
    };
  },
});
