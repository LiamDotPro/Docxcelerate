// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Works out a contingent or damages-based fee statement: every deduction in
 * order, as signed amounts that sum to the net. Installed by
 * `dxcl add contingency-settlement-statement`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper and statutory table is declared inside `run`.
 */
export default deriver({
  name: "contingencySettlementStatement",
  run: (settlement: unknown) => {
    const st = settlement as any;
    const SLOTS = 16;
    const STAGES = ["pre-suit", "filed", "trial", "appeal"];

    // Currency first: the exponent decides how every amount is read and printed.
    const currency = st?.currency;
    let money: Intl.NumberFormat;
    const locale = currency === "USD" ? "en-US" : "en-GB";
    try {
      // Intl accepts any three letters, so a code it does not know is checked separately.
      if (typeof currency !== "string" || !(Intl as any).supportedValuesOf("currency").includes(currency)) {
        throw new Error("unknown");
      }
      money = new Intl.NumberFormat(locale, { style: "currency", currency });
    } catch {
      throw new Error("settlement.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const unit = 10 ** exponent;
    // A zero is never printed negative: Intl would write -$0.00.
    const fmt = (minor: number) => money.format(minor === 0 ? 0 : minor / unit);
    const pctFmt = new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 });
    const pct = (p: number) => pctFmt.format(p) + "%";
    // Every computed figure is rounded once, half up, on a non-negative value.
    const round = (x: number) => Math.round(x);

    const amount = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
        throw new Error(field + " must be a finite amount of 0 or more");
      }
      return Math.round(value * unit);
    };
    const optAmount = (value: unknown, field: string) => value === undefined ? 0 : amount(value, field);

    // ---- 1. Validate ----
    if (st === null || typeof st !== "object") throw new Error("settlement is required");
    const regime = st.regime;
    if (regime !== "contingency" && regime !== "DBA" && regime !== "CFA") {
      throw new Error('settlement.regime must be "contingency", "DBA" or "CFA"');
    }
    const uk = regime !== "contingency";
    const gross = amount(st.grossRecovery, "settlement.grossRecovery");
    const stage = st.stage;
    if (stage !== undefined && !STAGES.includes(stage)) {
      throw new Error("settlement.stage must be one of " + STAGES.join(", "));
    }
    if (!Array.isArray(st.schedule) || st.schedule.length === 0) {
      throw new Error("settlement.schedule must have at least one row");
    }
    st.schedule.forEach((r: any, i: number) => {
      const f = "settlement.schedule[" + i + "]";
      if (r?.stage !== undefined && !STAGES.includes(r.stage)) throw new Error(f + ".stage is not a known stage");
      if (typeof r?.percent !== "number" || !Number.isFinite(r.percent) || r.percent < 0 || r.percent > 100) {
        throw new Error(f + ".percent must be between 0 and 100");
      }
      if (r.upTo !== undefined) {
        if (regime === "CFA") throw new Error(f + ".upTo is not allowed for a CFA");
        if (typeof r.upTo !== "number" || !Number.isFinite(r.upTo) || r.upTo <= 0) {
          throw new Error(f + ".upTo must be a positive amount");
        }
      }
    });
    const base = st.base;
    if (!uk || regime === "DBA") {
      if (base !== "gross" && base !== "net") throw new Error('settlement.base must be "gross" or "net"');
    }
    const statutoryCap = st.statutoryCap;
    if (statutoryCap !== undefined) {
      if (regime !== "contingency") throw new Error("settlement.statutoryCap applies to a contingency fee only");
      if (statutoryCap !== "ny-medical-malpractice" && statutoryCap !== "ca-medical-negligence") {
        throw new Error("settlement.statutoryCap is not a known cap");
      }
      if (statutoryCap === "ca-medical-negligence" && stage === undefined) {
        throw new Error("settlement.stage is required for the California cap");
      }
    }
    const claimType = st.claimType;
    let capBase = 0;
    let capPct = 0;
    let future = 0;
    let personalInjury = false;
    let general = 0;
    let past = 0;
    if (uk) {
      if (claimType !== "personal-injury" && claimType !== "employment" && claimType !== "other") {
        throw new Error('settlement.claimType must be "personal-injury", "employment" or "other"');
      }
      personalInjury = claimType === "personal-injury";
      if (personalInjury) {
        const d = st.damagesBreakdown;
        if (d === null || typeof d !== "object") throw new Error("settlement.damagesBreakdown is required for a personal injury claim");
        general = amount(d.general, "settlement.damagesBreakdown.general");
        past = amount(d.pastLoss, "settlement.damagesBreakdown.pastLoss");
        future = optAmount(d.futureLoss, "settlement.damagesBreakdown.futureLoss");
        if (general + past > gross) {
          throw new Error("settlement.damagesBreakdown: general plus pastLoss is above settlement.grossRecovery");
        }
      }
    }
    const baseCosts = regime === "CFA" ? amount(st.baseCosts, "settlement.baseCosts") : 0;
    let vatPercent = 0;
    if (uk && st.vatPercent !== undefined) {
      if (typeof st.vatPercent !== "number" || !Number.isFinite(st.vatPercent) || st.vatPercent < 0) {
        throw new Error("settlement.vatPercent must be a finite percentage of 0 or more");
      }
      vatPercent = st.vatPercent;
    }
    const counsel = uk ? optAmount(st.counselFees, "settlement.counselFees") : 0;
    const recoverable = uk ? optAmount(st.recoverableCostsPaid, "settlement.recoverableCostsPaid") : 0;

    const expenses: Array<{ label: string; minor: number }> = [];
    if (st.expenses !== undefined && !Array.isArray(st.expenses)) throw new Error("settlement.expenses must be a list");
    (st.expenses ?? []).forEach((e: any, i: number) => {
      const f = "settlement.expenses[" + i + "]";
      if (typeof e?.label !== "string" || e.label.trim() === "") throw new Error(f + ".label is required");
      expenses.push({ label: e.label, minor: amount(e.amount, f + ".amount") });
    });
    const liens: Array<{ holder: string; amount: number; payable: number; reduced: boolean; before: boolean; cru: boolean }> = [];
    if (st.liens !== undefined && !Array.isArray(st.liens)) throw new Error("settlement.liens must be a list");
    (st.liens ?? []).forEach((l: any, i: number) => {
      const f = "settlement.liens[" + i + "]";
      if (typeof l?.holder !== "string" || l.holder.trim() === "") throw new Error(f + ".holder is required");
      const a = amount(l.amount, f + ".amount");
      const reduced = l.reducedAmount !== undefined;
      const payable = reduced ? amount(l.reducedAmount, f + ".reducedAmount") : a;
      if (payable > a) throw new Error(f + ".reducedAmount must not be greater than amount");
      liens.push({ holder: l.holder, amount: a, payable, reduced, before: l.beforeFee === true, cru: l.cru === true });
    });

    // ---- 2. Totals ----
    const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
    const E = sum(expenses.map((e) => e.minor));
    const Lb = sum(liens.filter((l) => l.before).map((l) => l.payable));
    const cruTotal = sum(liens.filter((l) => l.cru).map((l) => l.payable));

    // ---- 5. Schedule selection ----
    if (stage === undefined && st.schedule.some((r: any) => r.stage !== undefined)) {
      throw new Error("settlement.stage is required when settlement.schedule names a stage");
    }
    let rows5 = st.schedule.filter((r: any) => stage !== undefined && r.stage === stage);
    if (rows5.length === 0) rows5 = st.schedule.filter((r: any) => r.stage === undefined);
    if (rows5.length === 0) throw new Error('no rate in settlement.schedule for stage "' + stage + '"');
    const stageWords: Record<string, string> = {
      "pre-suit": "before suit was filed",
      filed: "after filing",
      trial: "at trial",
      appeal: "on appeal",
    };
    const stagePhrase = stage === undefined ? "" : stageWords[stage];

    let uplift = 0;
    let bandRows: Array<{ top: number; percent: number }> = [];
    if (regime === "CFA") {
      if (rows5.length !== 1) throw new Error("settlement.schedule must give one uplift for a CFA");
      uplift = rows5[0].percent;
    } else {
      if (rows5.filter((r: any) => r.upTo === undefined).length !== 1) {
        throw new Error("settlement.schedule must have exactly one open band (no upTo)");
      }
      bandRows = rows5.map((r: any) => ({
        top: r.upTo === undefined ? Infinity : Math.round(r.upTo * unit),
        percent: r.percent,
      }));
      bandRows.sort((a, b) => a.top - b.top);
      for (let i = 1; i < bandRows.length; i++) {
        if (bandRows[i].top === bandRows[i - 1].top) throw new Error("settlement.schedule has two bands with the same upTo");
      }
    }

    // ---- 6. Band charge ----
    const bandsOf = (x: number, table: Array<{ top: number; percent: number }>) => {
      let prev = 0;
      let total = 0;
      const reached: Array<{ prev: number; top: number; portion: number; percent: number; charge: number }> = [];
      for (const band of table) {
        const portion = Math.min(Math.max(x, prev), band.top) - prev;
        const charge = round(portion * band.percent / 100);
        total += charge;
        reached.push({ prev, top: band.top, portion, percent: band.percent, charge });
        if (x <= band.top) break;
        prev = band.top;
      }
      return { total, reached };
    };

    // ---- 3. Fee base ----
    const baseOf = (b: string) => Math.max(0, gross - (b === "net" ? E : 0) - Lb);
    const B = regime === "CFA" ? 0 : baseOf(base);

    // ---- 7. Regime ----
    const out: Array<{ label: string; minor: number }> = [];
    const push = (label: string, minor: number) => out.push({ label, minor });
    const lienLabel = (l: (typeof liens)[number]) => l.reduced ? " (reduced from " + fmt(l.amount) + ")" : "";
    if (regime !== "CFA" && base === "net") expenses.forEach((e) => push(e.label, -e.minor));
    liens.filter((l) => l.before).forEach((l) => push(l.holder + ", paid before the fee" + lienLabel(l), -l.payable));

    let capText = "";
    let feeMinor = 0;
    let otherBase = "";
    const NY = [
      { top: 250000, percent: 30 },
      { top: 500000, percent: 25 },
      { top: 1000000, percent: 20 },
      { top: 1250000, percent: 15 },
      { top: Infinity, percent: 10 },
    ].map((b) => ({ top: b.top * unit, percent: b.percent }));
    const S = Math.max(0, gross - E);
    let statutory: number | undefined;
    if (statutoryCap === "ny-medical-malpractice") statutory = bandsOf(S, NY).total;
    if (statutoryCap === "ca-medical-negligence") statutory = round(S * (stage === "pre-suit" ? 25 : 33) / 100);

    const bandInfo = bandsOf(B, bandRows);
    let vat = 0;
    // The charge of a regime on a given base, so `otherBase` can ask the same question again.
    let dbaCap = 0;
    const dbaCharge = (agreed: number) => {
      const v = round(agreed * vatPercent / 100);
      const payment = agreed + v + counsel;
      return { vat: v, payment, charged: Math.min(payment, dbaCap) };
    };
    let actual = 0;

    if (regime === "contingency") {
      const agreed = bandInfo.total;
      const fee = Math.min(agreed, statutory ?? agreed);
      actual = fee;
      feeMinor = fee;
      push(statutory !== undefined && statutory < agreed ? "Contingent fee (statutory maximum)" : "Contingent fee", -fee);
      if (statutoryCap === "ny-medical-malpractice") {
        capText = "New York Judiciary Law § 474-a limits the fee in a medical malpractice action to " + fmt(statutory!) +
          ", worked out on the recovery after expenses (" + fmt(S) +
          "): 30% of the first $250,000, 25% of the next $250,000, 20% of the next $500,000, 15% of the next $250,000 and 10% above $1,250,000.";
      } else if (statutoryCap === "ca-medical-negligence") {
        capText =
          "California Business and Professions Code § 6146 limits the fee in a medical negligence action to 25% of the recovery after expenses if the case settles before a complaint or demand for arbitration is filed, and 33% after. This case " +
          (stage === "pre-suit" ? "settled before filing" : "was filed") + ", so the limit is " + fmt(statutory!) + ".";
      }
      if (capText) {
        capText += agreed <= statutory!
          ? " The agreement's fee of " + fmt(agreed) + " is within that limit, so it is charged."
          : " The agreement's fee of " + fmt(agreed) + " is above it, so the statutory maximum is charged.";
      }
    } else if (regime === "DBA") {
      if (claimType === "personal-injury") {
        capBase = general + Math.max(0, past - cruTotal);
        capPct = 25;
      } else {
        capBase = gross;
        capPct = claimType === "employment" ? 35 : 50;
      }
      dbaCap = round(capBase * capPct / 100);
      const agreed = bandInfo.total;
      const d = dbaCharge(agreed);
      vat = d.vat;
      const credit = Math.min(recoverable, d.charged);
      actual = d.charged;
      feeMinor = d.charged - credit;
      push("Payment under the agreement", -agreed);
      if (vat > 0) push("VAT at " + pct(vatPercent) + " on the payment", -vat);
      if (counsel > 0) push("Counsel's fees", -counsel);
      if (d.payment - d.charged > 0) push("Reduced to the statutory maximum", d.payment - d.charged);
      if (credit > 0) push("Costs paid by the other side, credited", credit);
      const what = claimType === "personal-injury"
        ? "25% of general damages and past losses net of sums recoverable by the CRU (" + fmt(capBase) + ")" +
          (future > 0 ? "; future loss of " + fmt(future) + " is excluded" : "")
        : pct(capPct) + " of the sums recovered (" + fmt(capBase) + ")";
      capText = "The Damages-Based Agreements Regulations 2013 limit the payment, including VAT and counsel's fees, to " +
        what + ", which is " + fmt(dbaCap) + "." +
        (d.payment <= dbaCap
          ? " The payment of " + fmt(d.payment) + " is within it."
          : " The payment of " + fmt(d.payment) + " is above it, so it is reduced to " + fmt(dbaCap) + ".");
    } else {
      const sfFee = round(baseCosts * uplift / 100);
      let cfaCap: number | undefined;
      if (claimType === "personal-injury") {
        capBase = general + Math.max(0, past - cruTotal);
        // Before VAT: CFA Order 2013 art. 5 does not say "including VAT", unlike DBA reg 4(2).
        cfaCap = round(capBase * 25 / 100);
      }
      const sfCharged = Math.min(sfFee, cfaCap ?? sfFee);
      const baseVat = round(baseCosts * vatPercent / 100);
      vat = baseVat + round(sfCharged * vatPercent / 100);
      // The success fee is not recoverable from the opponent, so costs never offset it.
      const credit = Math.min(recoverable, baseCosts + baseVat + counsel);
      feeMinor = baseCosts + sfCharged + vat + counsel - credit;
      push("Base costs", -baseCosts);
      push("Success fee (" + pct(uplift) + " of base costs)", -sfFee);
      if (sfFee - sfCharged > 0) push("Success fee reduced to the statutory maximum", sfFee - sfCharged);
      if (vat > 0) push("VAT at " + pct(vatPercent), -vat);
      if (counsel > 0) push("Counsel's fees", -counsel);
      if (credit > 0) push("Costs paid by the other side, credited", credit);
      if (cfaCap !== undefined) {
        capText = "The Conditional Fee Agreements Order 2013 limits the success fee to 25% of general damages and past losses net of sums recoverable by the CRU (" +
          fmt(capBase) + "), which is " + fmt(cfaCap) + "." +
          (sfFee <= cfaCap
            ? " The success fee of " + fmt(sfFee) + " is within it."
            : " The success fee of " + fmt(sfFee) + " is above it, so it is reduced to " + fmt(cfaCap) + ".");
      }
    }

    // ---- 8. After the fee ----
    if (regime === "CFA" || base === "gross") expenses.forEach((e) => push(e.label, -e.minor));
    liens.filter((l) => !l.before).forEach((l) => push(l.holder + lienLabel(l), -l.payable));

    // ---- 9. Slots ----
    if (out.length > SLOTS) {
      throw new Error("contingencySettlementStatement: " + out.length + " lines; this component prints at most " + SLOTS);
    }
    const rows = out.map((r) => ({ label: r.label, amount: fmt(r.minor) }));
    while (rows.length < SLOTS) rows.push({ label: "", amount: "" });

    // ---- 10. Net ----
    const netMinor = gross + sum(out.map((r) => r.minor));
    const shortfall = netMinor < 0;

    // ---- 11. Method ----
    const baseName = (b: string) => (b === "net" ? "the recovery after expenses" : "the gross recovery") +
      (Lb > 0 ? " and the liens paid before the fee" : "");
    let method: string;
    if (regime === "CFA") {
      method = "The success fee is " + pct(uplift) + " of base costs of " + fmt(baseCosts) +
        (stagePhrase ? ", the rate the agreement sets " + stagePhrase : "") + ".";
    } else if (B === 0) {
      method = "There is nothing to work the fee out on, so no fee is charged.";
    } else {
      const noun = regime === "DBA" ? "payment" : "fee";
      const tail = regime === "DBA"
        ? (vat > 0 && counsel > 0 ? ", plus VAT and counsel's fees" : vat > 0 ? ", plus VAT" : counsel > 0 ? ", plus counsel's fees" : "")
        : "";
      const reached = bandInfo.reached;
      const on = baseName(base) + " (" + fmt(B) + ")";
      if (reached.length === 1) {
        method = "The " + noun + " is " + pct(reached[0].percent) + " of " + on +
          (stagePhrase ? ", the rate the agreement sets " + stagePhrase : "") + tail + ".";
      } else {
        const parts = reached.map((b, i) => {
          const last = i === reached.length - 1;
          const where = i === 0
            ? "the first " + fmt(b.portion)
            : !last
            ? "the next " + fmt(b.portion)
            : b.top === Infinity
            ? "the amount above " + fmt(b.prev)
            : "the remaining " + fmt(b.portion);
          return pct(b.percent) + " of " + where + " (" + fmt(b.charge) + ")";
        });
        method = "The " + noun + " is worked out band by band on " + on +
          (stagePhrase ? ", at the rates the agreement sets " + stagePhrase : "") + ": " +
          parts.slice(0, -1).join(", ") + " and " + parts[parts.length - 1] + tail + ".";
      }
    }

    // ---- 12. Other base ----
    if (regime !== "CFA" && E > 0) {
      const other = base === "net" ? "gross" : "net";
      const agreed = bandsOf(baseOf(other), bandRows).total;
      const alt = regime === "DBA" ? dbaCharge(agreed).charged : Math.min(agreed, statutory ?? agreed);
      if (alt !== actual) {
        otherBase = "Worked out on " + (other === "gross" ? "the gross recovery" : "the recovery after expenses") +
          " instead, the " + (regime === "DBA" ? "payment" : "fee") + " would have been " + fmt(alt) + ".";
      }
    }

    return {
      gross: fmt(gross),
      rows,
      netLabel: shortfall ? "Shortfall" : "Net to you",
      net: fmt(Math.abs(netMinor)),
      outcome: shortfall ? "shortfall" : "remittance",
      method,
      otherBase,
      cap: capText,
      feeMinor,
      netMinor,
    };
  },
});
