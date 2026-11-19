// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Takes each claim line from billed charge to allowed amount to what the
 * patient owes, carrying the deductible and out-of-pocket balances from line to
 * line. Installed by `dxcl add patient-responsibility-breakdown`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "patientResponsibilityBreakdown",
  run: (patientResponsibility: unknown) => {
    const pr = patientResponsibility as any;

    // Currency first: the exponent decides how every amount is read and printed.
    const currency = pr?.currency ?? "USD";
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-US", { style: "currency", currency });
    } catch {
      throw new Error("patientResponsibility.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmt = (minor: number) => money.format(minor / 10 ** exponent);
    const dates = new Intl.DateTimeFormat("en-US", { dateStyle: "medium", timeZone: "UTC" });

    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
        throw new Error(field + " must be a non-negative integer of minor units");
      }
      return value;
    };

    const plan = pr?.plan;
    let ded = minorOf(plan?.deductibleRemainingMinor, "patientResponsibility.plan.deductibleRemainingMinor");
    let oop = minorOf(plan?.oopRemainingMinor, "patientResponsibility.plan.oopRemainingMinor");
    const rate = plan?.coinsuranceRate;
    if (typeof rate !== "number" || !Number.isFinite(rate) || rate < 0 || rate > 1) {
      throw new Error("patientResponsibility.plan.coinsuranceRate must be a number from 0 to 1");
    }
    // Basis points keep the rate in integers.
    const bp = Math.round(rate * 10000);
    const copayFirst = plan.copayAppliesBeforeDeductible ?? true;
    const counts = plan.copayCountsTowardDeductible ?? false;
    // Meeting the out-of-pocket maximum satisfies the deductible too.
    if (oop === 0) ded = 0;

    const raw: unknown[] = pr.lines === undefined || pr.lines === null ? [] : pr.lines;
    if (!Array.isArray(raw)) throw new Error("patientResponsibility.lines must be an array");

    // Validate everything before any arithmetic, so a bad line never half-prints.
    const checked = raw.map((l: any, i) => {
      const f = "patientResponsibility.lines[" + i + "]";
      if (typeof l?.serviceDate !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(l.serviceDate)) {
        throw new Error(f + ".serviceDate must be an ISO 8601 date");
      }
      if (typeof l.description !== "string" || l.description.trim() === "") {
        throw new Error(f + ".description is required");
      }
      const billed = minorOf(l.billedMinor, f + ".billedMinor");
      const covered = l.covered !== false;
      let allowed = 0;
      if (covered) {
        allowed = minorOf(l.allowedMinor, f + ".allowedMinor");
        if (allowed > billed) throw new Error(f + ".allowedMinor must not be above billedMinor");
      }
      const copay = l.copayMinor === undefined ? 0 : minorOf(l.copayMinor, f + ".copayMinor");
      if (l.network !== undefined && l.network !== "in" && l.network !== "out") {
        throw new Error(f + '.network must be "in" or "out"');
      }
      return { l, i, f, day: l.serviceDate.slice(0, 10), billed, covered, allowed, copay };
    });
    // Stable by construction: ties fall back to input order.
    checked.sort((a, b) => (a.day < b.day ? -1 : a.day > b.day ? 1 : a.i - b.i));

    type Row = {
      service: string;
      detail: string;
      billed: string;
      allowed: string;
      deductible: string;
      copay: string;
      coinsurance: string;
      youOwe: string;
      note: string | null;
    };
    const rows: Row[] = [];
    const sum = { billed: 0, allowed: 0, deductible: 0, copay: 0, coinsurance: 0, youOwe: 0 };
    let adjustmentTotal = 0;
    let excludedTotal = 0;

    for (const c of checked) {
      let deductible = 0;
      let copay = 0;
      let coinsurance = 0;
      let notCovered = 0;
      let balance = 0;
      let adjustment = 0;
      let note: string | null = null;

      if (!c.covered) {
        notCovered = c.billed;
        note = "Not covered by your plan. The full charge of " + fmt(c.billed) +
          " is yours to pay, and it does not count toward your deductible or out-of-pocket maximum.";
      } else if (c.allowed === 0 && c.billed > 0) {
        note = "Your plan allowed nothing for this service, so nothing is billed to you for it on this statement.";
      } else if (c.allowed > 0) {
        const out = c.l.network === "out";
        const unprotected = out && c.l.nsaProtected === false;
        if (unprotected) balance = c.billed - c.allowed;
        else adjustment = c.billed - c.allowed;

        const startOop = oop;
        // Every cost-share step passes through the cap, in plan order.
        const take = (x: number) => {
          const t = Math.min(x, oop);
          oop -= t;
          return t;
        };
        let left = c.allowed;
        const copayStep = () => {
          copay = take(Math.min(c.copay, left));
          left -= copay;
          if (counts) ded -= Math.min(copay, ded);
        };
        const deductibleStep = () => {
          deductible = take(Math.min(left, ded));
          left -= deductible;
          ded -= deductible;
        };
        if (c.l.copayOnly === true) copayStep();
        else {
          if (copayFirst) {
            copayStep();
            deductibleStep();
          } else {
            deductibleStep();
            copayStep();
          }
          if (!Number.isSafeInteger(left * bp)) throw new Error(c.f + ".allowedMinor is too large to apply the rate exactly");
          // Half up to the cent, on what is left after the deductible and copay.
          coinsurance = take(Math.floor((left * bp + 5000) / 10000));
        }
        if (oop === 0) ded = 0;

        if (balance > 0) {
          note = "Out of network. The provider bills you the " + fmt(balance) +
            " above your plan's allowed amount, which does not count toward your deductible or out-of-pocket maximum.";
        } else if (c.billed === 0) {
          note = null;
        } else if (startOop > 0 && oop === 0) {
          note = "You reached your out-of-pocket maximum with this service. Your plan pays the rest.";
        } else if (startOop === 0) {
          note = "Your plan pays 100%: you have already reached your out-of-pocket maximum.";
        } else if (out) {
          note = "Out of network, but protected by the No Surprises Act: you pay in-network cost sharing and cannot be billed the difference.";
        }
      }

      const youOwe = deductible + copay + coinsurance + notCovered + balance;
      sum.billed += c.billed;
      sum.allowed += c.covered ? c.allowed : 0;
      sum.deductible += deductible;
      sum.copay += copay;
      sum.coinsurance += coinsurance;
      sum.youOwe += youOwe;
      adjustmentTotal += adjustment;
      excludedTotal += notCovered + balance;

      const code = typeof c.l.code === "string" ? c.l.code.trim() : "";
      const date = dates.format(new Date(c.day + "T00:00:00Z"));
      rows.push({
        service: c.l.description,
        detail: code ? date + " · " + code : date,
        billed: fmt(c.billed),
        allowed: fmt(c.covered ? c.allowed : 0),
        deductible: fmt(deductible),
        copay: fmt(copay),
        coinsurance: fmt(coinsurance),
        youOwe: fmt(youOwe),
        note,
      });
    }

    return {
      outcome: (rows.length === 0 ? "none" : sum.youOwe > 0 ? "owes" : "nothing-owed") as
        | "none"
        | "owes"
        | "nothing-owed",
      rows,
      total: {
        billed: fmt(sum.billed),
        allowed: fmt(sum.allowed),
        deductible: fmt(sum.deductible),
        copay: fmt(sum.copay),
        coinsurance: fmt(sum.coinsurance),
        youOwe: fmt(sum.youOwe),
      },
      youOweMinor: sum.youOwe,
      adjustment: fmt(adjustmentTotal),
      hasAdjustment: adjustmentTotal > 0,
      excluded: fmt(excludedTotal),
      hasExcluded: excludedTotal > 0,
      deductibleRemaining: fmt(ded),
      oopRemaining: fmt(oop),
      oopMet: rows.length > 0 && oop === 0,
    };
  },
});
