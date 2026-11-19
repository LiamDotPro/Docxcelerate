import { deriver } from "docxcelerate/template";

/**
 * Decides every line of a 501(r) financial assistance block: the notice, the
 * collections and application dates, the suspension, and the AGB cap. Installed
 * by `dxcl add financial-assistance-notice`.
 *
 * It is self-contained on purpose:
 * it travels to the engine as source text, so a module-level constant or helper
 * would be a ReferenceError there. Date helpers, formatters and wording are all
 * declared inside it.
 */
export default deriver({
  name: "financialAssistanceNotice",
  run: (facility: unknown, billing: unknown) => {
    // deno-lint-ignore no-explicit-any
    const fac = (facility ?? {}) as any;
    // deno-lint-ignore no-explicit-any
    const bil = (billing ?? {}) as any;

    const out: Record<
      "notice" | "noticeIncomplete" | "collections" | "apply" | "suspended" | "free" | "agbBasis" | "owed",
      string | null
    > = {
      notice: null,
      noticeIncomplete: null,
      collections: null,
      apply: null,
      suspended: null,
      free: null,
      agbBasis: null,
      owed: null,
    };

    const forProfit = fac.taxExempt === false;
    if (forProfit && fac.stateNoticeRequired !== true) return out;

    const url = typeof fac.fapUrl === "string" ? fac.fapUrl.trim() : "";
    const phone = typeof fac.fapPhone === "string" ? fac.fapPhone.trim() : "";
    if (url !== "" && phone !== "") {
      out.notice = "Financial assistance is available. If you are worried about paying this bill, " +
        "you may qualify for free or discounted care under our financial assistance policy. " +
        "A free copy of the policy, the application form and a plain-language summary of the policy are at " +
        url + ", or call " + phone + " and we will send them to you.";
    } else {
      out.noticeIncomplete = "Financial assistance notice incomplete: this statement must give the phone number " +
        "and web address where the financial assistance policy can be obtained, and at least one is missing. " +
        "Add them before this statement is sent.";
    }
    if (forProfit) return out;

    const isoDate = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      const iso = value.slice(0, 10);
      const [y, m, d] = iso.split("-").map(Number);
      if (new Date(Date.UTC(y, m - 1, d)).toISOString().slice(0, 10) !== iso) {
        throw new Error(field + " is not a real calendar date");
      }
      return iso;
    };
    // Whole UTC calendar days, so the date does not shift on the engine.
    const addDays = (iso: string, n: number): string => {
      const [y, m, d] = iso.split("-").map(Number);
      return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10);
    };
    const dateFormat = new Intl.DateTimeFormat("en-US", { dateStyle: "long", timeZone: "UTC" });
    const fmtDate = (iso: string) => dateFormat.format(new Date(iso + "T00:00:00Z"));

    const first = isoDate(bil.firstPostDischargeStatementDate, "billing.firstPostDischargeStatementDate");
    const hasEca = bil.ecaNoticeDate !== undefined && bil.ecaNoticeDate !== null && bil.ecaNoticeDate !== "";
    const ecaNotice = hasEca ? isoDate(bil.ecaNoticeDate, "billing.ecaNoticeDate") : null;

    const status = bil.application?.status ?? "none";
    if (!["none", "pending", "approved", "denied"].includes(status)) {
      throw new Error("billing.application.status must be one of none, pending, approved, denied");
    }
    const discount = bil.application?.discount;
    if (status === "approved" && discount !== "free" && discount !== "partial") {
      throw new Error("billing.application.discount must be free or partial when the application is approved");
    }

    // Day 120 after the statement is the first permitted day; the later of the
    // two clocks wins, and the application window never ends before it.
    const ecaBar = addDays(first, 120);
    const earliestEca = ecaNotice !== null && addDays(ecaNotice, 30) > ecaBar ? addDays(ecaNotice, 30) : ecaBar;
    const lastApply = addDays(first, 240) > earliestEca ? addDays(first, 240) : earliestEca;

    const collections = "We will not take extraordinary collection action on this bill, such as reporting it to a " +
      "credit agency, selling the debt or taking legal action, before " +
      fmtDate(earliestEca) +
      (ecaNotice === null ? ", and not until at least 30 days after we have written to you to say that we intend to." : ".");

    if (status === "none" || status === "denied") {
      out.collections = collections;
      out.apply = "You can apply for financial assistance until " + fmtDate(lastApply) +
        ". If you apply in that time, we will consider your application even if collection action has begun.";
    } else if (status === "pending") {
      out.suspended = "We have received your application for financial assistance. While we consider it, we have " +
        "stopped all extraordinary collection action on this bill, and we will not start any until we have told you " +
        "our decision in writing.";
    } else if (discount === "free") {
      out.free = "You qualify for free care under our financial assistance policy. You owe nothing for this care.";
    } else {
      const gross = bil.grossCharges;
      const pct = bil.agbPercent;
      const due = bil.balanceDue;
      if (typeof gross !== "number" || !Number.isFinite(gross) || gross < 0) {
        throw new Error("billing.grossCharges must be a finite number of at least 0 for a partial approval");
      }
      if (typeof pct !== "number" || !Number.isFinite(pct) || pct <= 0 || pct > 100) {
        throw new Error("billing.agbPercent must be above 0 and at most 100 for a partial approval");
      }
      const hasDue = due !== undefined && due !== null;
      if (hasDue && (typeof due !== "number" || !Number.isFinite(due) || due < 0)) {
        throw new Error("billing.balanceDue must be a finite number of at least 0");
      }

      // Integer arithmetic: a float times a float rounds 2.01 * 50% down a cent.
      const grossCents = Math.round(gross * 100);
      const bp = Math.round(pct * 100);
      const agbCents = Math.floor((grossCents * bp + 5000) / 10000);
      const money = new Intl.NumberFormat("en-US", {
        style: "currency",
        currency: typeof bil.currency === "string" && bil.currency !== "" ? bil.currency : "USD",
      });
      const fmtMoney = (cents: number) => money.format(cents / 100);
      const pctText = new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(bp / 100);
      const agb = fmtMoney(agbCents);

      out.agbBasis = "You qualify for financial assistance, so you cannot be charged more than the amounts " +
        "generally billed (AGB) to patients who have insurance for this care. Our AGB is " + pctText +
        "% of gross charges: " + pctText + "% of " + fmtMoney(grossCents) + " is " + agb + ".";
      if (!hasDue) {
        out.owed = "Amount you owe for this care: " + agb + ".";
      } else {
        const balanceCents = Math.round(due * 100);
        out.owed = balanceCents > agbCents
          ? "Amount you owe for this care: " + agb + ". This is the AGB limit, not the " + fmtMoney(balanceCents) +
            " that would otherwise have been billed."
          : "Amount you owe for this care: " + fmtMoney(balanceCents) +
            ". This is below the AGB limit, so the limit does not change it.";
      }
      out.collections = collections;
    }
    return out;
  },
});
