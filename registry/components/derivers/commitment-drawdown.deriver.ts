// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Draws prepaid commitment and credit grants down against priced usage lines,
 * line by line and before tax, and decides the tax basis, the expiry warnings
 * and the minimum-commitment true-up. Installed by `dxcl add commitment-drawdown`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "commitmentDrawdown",
  run: (drawdown: unknown) => {
    const dd = drawdown as any;

    // Currency first: the exponent decides how every amount is read and printed.
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency: dd?.currency });
    } catch {
      throw new Error("drawdown.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmtMoney = (minor: number) => money.format(minor / 10 ** exponent);
    const fmtNum = (n: number, dec = 6) => new Intl.NumberFormat("en-GB", { maximumFractionDigits: dec }).format(n);

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
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const fmt = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));

    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
        throw new Error(field + " must be a non-negative integer of minor units");
      }
      return value;
    };
    const countOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
        throw new Error(field + " must be a finite number of 0 or more");
      }
      return value;
    };
    // Half away from zero; every input is >= 0, so this is round-half-up.
    const round = (x: number, dec = 0) => {
      const f = 10 ** dec;
      return (Math.sign(x) * Math.round(Math.abs(x) * f)) / f;
    };
    const decimalsOf = (n: number) => {
      const s = String(n);
      return s.includes("e") ? 6 : Math.min(6, (s.split(".")[1] ?? "").length);
    };
    const cmp = (a: any, b: any) => (a < b ? -1 : a > b ? 1 : 0);

    const invoiceDay = isoDay(dd?.invoiceDate, "drawdown.invoiceDate");
    const startDay = isoDay(dd?.period?.start, "drawdown.period.start");
    const endDay = isoDay(dd?.period?.end, "drawdown.period.end");
    if (endDay < startDay) throw new Error("drawdown.period.end is before drawdown.period.start");
    if (dd?.creditTaxBasis !== "taxed-at-purchase" && dd?.creditTaxBasis !== "multi-purpose") {
      throw new Error('drawdown.creditTaxBasis must be "taxed-at-purchase" or "multi-purpose"');
    }
    const basis: string = dd.creditTaxBasis;
    let warnDays = 30;
    if (dd.warnDays !== undefined) {
      if (typeof dd.warnDays !== "number" || !Number.isInteger(dd.warnDays) || dd.warnDays < 0) {
        throw new Error("drawdown.warnDays must be a non-negative integer");
      }
      warnDays = dd.warnDays;
    }

    let term: { start: string; end: string; commitment: number; toDate: number; billed: number } | undefined;
    if (dd.term) {
      term = {
        start: isoDay(dd.term.start, "drawdown.term.start"),
        end: isoDay(dd.term.end, "drawdown.term.end"),
        commitment: minorOf(dd.term.commitment, "drawdown.term.commitment"),
        toDate: dd.term.eligibleSpendToDate === undefined
          ? 0
          : minorOf(dd.term.eligibleSpendToDate, "drawdown.term.eligibleSpendToDate"),
        billed: dd.term.trueUpsBilled === undefined ? 0 : minorOf(dd.term.trueUpsBilled, "drawdown.term.trueUpsBilled"),
      };
    }

    const lines = (Array.isArray(dd.lines) ? dd.lines : []).map((l: any, i: number) => {
      const f = "drawdown.lines[" + i + "]";
      if (typeof l?.description !== "string" || l.description.trim() === "") throw new Error(f + ".description is required");
      const quantity = countOf(l.quantity, f + ".quantity");
      const decimals = l.quantityDecimals === undefined
        ? decimalsOf(quantity)
        : (Number.isInteger(l.quantityDecimals) && l.quantityDecimals >= 0 && l.quantityDecimals <= 6
          ? l.quantityDecimals
          : (() => { throw new Error(f + ".quantityDecimals must be an integer from 0 to 6"); })());
      return {
        i,
        key: typeof l.key === "string" && l.key !== "" ? l.key : String(i),
        description: l.description as string,
        usageDay: l.usageDate === undefined ? endDay : isoDay(l.usageDate, f + ".usageDate"),
        quantity,
        decimals,
        unit: l.unit as string | undefined,
        priceId: l.priceId as string | undefined,
        amount: minorOf(l.amount, f + ".amount"),
        raw: l,
      };
    });

    const kindRank: Record<string, number> = { rollover: 0, commit: 1, promotional: 2 };
    const kindLabel: Record<string, string> = { rollover: "rollover", commit: "prepaid commitment", promotional: "promotional" };
    const grants = (Array.isArray(dd.grants) ? dd.grants : []).map((g: any, i: number) => {
      const f = "drawdown.grants[" + i + "]";
      if (typeof g?.id !== "string" || g.id === "") throw new Error(f + ".id is required");
      if (typeof g.name !== "string" || g.name.trim() === "") throw new Error(f + ".name is required");
      if (!(g.kind in kindRank)) throw new Error(f + '.kind must be "rollover", "commit" or "promotional"');
      if (g.denomination !== "money" && g.denomination !== "units") throw new Error(f + '.denomination must be "money" or "units"');
      const units = g.denomination === "units";
      const balance = units ? countOf(g.balance, f + ".balance") : minorOf(g.balance, f + ".balance");
      const priceIds: string[] = Array.isArray(g.priceIds) ? g.priceIds : [];
      if (units && priceIds.length === 0) throw new Error(f + ".priceIds is required for a units grant");
      if (g.priority !== undefined && (typeof g.priority !== "number" || Number.isNaN(g.priority))) {
        throw new Error(f + ".priority must be a number");
      }
      return {
        i,
        id: g.id as string,
        name: g.name as string,
        kind: g.kind as string,
        units,
        unit: (g.unit ?? "units") as string,
        start: balance,
        balance,
        priority: g.priority === undefined ? Infinity : (g.priority as number),
        effective: isoDay(g.effective, f + ".effective"),
        expires: g.expires === undefined ? undefined : isoDay(g.expires, f + ".expires"),
        created: g.created === undefined ? "￿" : isoDay(g.created, f + ".created"),
        priceIds,
        dailyBurn: g.dailyBurn === undefined ? undefined : countOf(g.dailyBurn, f + ".dailyBurn"),
        dec: units ? decimalsOf(balance) : 0,
        drawnMinor: 0,
        drawnUnits: 0,
      };
    });

    // Burn order: type, priority, soonest expiry, earliest effective, earliest created, input.
    const order = grants.filter((g: any) => g.balance > 0).sort((a: any, b: any) =>
      cmp(kindRank[a.kind], kindRank[b.kind]) || cmp(a.priority, b.priority) ||
      cmp(a.expires ?? "￿", b.expires ?? "￿") || cmp(a.effective, b.effective) ||
      cmp(a.created, b.created) || a.i - b.i
    );

    const eligible = (g: any, l: any) =>
      g.effective <= l.usageDay && (g.expires === undefined || g.expires > l.usageDay) &&
      (g.priceIds.length === 0 || (l.priceId !== undefined && g.priceIds.includes(l.priceId)));

    const results = lines.map((l: any) => {
      const draws: Array<{ grantId: string; amountMinor: number; units?: number }> = [];
      let remAmt = l.amount;
      let remQty = l.quantity;
      if (l.amount > 0) {
        for (const g of order) {
          if (remAmt === 0) break;
          if (g.balance <= 0 || !eligible(g, l)) continue;
          if (!g.units) {
            const take = Math.min(g.balance, remAmt);
            g.balance -= take;
            remAmt -= take;
            remQty = remAmt === 0 ? 0 : round((l.quantity * remAmt) / l.amount, l.decimals);
            g.drawnMinor += take;
            draws.push({ grantId: g.id, amountMinor: take });
          } else {
            if (remQty <= 0) continue;
            const unitPrice = l.raw.unitPrice;
            if (typeof unitPrice !== "number" || !Number.isFinite(unitPrice) || unitPrice < 0) {
              throw new Error("drawdown.lines[" + l.i + "].unitPrice must be a finite number of 0 or more");
            }
            const takeUnits = Math.min(g.balance, remQty);
            g.balance = round(g.balance - takeUnits, 6);
            remQty = round(remQty - takeUnits, l.decimals);
            const newAmt = remQty === 0 ? 0 : Math.min(remAmt, round(remQty * unitPrice));
            g.drawnMinor += remAmt - newAmt;
            g.drawnUnits = round(g.drawnUnits + takeUnits, 6);
            g.dec = Math.max(g.dec, l.decimals);
            draws.push({ grantId: g.id, amountMinor: remAmt - newAmt, units: takeUnits });
            remAmt = newAmt;
          }
        }
      }
      return { l, draws, charged: remAmt, covered: l.amount - remAmt, chargedQuantity: remQty };
    });

    // Totals, in integer minor units.
    let coveredPaid = 0;
    let coveredPromotional = 0;
    for (const g of order) {
      if (g.kind === "promotional") coveredPromotional += g.drawnMinor;
      else coveredPaid += g.drawnMinor;
    }
    const usage = results.reduce((s: number, r: any) => s + r.l.amount, 0);
    const covered = coveredPaid + coveredPromotional;
    const charged = usage - covered;

    // Tax basis.
    const taxArm = covered === 0 ? "none" : basis === "taxed-at-purchase" ? "purchase" : "redemption";
    const vatBase = basis === "multi-purpose" ? charged + coveredPaid : charged;
    const deductionAfterTax = basis === "multi-purpose" ? coveredPaid : 0;
    const promoSentence = coveredPromotional > 0
      ? " Promotional credit of " + fmtMoney(coveredPromotional) + " is a discount and is taken off before VAT."
      : "";
    const taxNote = taxArm === "none"
      ? ""
      : taxArm === "purchase" || coveredPaid === 0
      ? (taxArm === "purchase"
        ? "Usage covered by prepaid credit carries no VAT on this invoice, because the credit was taxed when it was bought. "
        : "") + "VAT applies to the " + fmtMoney(charged) + " of usage charged." + promoSentence
      : "VAT applies to " + fmtMoney(vatBase) +
        " of usage, because prepaid credit is a multi-purpose voucher and is taxed when it is used. The " +
        fmtMoney(coveredPaid) + " of credit applied is deducted after VAT, like a payment." + promoSentence;

    // True-up: only on the invoice that ends the term.
    const eligibleSpend = (term?.toDate ?? 0) + charged + coveredPaid;
    let trueUpArm = "none";
    let trueUpNote = "";
    let trueUp = 0;
    if (term && term.commitment > 0 && endDay === term.end) {
      const shortfall = Math.max(0, term.commitment - eligibleSpend - term.billed);
      const span = "the term " + fmt(term.start) + " to " + fmt(term.end);
      if (shortfall > 0) {
        trueUpArm = "shortfall";
        trueUp = shortfall;
        trueUpNote = "Your minimum commitment for " + span + " is " + fmtMoney(term.commitment) +
          ". Eligible spend over the term came to " + fmtMoney(eligibleSpend) +
          (term.billed > 0 ? ", and " + fmtMoney(term.billed) + " of true-ups were billed earlier" : "") +
          ", so a true-up of " + fmtMoney(shortfall) +
          " is charged on this invoice. Usage covered by promotional credit does not count towards the commitment.";
      } else {
        trueUpArm = "met";
        trueUpNote = "Your minimum commitment of " + fmtMoney(term.commitment) + " for " + span +
          " has been met: eligible spend came to " + fmtMoney(eligibleSpend) + ". No true-up is due.";
      }
    }

    // Grants shown, with a note each.
    const amt = (g: any, n: number) => (g.units ? fmtNum(n, g.dec) + " " + g.unit : fmtMoney(n));
    const warned: Array<{ name: string; expires: string }> = [];
    const shown = order.filter((g: any) =>
      g.drawnMinor > 0 || g.drawnUnits > 0 ||
      (g.effective <= invoiceDay && (g.expires === undefined || invoiceDay < g.expires))
    ).map((g: any) => {
      const remaining = g.units ? round(g.balance, g.dec) : g.balance;
      const days = g.expires === undefined ? 0 : dayIndex(g.expires) - dayIndex(invoiceDay);
      const warns = remaining > 0 && g.expires !== undefined && days > 0 && days <= warnDays;
      let note: string;
      if (warns) {
        note = "Expires in " + days + (days === 1 ? " day" : " days") + ", on " + fmt(g.expires) + ".";
        if (g.dailyBurn !== undefined) {
          const unused = remaining - round(g.dailyBurn * days, g.units ? g.dec : 0);
          if (unused > 0) note += " At recent usage about " + amt(g, round(unused, g.units ? g.dec : 0)) + " will go unused.";
        }
        warned.push({ name: g.name, expires: g.expires });
      } else if (g.expires === undefined) note = "No expiry.";
      else if (g.expires <= invoiceDay) note = "Expired on " + fmt(g.expires) + "." + (remaining > 0 ? " The rest has lapsed." : "");
      else note = "Expires on " + fmt(g.expires) + ".";
      return { g, remaining, warns, note };
    });
    const expiring = warned.length === 0
      ? ""
      : warned.length + (warned.length === 1 ? " credit expires" : " credits expire") + " within " + warnDays +
        " days of this invoice: " +
        new Intl.ListFormat("en-GB", { type: "conjunction" }).format(warned.map((w) => w.name + " on " + fmt(w.expires))) +
        ".";

    const grantRow = (s: any) => {
      const g = s.g;
      return {
        name: g.name + " (" + kindLabel[g.kind] + ")",
        applied: g.drawnMinor === 0 && g.drawnUnits === 0
          ? "—"
          : g.units
          ? fmtNum(g.drawnUnits, g.dec) + " " + g.unit + " (" + fmtMoney(g.drawnMinor) + ")"
          : fmtMoney(g.drawnMinor),
        remaining: amt(g, s.remaining),
        note: s.note,
      };
    };
    const blankGrant = () => ({ name: "", applied: "", remaining: "", note: "" });
    const grantSlots = shown.length <= 4 ? shown.map(grantRow) : shown.slice(0, 3).map(grantRow);
    if (shown.length > 4) {
      const rest = shown.slice(3);
      const unitCount = rest.filter((s: any) => s.g.units).length;
      const applied = rest.reduce((n: number, s: any) => n + s.g.drawnMinor, 0);
      grantSlots.push({
        name: rest.length + " other credits",
        applied: applied === 0 ? "—" : fmtMoney(applied),
        remaining: fmtMoney(rest.reduce((n: number, s: any) => n + (s.g.units ? 0 : s.remaining), 0)) +
          (unitCount > 0 ? " plus " + unitCount + (unitCount === 1 ? " unit credit" : " unit credits") : ""),
        note: "",
      });
    }
    while (grantSlots.length < 4) grantSlots.push(blankGrant());

    // Display lines: zero-amount lines are not printed.
    const cells = (usage: number, covered: number, charged: number) => ({
      usage: fmtMoney(usage),
      covered: covered === 0 ? "—" : fmtMoney(covered),
      charged: charged === 0 ? "—" : fmtMoney(charged),
    });
    const shownLines = results.filter((r: any) => r.l.amount > 0);
    const lineRow = (r: any) => {
      const c = cells(r.l.amount, r.covered, r.charged);
      if (r.charged > 0 && r.charged < r.l.amount) {
        c.charged += " (" + fmtNum(r.chargedQuantity, r.l.decimals) + (r.l.unit ? " " + r.l.unit : "") + ")";
      }
      return { description: r.l.description, ...c };
    };
    let lineSlots: any[];
    if (shownLines.length <= 6) lineSlots = shownLines.map(lineRow);
    else {
      const drew = shownLines.filter((r: any) => r.covered > 0).slice(0, 5);
      const folded = shownLines.filter((r: any) => !drew.includes(r));
      lineSlots = drew.map(lineRow);
      while (lineSlots.length < 5) lineSlots.push({ description: "", usage: "", covered: "", charged: "" });
      const sum = (k: string) => folded.reduce((n: number, r: any) => n + (k === "usage" ? r.l.amount : r[k]), 0);
      lineSlots.push({
        description: folded.length + (folded.some((r: any) => r.covered > 0) ? " further lines" : " lines not covered by credit"),
        ...cells(sum("usage"), sum("covered"), sum("charged")),
      });
    }
    while (lineSlots.length < 6) lineSlots.push({ description: "", usage: "", covered: "", charged: "" });

    const showLines = covered > 0;
    const showGrants = shown.length > 0;
    return {
      show: showLines || showGrants || trueUpArm !== "none",
      periodLabel: fmt(startDay) + " to " + fmt(endDay),
      showLines,
      lines: lineSlots,
      totals: { usage: fmtMoney(usage), covered: fmtMoney(covered), charged: fmtMoney(charged) },
      showGrants,
      grants: grantSlots,
      expiring,
      taxArm,
      taxNote,
      trueUpArm,
      trueUpNote,
      minor: { usage, covered, coveredPaid, coveredPromotional, charged, vatBase, deductionAfterTax, eligibleSpend, trueUp },
      allocation: results.map((r: any) => ({
        key: r.l.key,
        coveredMinor: r.covered,
        chargedMinor: r.charged,
        chargedQuantity: r.chargedQuantity,
        draws: r.draws,
      })),
      balances: order.map((g: any) => {
        const s = shown.find((x: any) => x.g === g);
        const remaining = g.units ? round(g.balance, g.dec) : g.balance;
        const b: { id: string; drawnMinor: number; drawnUnits?: number; remaining: number; warns: boolean } = {
          id: g.id,
          drawnMinor: g.drawnMinor,
          remaining,
          warns: s?.warns ?? false,
        };
        if (g.units) b.drawnUnits = g.drawnUnits;
        return b;
      }),
    };
  },
});
