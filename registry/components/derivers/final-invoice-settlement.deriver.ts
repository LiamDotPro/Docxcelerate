// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Settles a final invoice against its advances: the whole supply by VAT rate,
 * each deducted advance exactly as printed on it, and the balance. Installed by
 * `dxcl add final-invoice-settlement`.
 *
 * Everything is integer minor units, and rates are basis points, so no step
 * goes through a float that could drift a penny. The VAT of a band is rounded
 * half away from zero once per band, never per line, and an advance's net and
 * VAT are subtracted as they were printed, never recomputed from its gross.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "finalInvoiceSettlement",
  run: (invoice: unknown, seller: unknown, locale: unknown) => {
    const inv = invoice as any;
    const loc = locale as any;

    const jurisdiction = inv?.jurisdiction;
    if (!["UK", "EU", "DE", "FR", "US"].includes(jurisdiction)) {
      throw new Error("invoice.jurisdiction must be one of UK, EU, DE, FR, US");
    }
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat(loc, { style: "currency", currency: inv?.currency });
    } catch {
      throw new Error("invoice.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmt = (minor: number) => money.format(minor / 10 ** exponent);
    const percent = new Intl.NumberFormat(loc, { style: "percent", maximumFractionDigits: 2 });
    const dates = new Intl.DateTimeFormat(loc, { dateStyle: "long", timeZone: "UTC" });

    const vatShown = jurisdiction !== "US" && (seller as any)?.smallBusinessExempt !== true;

    const safe = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isSafeInteger(value)) {
        throw new Error(field + " must be an integer of minor units");
      }
      return value;
    };
    // Basis points, so 5.5 is 550 and compares exactly.
    const bp = (rate: unknown, field: string): number => {
      if (!vatShown) {
        if (rate === undefined || rate === null || rate === 0) return 0;
        throw new Error(field + ": no VAT is shown, so " + field + " must be 0");
      }
      if (typeof rate !== "number" || !Number.isFinite(rate) || rate < 0 || rate >= 100) {
        throw new Error(field + " must be a VAT rate of at least 0 and under 100");
      }
      const scaled = rate * 100;
      if (Math.abs(scaled - Math.round(scaled)) > 1e-9) {
        throw new Error(field + " must have at most two decimals");
      }
      return Math.round(scaled);
    };
    // Half away from zero on exact integers.
    const roundHalfAway = (p: number) => {
      const q = Math.trunc(p / 10000);
      const r = p - q * 10000;
      return Math.abs(r) * 2 >= 10000 ? q + Math.sign(p) : q;
    };
    const pct = (b: number) => percent.format(b / 10000);
    const day = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));
    const blank = { number: "", date: "", rate: "", net: "", vat: "", gross: "" };

    // Group the whole supply into one band per rate.
    if (!Array.isArray(inv.lines) || inv.lines.length === 0) {
      throw new Error("invoice.lines must be a non-empty array");
    }
    const bands = new Map<number, { net: number; vat: number; fixed: boolean }>();
    inv.lines.forEach((line: any, i: number) => {
      const f = "invoice.lines[" + i + "]";
      const net = safe(line?.net, f + ".net");
      const key = bp(line?.vatRate, f + ".vatRate");
      const band = bands.get(key) ?? { net: 0, vat: 0, fixed: false };
      band.net = safe(band.net + net, "the lines at one rate");
      bands.set(key, band);
    });
    for (const [key, band] of bands) {
      if (band.net < 0) throw new Error("the lines at " + pct(key) + " net below zero; that is a credit note");
    }
    const lineRates = new Set(bands.keys());

    // Every advance is validated, paid or not, then listed by issue date.
    const rawAdvances: unknown[] = inv.advances == null ? [] : inv.advances;
    if (!Array.isArray(rawAdvances)) throw new Error("invoice.advances must be an array");
    const seen = new Set<string>();
    const advances = rawAdvances.map((a: any, i: number) => {
      const f = "invoice.advances[" + i + "]";
      const number = typeof a?.number === "string" ? a.number.trim() : "";
      if (number === "") throw new Error(f + ".number is required");
      if (seen.has(number)) throw new Error(f + ".number " + number + " is used twice");
      seen.add(number);
      const date = typeof a.issueDate === "string" ? a.issueDate.slice(0, 10) : "";
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || Number.isNaN(new Date(date + "T00:00:00Z").getTime())) {
        throw new Error(f + ".issueDate must be an ISO 8601 date");
      }
      const net = safe(a.net, f + ".net");
      if (net <= 0) throw new Error(f + ".net must be above zero");
      let vat = 0;
      if (vatShown) {
        vat = safe(a.vat, f + ".vat");
        if (vat < 0) throw new Error(f + ".vat must be 0 or more");
      } else if (a.vat !== undefined && a.vat !== 0) {
        throw new Error(f + ".vat: no VAT is shown, so it must be 0");
      }
      const rate = bp(a.vatRate, f + ".vatRate");
      if (typeof a.received !== "boolean") throw new Error(f + ".received must be true or false");
      return { i, number, date, net, vat, rate, received: a.received as boolean, supplyRate: a.supplyRate };
    });
    advances.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.i - b.i));

    // The UK counts an issued advance VAT invoice as a tax point, paid or not.
    const deductAll = jurisdiction === "UK" && vatShown;
    const deducted = advances.filter((a) => deductAll || a.received);
    const unpaid = advances.filter((a) => !a.received);
    if (deducted.length > 8) {
      throw new Error("more than eight advances to deduct; this component prints at most eight");
    }

    // A rate change between the advance and completion.
    let correction = 0;
    if (vatShown) {
      for (const a of deducted) {
        if (lineRates.has(a.rate)) continue;
        let target: number;
        if (a.supplyRate !== undefined && a.supplyRate !== null) {
          target = bp(a.supplyRate, "advance " + a.number + " supplyRate");
          if (!lineRates.has(target)) {
            throw new Error("advance " + a.number + " supplyRate is not a rate any line carries");
          }
        } else if (lineRates.size === 1) {
          target = [...lineRates][0];
        } else {
          throw new Error(
            "advance " + a.number + " is at " + pct(a.rate) + ", which no line carries; give its supplyRate",
          );
        }
        // The line an owner edits: which rule a rate change follows per jurisdiction. The UK keeps
        // the advance at its own rate; EU and DE re-rate the whole supply to the rate at
        // completion (BMF letter of 30 June 2020), FR on the BOFiP regularisation of acomptes at
        // the 2014 rate change. Some member states (Ireland, for one) keep the advance's rate:
        // an owner has to verify EU and FR.
        if (jurisdiction === "UK") {
          const from = bands.get(target)!;
          from.net -= a.net;
          if (from.net < 0) {
            throw new Error(
              "advance " + a.number +
                " exceeds the supply it was paid towards; correct the advance invoice with a credit note",
            );
          }
          const old = bands.get(a.rate) ?? { net: 0, vat: 0, fixed: true };
          old.net += a.net;
          old.vat += a.vat;
          bands.set(a.rate, old);
        } else {
          const p = a.net * target;
          if (!Number.isSafeInteger(p)) throw new Error("advance " + a.number + " is too large");
          correction += roundHalfAway(p) - a.vat;
        }
      }
    }

    // VAT per band, rounded once per band.
    const printedBands: Array<{ key: number; net: number; vat: number }> = [];
    for (const [key, band] of bands) {
      let vat = band.vat;
      if (!band.fixed) {
        const p = band.net * key;
        if (!Number.isSafeInteger(p)) throw new Error("the lines at " + pct(key) + " are too large");
        vat = roundHalfAway(p);
      }
      if (band.net === 0 && vat === 0) continue;
      printedBands.push({ key, net: band.net, vat });
    }
    printedBands.sort((a, b) => b.key - a.key);
    if (printedBands.length > 4) throw new Error("more than four VAT rates; this component prints at most four");

    const supplyNet = printedBands.reduce((sum, b) => sum + b.net, 0);
    const supplyVat = vatShown ? printedBands.reduce((sum, b) => sum + b.vat, 0) : 0;
    if (supplyNet <= 0) throw new Error("the supply nets to zero or less; that is not a final invoice");

    const deductedNet = deducted.reduce((sum, a) => sum + a.net, 0);
    const deductedVat = deducted.reduce((sum, a) => sum + a.vat, 0);
    const dueNet = supplyNet - deductedNet;
    const dueVat = supplyVat - deductedVat;
    const dueGross = dueNet + dueVat;
    const overpaid = dueGross < 0;
    if (overpaid && deducted.some((a) => !a.received)) {
      throw new Error(
        "the advances invoiced exceed the final total while some are unpaid; " +
          "cancel the unpaid advance invoice with a credit note first",
      );
    }
    const sign = overpaid ? -1 : 1;
    const vatCell = (minor: number) => (vatShown ? fmt(minor) : "");

    const list = (items: typeof advances) =>
      items.length === 0 ? "" : new Intl.ListFormat(loc, { type: "conjunction" }).format(
        items.map((a) => a.number + " (" + day(a.date) + ", " + fmt(a.net + a.vat) + ")"),
      );

    const slots = <T>(items: T[], size: number, empty: T): T[] => {
      const out = items.slice();
      while (out.length < size) out.push(empty);
      return out;
    };

    return {
      vatShown,
      bands: slots(
        (vatShown ? printedBands : []).map((b) => ({
          ...blank,
          rate: pct(b.key),
          net: fmt(b.net),
          vat: fmt(b.vat),
          gross: fmt(b.net + b.vat),
        })),
        4,
        blank,
      ),
      supply: { net: fmt(supplyNet), vat: vatCell(supplyVat), gross: fmt(supplyNet + supplyVat) },
      deducted: slots(
        deducted.map((a) => ({
          number: a.number,
          date: day(a.date),
          rate: vatShown ? pct(a.rate) : "",
          net: fmt(0 - a.net),
          vat: vatCell(0 - a.vat),
          gross: fmt(0 - (a.net + a.vat)),
        })),
        8,
        blank,
      ),
      overpaid,
      due: {
        net: fmt(sign * dueNet + 0),
        vat: vatCell(sign * dueVat + 0),
        gross: fmt(sign * dueGross + 0),
      },
      correction: vatShown && jurisdiction !== "UK" && correction !== 0 ? fmt(correction) : "",
      pendingOpen: deductAll ? "" : list(unpaid),
      pendingDeducted: deductAll ? list(unpaid) : "",
    };
  },
});
