import { deriver } from "docxcelerate/template";

/**
 * Works out how GST shows on an AU or NZ invoice, and checks who it names.
 * Installed by `dxcl add gst-inclusive-statement`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level helper or constant would be a ReferenceError there. Every
 * helper, checksum and label is declared inside `run`.
 */
export default deriver({
  name: "gstInclusiveStatement",
  run: (input: unknown) => {
    // deno-lint-ignore no-explicit-any
    const gst = input as any;
    const blank = (v: unknown) => typeof v !== "string" || v.trim() === "";

    if (typeof gst !== "object" || gst === null) throw new Error("gst is required");
    if (gst.country !== "AU" && gst.country !== "NZ") throw new Error('gst.country must be "AU" or "NZ"');
    const au = gst.country === "AU";
    if (gst.rounding !== undefined && gst.rounding !== "total-invoice" && gst.rounding !== "taxable-sale") {
      throw new Error("gst.rounding must be total-invoice or taxable-sale");
    }
    if (!Array.isArray(gst.lines) || gst.lines.length === 0) throw new Error("gst.lines must be a non-empty array");

    const treatments = au
      ? ["taxable", "gst-free", "input-taxed", "margin-scheme"]
      : ["taxable", "zero-rated", "exempt"];
    let absSum = 0;
    // deno-lint-ignore no-explicit-any
    gst.lines.forEach((line: any, i: number) => {
      const at = "gst.lines[" + i + "]";
      if (blank(line?.description)) throw new Error(at + ".description is required");
      if (!Number.isInteger(line.amountMinor)) throw new Error(at + ".amountMinor must be an integer");
      if (!treatments.includes(line.treatment)) {
        throw new Error(at + ".treatment must be one of " + treatments.join(", "));
      }
      absSum += Math.abs(line.amountMinor);
    });
    // Keeps 2 x 69 x sum a safe integer in the rounding below.
    if (absSum > 1e13) throw new Error("gst.lines total is too large");

    const registered = gst.sellerRegistered === true;
    // deno-lint-ignore no-explicit-any
    const taxable = gst.lines.filter((l: any) => l.treatment === "taxable");
    const counted = registered && taxable.length > 0;
    if (counted && typeof gst.pricesIncludeGst !== "boolean") {
      throw new Error("gst.pricesIncludeGst is required when a line is taxable");
    }
    const inclusive = gst.pricesIncludeGst === true;

    // Rate p/q; half a cent away from zero.
    const p = au ? 1 : 3;
    const q = au ? 10 : 20;
    const div = (n: number, d: number) => Math.sign(n) * Math.floor((2 * Math.abs(n) + d) / (2 * d)) + 0;
    const gstOf = (a: number, incl: boolean) => incl ? div(a * p, p + q) : div(a * p, q);

    // deno-lint-ignore no-explicit-any
    const sum = (ls: any[]) => ls.reduce((t: number, l: any) => t + l.amountMinor, 0);
    const linesTotal = sum(gst.lines);
    const taxableTotal = sum(taxable);
    if (linesTotal < 0) throw new Error("gst.lines total is negative; adjustment notes are not supported");
    if (counted && taxableTotal < 0) {
      throw new Error("gst.lines taxable total is negative; adjustment notes are not supported");
    }

    let gstMinor = 0;
    if (counted) {
      gstMinor = gst.rounding === "taxable-sale"
        // deno-lint-ignore no-explicit-any
        ? taxable.reduce((t: number, l: any) => t + gstOf(l.amountMinor, inclusive), 0)
        : gstOf(taxableTotal, inclusive);
    }
    const totalMinor = counted && !inclusive ? linesTotal + gstMinor : linesTotal;

    // Checksums.
    const digitsOf = (v: string, field: string, pattern: RegExp, what: string) => {
      const d = v.replace(pattern, "");
      if (!/^\d+$/.test(d)) throw new Error(field + " is not a valid " + what);
      return d;
    };
    const abn = (v: string, field: string) => {
      const d = digitsOf(v, field, /\s/g, "ABN");
      const w = [10, 1, 3, 5, 7, 9, 11, 13, 15, 17, 19];
      let s = 0;
      if (d.length === 11) {
        for (let i = 0; i < 11; i++) s += (Number(d[i]) - (i === 0 ? 1 : 0)) * w[i];
      }
      if (d.length !== 11 || s % 89 !== 0) throw new Error(field + " is not a valid ABN");
      return d.slice(0, 2) + " " + d.slice(2, 5) + " " + d.slice(5, 8) + " " + d.slice(8);
    };
    const ird = (v: string, field: string) => {
      const d = digitsOf(v, field, /[\s-]/g, "IRD number");
      const n = Number(d);
      if ((d.length !== 8 && d.length !== 9) || n < 10000000 || n > 150000000) {
        throw new Error(field + " is not a valid IRD number");
      }
      const nine = d.padStart(9, "0");
      const check = (weights: number[]) => {
        let s = 0;
        for (let i = 0; i < 8; i++) s += Number(nine[i]) * weights[i];
        const r = s % 11;
        return r === 0 ? 0 : 11 - r;
      };
      let c = check([3, 2, 7, 6, 5, 4, 3, 2]);
      if (c === 10) c = check([7, 4, 3, 2, 5, 2, 7, 6]);
      if (c === 10 || c !== Number(nine[8])) throw new Error(field + " is not a valid IRD number");
      return d.length === 8
        ? d.slice(0, 2) + "-" + d.slice(2, 5) + "-" + d.slice(5)
        : d.slice(0, 3) + "-" + d.slice(3, 6) + "-" + d.slice(6);
    };
    const nzbn = (v: string, field: string) => {
      const d = digitsOf(v, field, /\s/g, "NZBN");
      let s = 0;
      for (let i = 0; i < 12 && d.length === 13; i++) s += Number(d[i]) * (i % 2 === 0 ? 1 : 3);
      if (d.length !== 13 || (10 - s % 10) % 10 !== Number(d[12])) throw new Error(field + " is not a valid NZBN");
    };

    // The seller's ID is printed whenever given: an AU ABN even when not
    // registered; an NZ GST number only when registered.
    let sellerTaxId: string | null = null;
    const given = !blank(gst.sellerTaxId);
    if (gst.sellerTaxId !== undefined && gst.sellerTaxId !== null && typeof gst.sellerTaxId !== "string") {
      throw new Error("gst.sellerTaxId must be a string");
    }
    if (given && au) sellerTaxId = "ABN " + abn(gst.sellerTaxId, "gst.sellerTaxId");
    if (given && !au && registered) sellerTaxId = "GST number " + ird(gst.sellerTaxId, "gst.sellerTaxId");

    const buyer = gst.buyer ?? {};
    const buyerNamed = !blank(buyer.name);
    let buyerTaxId: string | null = null;
    let buyerIdentifier: string | null = null;
    let buyerIdentified = false;
    if (au && !blank(buyer.abn)) {
      buyerTaxId = "Buyer ABN " + abn(buyer.abn, "gst.buyer.abn");
    }
    if (!au && buyer.identifier !== undefined && buyer.identifier !== null) {
      const labels: Record<string, string> = {
        address: "Buyer address",
        phone: "Buyer phone",
        email: "Buyer email",
        nzbn: "Buyer NZBN",
        url: "Buyer website",
        "trading-name": "Buyer trading name",
      };
      const kind = buyer.identifier.kind;
      if (typeof kind !== "string" || !Object.prototype.hasOwnProperty.call(labels, kind)) {
        throw new Error("gst.buyer.identifier.kind must be one of " + Object.keys(labels).join(", "));
      }
      if (!blank(buyer.identifier.value)) {
        const value = buyer.identifier.value.trim();
        if (kind === "nzbn") nzbn(value, "gst.buyer.identifier.value");
        buyerIdentifier = labels[kind] + ": " + value;
        buyerIdentified = true;
      }
    }

    // Thresholds are tested against the GST-inclusive total.
    if (counted) {
      if (au) {
        if (!given) throw new Error("gst.sellerTaxId: an AU tax invoice must show the supplier's ABN");
        if (totalMinor >= 100000 && !buyerNamed && buyerTaxId === null) {
          throw new Error("gst.buyer: an AU tax invoice of A$1,000 or more must show the buyer's name or ABN");
        }
      } else {
        if (totalMinor > 20000 && !given) {
          throw new Error("gst.sellerTaxId: an NZ tax invoice above NZ$200 must show the supplier's GST number");
        }
        if (totalMinor > 100000 && !(buyerNamed && buyerIdentified)) {
          throw new Error(
            "gst.buyer: an NZ tax invoice above NZ$1,000 must show the buyer's name and one identifier (gst.buyer.name and gst.buyer.identifier)",
          );
        }
      }
    }

    const money = new Intl.NumberFormat(au ? "en-AU" : "en-NZ", { style: "currency", currency: au ? "AUD" : "NZD" });
    const fmt = (minor: number) => money.format(minor / 100);

    const mixed = counted && taxable.length < gst.lines.length;
    // The statement stands in for the figure only when it is exactly true.
    const exact = counted && inclusive && !mixed && gstMinor === gstOf(totalMinor, true);
    const showGst = counted && !exact;

    return {
      heading: counted ? "Tax invoice" : "Invoice",
      sellerTaxId,
      buyerName: buyerNamed ? "Buyer: " + buyer.name.trim() : null,
      buyerTaxId,
      buyerIdentifier,
      // deno-lint-ignore no-explicit-any
      items: gst.lines.map((l: any) => ({
        description: l.description.trim() + (mixed && l.treatment === "taxable" ? " *" : ""),
        amount: fmt(l.amountMinor),
      })),
      subtotal: counted && !inclusive ? fmt(linesTotal) : null,
      gst: showGst ? fmt(gstMinor) : null,
      gstLabel: showGst ? (inclusive ? "GST included" : "GST") : "",
      total: fmt(totalMinor),
      totalLabel: showGst ? "Total including GST" : "Total",
      statement: exact
        ? "Total price includes GST."
        : registered && taxable.length === 0
        ? "No GST is charged on this invoice."
        : null,
      legend: mixed
        ? inclusive
          ? "* Taxable supply. The price of each marked item includes GST."
          : "* Taxable supply. GST is charged on each marked item."
        : null,
      marginNote: registered && au && gst.lines.some((l: { treatment: string }) => l.treatment === "margin-scheme")
        ? "Items sold under the margin scheme carry no GST you can claim as a credit."
        : null,
    };
  },
});
