import { deriver } from "docxcelerate/template";

/**
 * Works out the whole GST block of an Indian tax invoice. Installed by
 * `dxcl add india-gst-split`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The state
 * codes, the number-to-words function and the formatters live inside `run`.
 * Money is integer paise throughout; no step multiplies a rounded figure.
 */
export default deriver({
  name: "indiaGstSplit",
  run: (seller: unknown, placeOfSupply: unknown, customer: unknown, supply: unknown, lines: unknown) => {
    // deno-lint-ignore no-explicit-any
    const sel = seller as any;
    // deno-lint-ignore no-explicit-any
    const pos = placeOfSupply as any;
    // deno-lint-ignore no-explicit-any
    const cus = customer as any;
    // deno-lint-ignore no-explicit-any
    const sup = supply as any;
    // deno-lint-ignore no-explicit-any
    const items = lines as any;

    const validCode = (code: string) => {
      const n = Number(code);
      return /^\d{2}$/.test(code) && ((n >= 1 && n <= 38) || n === 97);
    };

    if (typeof sel?.gstin !== "string") throw new Error("seller.gstin must be a 15-character GSTIN");
    const gstin = sel.gstin.trim().toUpperCase();
    if (!/^[0-9A-Z]{15}$/.test(gstin)) throw new Error("seller.gstin must be a 15-character GSTIN");
    const supplierCode = gstin.slice(0, 2);
    if (!validCode(supplierCode)) throw new Error("seller.gstin must start with a known state code");
    if (sel.stateCode !== undefined && sel.stateCode !== supplierCode) {
      throw new Error("seller.stateCode does not match the state code in seller.gstin");
    }

    const posCode = pos?.stateCode;
    if (typeof posCode !== "string" || !(validCode(posCode) || posCode === "96")) {
      throw new Error("placeOfSupply.stateCode must be a two-digit state code (01 to 38, 96 or 97)");
    }
    if (typeof pos.stateName !== "string" || pos.stateName.trim() === "") {
      throw new Error("placeOfSupply.stateName is required");
    }

    const sez = cus?.isSez === true;
    const export_ = posCode === "96";
    if (sez && export_) throw new Error("customer.isSez cannot be true with place of supply 96");
    const zeroRated = sez || export_;
    if (typeof sup?.reverseCharge !== "boolean") throw new Error("supply.reverseCharge must be a boolean");
    const reverseCharge: boolean = sup.reverseCharge;
    const lut = sup.exportUnderLut === true;
    if (lut && !zeroRated) {
      throw new Error("supply.exportUnderLut applies only to an export or a supply to an SEZ");
    }

    if (!Array.isArray(items) || items.length === 0) throw new Error("lines must have at least one line");

    const intraState = !zeroRated && supplierCode === posCode;
    const stateTaxLabel = !intraState
      ? ""
      : ["04", "25", "26", "31", "35", "38", "97"].includes(posCode)
      ? "UTGST"
      : "SGST";

    const groups = new Map<string, { hsnSac: string; rate: number; taxable: number }>();
    for (const [i, line] of items.entries()) {
      const hsnSac = typeof line?.hsnSac === "string" ? line.hsnSac.trim() : "";
      if (!/^(\d{4}|\d{6}|\d{8})$/.test(hsnSac)) {
        throw new Error(`lines[${i}].hsnSac must be a 4, 6 or 8 digit HSN/SAC code`);
      }
      if (!Number.isSafeInteger(line.taxableMinor)) {
        throw new Error(`lines[${i}].taxableMinor must be an integer number of paise`);
      }
      const rate = line.rate;
      if (
        typeof rate !== "number" || !Number.isFinite(rate) || rate < 0 ||
        Math.abs(rate * 100 - Math.round(rate * 100)) > 1e-6
      ) {
        throw new Error(`lines[${i}].rate must be 0 or more, with at most two decimal places`);
      }
      const key = hsnSac + "|" + rate;
      const group = groups.get(key);
      if (group) group.taxable += line.taxableMinor;
      else groups.set(key, { hsnSac, rate, taxable: line.taxableMinor });
      if (!Number.isSafeInteger(groups.get(key)!.taxable)) throw new Error(`lines[${i}] overflows the safe integer range`);
    }

    // Half away from zero, so a credit note mirrors its invoice. `+ 0` drops -0.
    const roundDiv = (n: number, d: number) => Math.sign(n) * Math.floor((Math.abs(n) + d / 2) / d) + 0;
    const money = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR" });
    const signed = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", signDisplay: "exceptZero" });
    const fmt = (paise: number) => money.format(paise / 100);

    let taxableSum = 0;
    let centralSum = 0;
    let stateSum = 0;
    let integratedSum = 0;
    const rows = [...groups.values()].map((g) => {
      const bp = Math.round(g.rate * 100);
      const n = g.taxable * bp;
      if (Math.abs(n) > Number.MAX_SAFE_INTEGER) throw new Error(`tax on ${g.hsnSac} overflows the safe integer range`);
      // Each half is rounded on its own; the base is never split after rounding.
      const half = intraState ? roundDiv(n, 20000) : 0;
      const igst = intraState || lut ? 0 : roundDiv(n, 10000);
      taxableSum += g.taxable;
      centralSum += half;
      stateSum += half;
      integratedSum += igst;
      return {
        hsnSac: g.hsnSac,
        taxable: fmt(g.taxable),
        rate: lut ? "Nil" : bp / 100 + "%",
        halfRate: intraState ? bp / 200 + "%" : "",
        centralTax: intraState ? fmt(half) : "",
        stateTax: intraState ? fmt(half) : "",
        integratedTax: intraState ? "" : fmt(igst),
        totalTax: fmt(intraState ? half * 2 : igst),
      };
    });

    const taxSum = centralSum + stateSum + integratedSum;
    const unrounded = taxableSum + (reverseCharge ? 0 : taxSum);
    const grand = roundDiv(unrounded, 100) * 100;

    const ones = [
      "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve",
      "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen",
    ];
    const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
    const words = (n: number): string => {
      if (n === 0) return "Zero";
      if (n < 20) return ones[n];
      if (n < 100) return tens[Math.floor(n / 10)] + (n % 10 ? " " + ones[n % 10] : "");
      const [unit, name] = n >= 1e7
        ? [1e7, "Crore"]
        : n >= 1e5
        ? [1e5, "Lakh"]
        : n >= 1e3
        ? [1e3, "Thousand"]
        : [100, "Hundred"];
      const rest = n % unit;
      return words(Math.floor(n / unit)) + " " + name + (rest ? " " + words(rest) : "");
    };

    const taxTotal = fmt(taxSum);
    const zeroRatedMention = !zeroRated
      ? ""
      : (export_ ? "SUPPLY MEANT FOR EXPORT" : "SUPPLY TO SEZ UNIT OR SEZ DEVELOPER FOR AUTHORISED OPERATIONS") +
        (lut ? " UNDER BOND OR LETTER OF UNDERTAKING WITHOUT PAYMENT OF INTEGRATED TAX" : " ON PAYMENT OF INTEGRATED TAX");

    return {
      intraState,
      stateTaxLabel,
      placeOfSupply: pos.stateName.trim() + " (" + posCode + ")",
      reverseCharge: reverseCharge ? "Yes" : "No",
      zeroRatedMention,
      reverseChargeNote: reverseCharge
        ? `Tax of ${taxTotal} is payable by the recipient under reverse charge and is not included in the amount payable.`
        : "",
      rows,
      taxableTotal: fmt(taxableSum),
      centralTaxTotal: fmt(centralSum),
      stateTaxTotal: fmt(stateSum),
      integratedTaxTotal: fmt(integratedSum),
      taxTotal,
      roundOff: signed.format((grand - unrounded) / 100),
      grandTotal: fmt(grand),
      grandTotalInWords: (grand < 0 ? "Minus " : "") + "Rupees " + words(Math.abs(grand) / 100) + " Only",
    };
  },
});
