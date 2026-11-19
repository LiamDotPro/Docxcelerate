import { deriver } from "docxcelerate/template";

/**
 * Decides whose VAT applies to a sale and at what rate. Installed by
 * `dxcl add vat-place-of-supply`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The EU
 * list, the helpers and the formatters are all declared inside `run`.
 */
export default deriver({
  name: "vatPlaceOfSupply",
  run: (seller: unknown, customer: unknown, supply: unknown, vatRates: unknown) => {
    // deno-lint-ignore no-explicit-any
    const sel = seller as any;
    // deno-lint-ignore no-explicit-any
    const cus = customer as any;
    // deno-lint-ignore no-explicit-any
    const sup = supply as any;
    // deno-lint-ignore no-explicit-any
    const rates = vatRates as any;

    const EU27 = [
      "AT", "BE", "BG", "CY", "CZ", "DE", "DK", "EE", "ES", "FI", "FR", "GR", "HR", "HU",
      "IE", "IT", "LT", "LU", "LV", "MT", "NL", "PL", "PT", "RO", "SE", "SI", "SK",
    ];
    const code = (x: unknown, path: string): string => {
      if (x == null) throw new Error(path + " is required");
      const v = String(x).trim().toUpperCase();
      const r = v === "EL" ? "GR" : v;
      if (!/^[A-Z]{2}$/.test(r)) throw new Error(path + " must be an ISO 3166-1 alpha-2 country code");
      return r;
    };

    const s = code(sel?.country, "seller.country");
    if (!EU27.includes(s) && s !== "GB" && s !== "XI") {
      throw new Error("seller.country must be an EU member state, GB or XI");
    }
    const c = code(cus?.country, "customer.country");
    const kind = sup?.kind;
    if (kind !== "goods" && kind !== "services" && kind !== "digital") {
      throw new Error("supply.kind must be goods, services or digital");
    }
    const goods = kind === "goods";
    const O = goods ? code(sup.shipsFrom ?? sel.country, "supply.shipsFrom") : s;
    const D = goods ? code(sup.shipsTo ?? cus.country, "supply.shipsTo") : c;

    const jur = (x: string) => (x === "XI" ? "GB" : x);
    const area = (x: string) => {
      if (EU27.includes(x)) return "EU";
      if (x === "XI") return goods ? "EU" : "UK";
      return x === "GB" ? "UK" : "outside";
    };
    if (goods && area(O) === "outside") throw new Error("supply.shipsFrom must be in the EU, GB or XI");

    const rateOf = (x: string): number => {
      const key = jur(x);
      const r = rates?.[key];
      if (typeof r !== "number" || !Number.isFinite(r) || r < 0 || r >= 100) {
        throw new Error("vatRates." + key + " must be the current standard VAT rate as a percentage");
      }
      return r;
    };
    const name = (x: string): string =>
      x === "XI" ? "Northern Ireland" : new Intl.DisplayNames("en", { type: "region" }).of(x) ?? x;
    const pct = (r: number) => new Intl.NumberFormat("en-GB", { maximumFractionDigits: 2 }).format(r) + "%";

    const isBusiness = cus?.isBusiness === true;
    const validB2B = isBusiness &&
      typeof cus.vatNumber === "string" && cus.vatNumber.trim() !== "" &&
      cus.vatNumberValidated === true;

    const standard = (at: string, suffix = "") => {
      const rate = rateOf(at);
      return {
        treatment: "standard", rate,
        line: "VAT charged at " + pct(rate) + " (" + name(at) + ")" + suffix + ".", note: null as string | null,
      };
    };
    const zero = (treatment: string, line: string) => ({ treatment, rate: 0, line, note: null as string | null });

    // 1. Domestic. GB and XI are each other's domestic market.
    if (jur(O) === jur(D)) return standard(O);

    // 2. Across VAT areas.
    if (area(O) !== area(D)) {
      if (goods) return zero("zero-rated-export", "Zero-rated: export of goods (" + name(D) + ").");
      const outside = zero(
        "outside-scope",
        "No VAT charged: outside the scope of " + area(O) + " VAT (" + name(D) + ").",
      );
      if (isBusiness) return outside;
      if (kind === "services") return standard(O);
      if (area(D) === "EU") return standard(D, " under the non-Union OSS scheme");
      return area(D) === "UK" ? standard(D) : outside;
    }

    // 3. Cross-border inside the EU.
    if (validB2B) {
      return goods
        ? zero("intra-community-supply", "No VAT charged: intra-Community supply of goods (" + name(O) + " to " + name(D) + ").")
        : zero("reverse-charge-services", "No VAT charged: reverse charge, VAT to be accounted for by the customer (" + name(D) + ").");
    }
    const note = isBusiness
      ? "VAT has been charged because the customer's VAT number is missing or has not been validated."
      : null;
    let result;
    if (kind === "services") {
      result = standard(O);
    } else {
      let destination = sel.ossRegistered === true || (goods && jur(O) !== jur(s));
      if (!destination) {
        const f = sel.euDistanceSales;
        const ok = (v: unknown) => typeof v === "number" && Number.isFinite(v) && v >= 0;
        if (!ok(f?.currentYear) || !ok(f?.priorYear)) {
          throw new Error("seller.euDistanceSales must give currentYear and priorYear as numbers of 0 or more");
        }
        const cents = (v: number) => Math.round(v * 100) / 100;
        destination = cents(f.currentYear) > 10000 || cents(f.priorYear) > 10000;
      }
      // Northern Ireland VAT is UK VAT, which the Union OSS does not cover.
      result = destination ? standard(D, D === "XI" ? "" : " under the Union OSS scheme") : standard(O);
    }
    return { ...result, note };
  },
});
