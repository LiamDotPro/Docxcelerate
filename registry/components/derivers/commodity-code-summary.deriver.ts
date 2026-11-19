import { deriver } from "docxcelerate/template";

/**
 * Checks each goods line's tariff code, origin and mass, groups the lines the
 * way a customs declaration is split, and totals them. Installed by
 * `dxcl add commodity-code-summary`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * factor and message is declared inside `run`. Every string it returns is
 * ready to print, and "" means there is nothing to print.
 */
export default deriver({
  name: "commodityCodeSummary",
  run: (customs: unknown) => {
    // deno-lint-ignore no-explicit-any
    const c = customs as any;

    let money: Intl.NumberFormat;
    try {
      // Intl accepts any well-formed three letters, so XXZ would format; the list is what knows it is not a currency.
      if (!Intl.supportedValuesOf("currency").includes(c?.currency)) throw new Error("unknown");
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency: c?.currency });
    } catch {
      throw new Error("customs.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmtMoney = (minor: number) => money.format(minor / 10 ** exponent);

    const lengths = [6, 8, 10];
    if (!lengths.includes(c?.destinationCodeLength)) {
      throw new Error("customs.destinationCodeLength must be one of " + lengths.join(", "));
    }
    const L: number = c.destinationCodeLength;

    const isNum = (v: unknown): v is number => typeof v === "number" && Number.isFinite(v) && v >= 0;
    const hasGross = c.grossMassKg !== undefined && c.grossMassKg !== null;
    if (hasGross && !isNum(c.grossMassKg)) {
      throw new Error("customs.grossMassKg must be a finite number of 0 or more");
    }

    const qtyFmt = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 3 });
    const kgFmt = new Intl.NumberFormat("en-GB", { minimumFractionDigits: 3, maximumFractionDigits: 3 });
    // Strips binary noise (0.1 + 0.2) before anything is rounded or compared.
    const clean = (n: number) => Math.round(n * 1e6) / 1e6;
    const wholeGrams = (n: number) => Math.floor(clean(n) + 0.5);
    const kg = (grams: number) => kgFmt.format(grams / 1000);
    const factors: Record<string, number> = { kg: 1000, g: 1, lb: 453.59237 };

    const raw: unknown[] = Array.isArray(c.lines) ? c.lines : [];
    if (raw.length === 0) {
      return {
        empty: true,
        basis: "",
        attention: "",
        lines: [],
        groups: [],
        totalNetMass: "",
        totalValue: "",
        grossMass: "",
        grossWarning: "",
      };
    }

    interface Entry {
      code: string;
      origin: string;
      preference: boolean;
      grams: number;
      valueMinor: number;
      supp: { code: string; quantity: number } | null;
    }
    const entries: Entry[] = [];
    let flagged = 0;

    const lines = raw.map((item, i) => {
      // deno-lint-ignore no-explicit-any
      const l = item as any;
      const at = "customs.lines[" + i + "]";
      if (!isNum(l?.quantity)) throw new Error(at + ".quantity must be a finite number of 0 or more");
      if (!isNum(l.netMass?.value)) throw new Error(at + ".netMass.value must be a finite number of 0 or more");
      if (factors[l.netMass.unit] === undefined) throw new Error(at + '.netMass.unit must be "kg", "g" or "lb"');
      if (!Number.isInteger(l.valueMinor) || l.valueMinor < 0) {
        throw new Error(at + ".valueMinor must be a non-negative integer");
      }
      const s = l.supplementaryUnit;
      const hasSupp = s !== undefined && s !== null;
      if (hasSupp && (typeof s.code !== "string" || s.code === "")) {
        throw new Error(at + ".supplementaryUnit.code must not be empty");
      }
      if (hasSupp && !isNum(s.quantity)) {
        throw new Error(at + ".supplementaryUnit.quantity must be a finite number of 0 or more");
      }

      const warnings: string[] = [];
      const tariff = l.tariffCode;
      if (typeof tariff === "number") {
        warnings.push("The tariff code was supplied as a number, so a leading zero may have been lost. Supply it as text.");
      }
      const code = String(tariff ?? "").replace(/[\s.]/g, "");
      if (code === "") {
        warnings.push("No tariff code: this line cannot be declared.");
      } else if (!/^\d+$/.test(code)) {
        warnings.push("“" + String(tariff) + "” is not a tariff code.");
      } else if (code.length < L) {
        warnings.push("The tariff code has " + code.length + " digits; this destination needs " + L + ".");
      }

      const origin = String(l.origin ?? "").trim().toUpperCase();
      if (origin === "") {
        warnings.push("No country of origin.");
      } else if (!/^[A-Z]{2}$/.test(origin)) {
        warnings.push("“" + origin + "” is not a two-letter country code.");
      }

      const grams = clean(l.netMass.value * factors[l.netMass.unit]);
      const preference = l.preference === true;
      entries.push({
        code,
        origin,
        preference,
        grams,
        valueMinor: l.valueMinor,
        supp: hasSupp ? { code: s.code, quantity: s.quantity } : null,
      });
      if (warnings.length > 0) flagged += 1;

      const ref = typeof l.ref === "string" ? l.ref.trim() : "";
      const description = typeof l.description === "string" ? l.description : "";
      return {
        item: (ref === "" ? "Line " + (i + 1) : ref) + (description === "" ? "" : " — " + description),
        tariffCode: code === "" ? "None" : code,
        origin: origin === "" ? "None" : origin,
        preference: preference ? "Yes" : "No",
        quantity: qtyFmt.format(l.quantity),
        supplementary: hasSupp ? qtyFmt.format(s.quantity) + " " + s.code : "",
        netMass: kgFmt.format(wholeGrams(grams) / 1000),
        value: fmtMoney(l.valueMinor),
        warning: warnings.join(" "),
      };
    });

    // Missing sorts last, so a broker scanning down finds real codes first.
    const cmp = (a: string, b: string) => a === b ? 0 : a === "" ? 1 : b === "" ? -1 : a < b ? -1 : 1;
    const buckets = new Map<string, Entry[]>();
    for (const e of entries) {
      const key = JSON.stringify([e.code, e.origin, e.preference]);
      const bucket = buckets.get(key);
      if (bucket) bucket.push(e);
      else buckets.set(key, [e]);
    }
    const sorted = [...buckets.values()].sort((a, b) =>
      cmp(a[0].code, b[0].code) || cmp(a[0].origin, b[0].origin) || Number(a[0].preference) - Number(b[0].preference)
    );

    let totalGrams = 0;
    let totalMinor = 0;
    const groups = sorted.map((members) => {
      const first = members[0];
      const grams = wholeGrams(members.reduce((sum, e) => sum + e.grams, 0));
      const minor = members.reduce((sum, e) => sum + e.valueMinor, 0);
      totalGrams += grams;
      totalMinor += minor;

      const given = members.filter((e) => e.supp !== null);
      let supplementary = "";
      let warning = "";
      if (given.length > 0) {
        const unit = given[0].supp!.code;
        if (given.length === members.length && given.every((e) => e.supp!.code === unit)) {
          supplementary = qtyFmt.format(clean(given.reduce((sum, e) => sum + e.supp!.quantity, 0))) + " " + unit;
        } else {
          warning = "The lines in this group give different supplementary units, so none is totalled.";
        }
      }
      return {
        tariffCode: first.code === "" ? "None" : first.code,
        origin: first.origin === "" ? "None" : first.origin,
        preference: first.preference ? "Yes" : "No",
        lineCount: members.length === 1 ? "1 line" : members.length + " lines",
        supplementary,
        netMass: kg(grams),
        value: fmtMoney(minor),
        warning,
      };
    });

    let grossMass = "";
    let grossWarning = "";
    if (hasGross) {
      const grossGrams = Math.round(c.grossMassKg * 1000);
      grossMass = kg(grossGrams);
      if (totalGrams > grossGrams) {
        grossWarning = "The net mass of the lines, " + kg(totalGrams) + " kg, is more than the declared gross mass of " +
          grossMass + " kg. One of them is wrong.";
      }
    }

    return {
      empty: false,
      basis: "Tariff codes are checked to " + L + " digits for this destination. Masses are net, in kilograms.",
      attention: flagged === 0 ? "" : flagged + " of " + raw.length + " lines " + (flagged === 1 ? "needs" : "need") +
        " attention before these goods can be declared.",
      lines,
      groups,
      totalNetMass: kg(totalGrams),
      totalValue: fmtMoney(totalMinor),
      grossMass,
      grossWarning,
    };
  },
});
