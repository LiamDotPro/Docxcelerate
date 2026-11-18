// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Works out a US sales tax breakdown: one row per jurisdiction, each on its own
 * base, adding up to the tax charged. Installed by
 * `dxcl add us-sales-tax-breakdown`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The order
 * map, the rounding, the rate parser and the formatters are all declared inside
 * it.
 *
 * It rounds half away from zero, so a credit memo mirrors its invoice to the
 * cent. It throws rather than guess whether something is taxable: a line with
 * no `taxable` flag, a rate that is not a decimal string and a missing state
 * row are all errors, because a guess is wrong in one direction or the other.
 *
 * It has no placeholder, so previews run it.
 */
export default deriver({
  name: "usSalesTaxBreakdown",
  run: (salesTax: unknown) => {
   
    const st = salesTax as any;
    if (st === null || typeof st !== "object") throw new Error("salesTax is required");

    const collectedBy = st.collectedBy ?? "seller";
    if (collectedBy !== "seller" && collectedBy !== "marketplace") {
      throw new Error('salesTax.collectedBy must be "seller" or "marketplace"');
    }

    if (collectedBy === "marketplace") {
      const name = typeof st.marketplaceName === "string" && st.marketplaceName.trim() !== ""
        ? st.marketplaceName.trim()
        : "the marketplace facilitator";
      return {
        marketplace: "Sales tax on this order was collected and remitted by " + name +
          " as the marketplace facilitator, so it is not charged on this invoice.",
        rows: [] as Array<{ name: string; taxable: string; rate: string; tax: string; taxableMinor: number; taxMinor: number }>,
        total: { taxable: "", rate: "", tax: "", taxMinor: 0 },
        exemptions: [] as Array<{ reason: string; items: string }>,
        shippingWarning: "",
      };
    }

    const roundHalfAway = (n: bigint, d: bigint): bigint =>
      n >= 0n ? (2n * n + d) / (2n * d) : -((-2n * n + d) / (2n * d));
    const abs = (n: bigint): bigint => n < 0n ? -n : n;

    const currency: string = st.currency ?? "USD";
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-US", { style: "currency", currency });
    } catch {
      throw new Error("salesTax.currency must be an ISO 4217 code");
    }
    const digits = money.resolvedOptions().maximumFractionDigits ?? 2;
    const fmt = (minor: bigint): string => money.format(Number(minor) / 10 ** digits);
    const pct = (millionths: bigint): string => {
      const whole = millionths / 10000n;
      let frac = String(millionths % 10000n).padStart(4, "0");
      while (frac.length > 2 && frac.endsWith("0")) frac = frac.slice(0, -1);
      return whole + "." + frac + "%";
    };

    const jur = st.jurisdictions;
    if (!Array.isArray(jur) || jur.length === 0) {
      throw new Error("salesTax.jurisdictions must be a non-empty array");
    }
    const levels = ["state", "county", "city", "special"];
    const parsed = jur.map((j: any, index: number) => {
      const at = "salesTax.jurisdictions[" + index + "]";
      if (j === null || typeof j !== "object") throw new Error(at + " must be an object");
      if (!levels.includes(j.level)) throw new Error(at + ".level must be one of " + levels.join(", "));
      if (typeof j.name !== "string" || j.name.trim() === "") throw new Error(at + ".name is required");
      const m = typeof j.rate === "string" ? /^(0|[1-9]\d*)(\.(\d{1,6}))?$/.exec(j.rate) : null;
      const rate = m ? BigInt(m[1]) * 1000000n + BigInt((m[3] ?? "").padEnd(6, "0")) : null;
      if (rate === null || rate >= 1000000n) {
        throw new Error(at + '.rate must be a decimal string such as "0.0625"');
      }
      if (j.perItemCapMinor !== undefined && j.perItemCapMinor !== null &&
        (!Number.isSafeInteger(j.perItemCapMinor) || j.perItemCapMinor <= 0)) {
        throw new Error(at + ".perItemCapMinor must be a positive integer");
      }
      if (j.exemptCategories !== undefined && j.exemptCategories !== null && !Array.isArray(j.exemptCategories)) {
        throw new Error(at + ".exemptCategories must be an array");
      }
      return {
        index,
        order: levels.indexOf(j.level),
        name: j.name as string,
        rate,
        taxesShipping: j.taxesShipping as boolean | undefined,
        exempt: (j.exemptCategories ?? []) as string[],
        cap: j.perItemCapMinor == null ? null : BigInt(j.perItemCapMinor),
      };
    });
    if (parsed.filter((j: { order: number }) => j.order === 0).length !== 1) {
      throw new Error("salesTax.jurisdictions must have exactly one state entry");
    }
    const js = [...parsed].sort((a, b) => a.order - b.order || a.index - b.index);
    const state = js[0];

    const rawLines = st.lines ?? [];
    if (!Array.isArray(rawLines)) throw new Error("salesTax.lines must be an array");
    const lines = rawLines.map((l: any, index: number) => {
      const at = "salesTax.lines[" + index + "]";
      if (l === null || typeof l !== "object") throw new Error(at + " must be an object");
      if (!Number.isSafeInteger(l.amountMinor)) throw new Error(at + ".amountMinor must be an integer");
      if (typeof l.taxable !== "boolean") throw new Error(at + ".taxable must be true or false");
      return { index, line: l, amount: BigInt(l.amountMinor) };
    });

    const ship = st.shipping === undefined || st.shipping === null ? null : st.shipping;
    if (ship !== null && !Number.isSafeInteger(ship.amountMinor)) {
      throw new Error("salesTax.shipping.amountMinor must be an integer");
    }
    const shipAmount = ship === null ? 0n : BigInt(ship.amountMinor);
    const combined = ship !== null && ship.combinedWithHandling === true && ship.separatelyStated !== false;

    const taxedBy = (j: { exempt: string[] }, l: { taxable: boolean; category?: string }): boolean =>
      l.taxable === true && !(l.category !== undefined && j.exempt.includes(l.category));

    const all = lines.reduce((sum: bigint, l: { amount: bigint }) => sum + l.amount, 0n);
    let warn = false;
    const bases = js.map((j) => {
      let goods = 0n;
      let base = 0n;
      for (const { line, amount } of lines) {
        if (!taxedBy(j, line)) continue;
        goods += amount;
        base += j.cap === null ? amount : amount >= 0n ? (amount < j.cap ? amount : j.cap) : (amount > -j.cap ? amount : -j.cap);
      }
      if (ship !== null) {
        const ships = j.taxesShipping ?? state.taxesShipping ?? false;
        if (ships || ship.separatelyStated === false || ship.combinedWithHandling === true) {
          const none = all === 0n || goods === 0n || (goods > 0n) !== (all > 0n);
          const share = none ? 0n : abs(goods) >= abs(all) ? shipAmount : roundHalfAway(shipAmount * abs(goods), abs(all));
          base += share;
          if (!ships && share !== 0n) warn = true;
        }
      }
      return base;
    });

    const exact = js.map((j, i) => bases[i] * j.rate);
    const taxes = exact.map((e) => roundHalfAway(e, 1000000n));
    const totalTax = roundHalfAway(exact.reduce((a, b) => a + b, 0n), 1000000n);
    const residual = totalTax - taxes.reduce((a, b) => a + b, 0n);
    const target = state.rate !== 0n ? 0 : js.findIndex((j) => j.rate !== 0n);
    if (target >= 0) taxes[target] += residual;

    const rows = js.map((j, i) => ({
      name: j.name,
      taxable: fmt(bases[i]),
      rate: pct(j.rate),
      tax: fmt(taxes[i]),
      taxableMinor: Number(bases[i]),
      taxMinor: Number(taxes[i]),
    }));
    const same = bases.every((b) => b === bases[0]);
    const total = {
      taxable: same ? fmt(bases[0]) : "",
      rate: same ? pct(js.reduce((sum, j) => sum + j.rate, 0n)) : "",
      tax: fmt(totalTax),
      taxMinor: Number(totalTax),
    };

    const list = new Intl.ListFormat("en-US", { type: "conjunction" });
    const groups: Array<{ reason: string; items: string[] }> = [];
    for (const { index, line, amount } of lines) {
      let reason: string;
      if (!line.taxable) {
        reason = typeof line.exemptReason === "string" && line.exemptReason.trim() !== ""
          ? line.exemptReason.trim()
          : "Not taxable";
      } else {
        const skipping = js.filter((j) => !taxedBy(j, line));
        if (skipping.length === 0) continue;
        reason = skipping.length === js.length
          ? "Exempt from sales tax (" + line.category + ")"
          : "Not taxed by " + list.format(skipping.map((j) => j.name));
      }
      if (typeof line.description !== "string" || line.description.trim() === "") {
        throw new Error("salesTax.lines[" + index + "].description is required");
      }
      const item = line.description + " (" + fmt(amount) + ")";
      const group = groups.find((g) => g.reason === reason);
      if (group) group.items.push(item);
      else groups.push({ reason, items: [item] });
    }

    const shippingWarning = warn && combined
      ? "Shipping and handling were charged together as " + fmt(shipAmount) + ". " + state.name +
        " does not tax delivery charges stated on their own, but a combined shipping and handling charge is taxable, so its taxable share has been taxed here. Invoice handling as a separate line to keep delivery out of the tax."
      : "";

    return {
      marketplace: "",
      rows,
      total,
      exemptions: groups.map((g) => ({ reason: g.reason, items: g.items.join("; ") })),
      shippingWarning,
    };
  },
});
