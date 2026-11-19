// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Places every voucher, coupon and gift card where its tax treatment puts it
 * and works out the tax that is actually due. Installed by
 * `dxcl add voucher-redemption-lines`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "voucherRedemptionLines",
  run: (sale: unknown) => {
    const s = sale as any;
    const MAX = 90_000_000_000;

    // Exact integer rounding, half away from zero. BigInt because a coupon times a
    // line can pass 2^53 well before either is unreasonable.
    const roundHalfAway = (numerator: number, denominator: number): number => {
      const n = BigInt(numerator);
      const d = BigInt(denominator);
      const abs = n < BigInt(0) ? -n : n;
      const q = (BigInt(2) * abs + d) / (BigInt(2) * d);
      return Number(n < BigInt(0) ? -q : q);
    };

    // Largest remainder: floors first, the leftover units one each to the largest
    // remainders, ties to the earlier line. A <= sum, so no share exceeds its line.
    const allocate = (amount: number, remaining: number[]): number[] => {
      const total = remaining.reduce((a, b) => a + b, 0);
      if (amount === 0 || total === 0) return remaining.map(() => 0);
      const A = BigInt(amount);
      const N = BigInt(total);
      const shares = remaining.map((r) => (A * BigInt(r)) / N);
      const rems = remaining.map((r) => (A * BigInt(r)) % N);
      let left = Number(A - shares.reduce((a, b) => a + b, BigInt(0)));
      const order = remaining.map((_, i) => i).sort((a, b) => (rems[a] === rems[b] ? a - b : rems[a] > rems[b] ? -1 : 1));
      const out = shares.map((x) => Number(x));
      for (const i of order) {
        if (left <= 0) break;
        out[i] += 1;
        left -= 1;
      }
      return out;
    };

    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > MAX) {
        throw new Error(field + " must be an integer of minor units from 0 to " + MAX);
      }
      return value;
    };
    const rateOf = (value: unknown, field: string): number => {
      if (
        typeof value !== "number" || !Number.isFinite(value) || value < 0 || value > 100 ||
        Math.abs(value * 1000 - Math.round(value * 1000)) > 1e-6
      ) {
        throw new Error(field + " must be a number from 0 to 100 with at most three decimal places");
      }
      return Math.round(value * 1000);
    };

    const currency = s?.currency;
    const jurisdiction = s?.jurisdiction;
    if (jurisdiction !== "UK" && jurisdiction !== "EU" && jurisdiction !== "US") {
      throw new Error('sale.jurisdiction must be "UK", "EU" or "US"');
    }
    const us = jurisdiction === "US";
    let money: Intl.NumberFormat;
    try {
      if (typeof currency !== "string" || !/^[A-Za-z]{3}$/.test(currency)) throw new Error();
      if (!(Intl as any).supportedValuesOf("currency").includes(currency.toUpperCase())) throw new Error();
      money = new Intl.NumberFormat(us ? "en-US" : "en-GB", { style: "currency", currency });
    } catch {
      throw new Error("sale.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmt = (minor: number) => money.format(minor / 10 ** exponent);
    const rateLabel = (milli: number) => new Intl.NumberFormat("en-GB", { maximumFractionDigits: 3 }).format(milli / 1000);
    const taxName = us ? "Sales tax" : "VAT";

    if (!Array.isArray(s.lines) || s.lines.length === 0) throw new Error("sale.lines must be a non-empty array");
    const seen = new Set<string>();
    const lines = s.lines.map((l: any, i: number) => {
      const f = "sale.lines[" + i + "]";
      if (typeof l?.key !== "string" || l.key === "") throw new Error(f + ".key must be a non-empty string");
      if (seen.has(l.key)) throw new Error(f + ".key is a duplicate of an earlier line: " + l.key);
      seen.add(l.key);
      return {
        key: l.key as string,
        net: minorOf(l.netMinor, f + ".netMinor"),
        rate: rateOf(l.ratePercent, f + ".ratePercent"),
      };
    });
    const remaining: number[] = lines.map((l: any) => l.net);

    const allowed = us
      ? ["retailer-coupon", "manufacturer-coupon", "gift-card", "promo-card"]
      : ["spv", "mpv", "discount-coupon"];
    const rawRedemptions = s.redemptions === undefined || s.redemptions === null ? [] : s.redemptions;
    if (!Array.isArray(rawRedemptions)) throw new Error("sale.redemptions must be an array");
    const reduceBase = us && s.manufacturerCouponsReduceBase === true;
    const reds = rawRedemptions.map((r: any, i: number) => {
      const f = "sale.redemptions[" + i + "]";
      if (!allowed.includes(r?.kind)) {
        throw new Error(f + ".kind must be one of " + allowed.join(", ") + " for " + jurisdiction);
      }
      const amount = minorOf(r.amountMinor, f + ".amountMinor");
      let eligible: number[] = lines.map((_: any, k: number) => k);
      if (r.appliesToLineKeys !== undefined) {
        if (!Array.isArray(r.appliesToLineKeys) || r.appliesToLineKeys.length === 0) {
          throw new Error(f + ".appliesToLineKeys must be a non-empty array of line keys");
        }
        eligible = r.appliesToLineKeys.map((key: unknown, k: number) => {
          const at = lines.findIndex((l: any) => l.key === key);
          if (at < 0) throw new Error(f + ".appliesToLineKeys[" + k + "] is not a line key on the sale: " + String(key));
          return at;
        });
        const dup = eligible.findIndex((at, k) => eligible.indexOf(at) !== k);
        if (dup >= 0) throw new Error(f + ".appliesToLineKeys[" + dup + "] is a duplicate");
        // Ties in the allocation go to the earlier sale line, not the earlier listed key.
        eligible.sort((a, b) => a - b);
      }
      let rate = 0;
      if (r.kind === "spv") {
        const rates = [...new Set(eligible.map((k) => lines[k].rate))];
        if (r.spvRatePercent !== undefined) {
          rate = rateOf(r.spvRatePercent, f + ".spvRatePercent");
          if (rates.some((x) => x !== rate)) {
            throw new Error(f + ".spvRatePercent differs from the rate of a line the voucher applies to");
          }
        } else {
          if (rates.length !== 1) {
            throw new Error(f + " is a single-purpose voucher over lines at different rates; it would be multi-purpose");
          }
          rate = rates[0];
        }
      }
      const code = typeof r.code === "string" && r.code !== "" ? " " + r.code : "";
      return { kind: r.kind as string, amount, eligible, rate, code };
    });

    const subtotalMinor = lines.reduce((a: number, l: any) => a + l.net, 0);

    // Discounts: before tax, each against what earlier ones left.
    const discounts: Array<{ description: string; amount: string }> = [];
    for (const r of reds) {
      const isDiscount = r.kind === "discount-coupon" || r.kind === "retailer-coupon" || r.kind === "promo-card" ||
        (r.kind === "manufacturer-coupon" && reduceBase);
      if (!isDiscount || r.amount === 0) continue;
      const room = r.eligible.reduce((a: number, k: number) => a + remaining[k], 0);
      const applied = Math.min(r.amount, room);
      const shares = allocate(applied, r.eligible.map((k: number) => remaining[k]));
      r.eligible.forEach((k: number, j: number) => (remaining[k] -= shares[j]));
      if (applied === 0) continue;
      const name = r.kind === "promo-card"
        ? "Promotional card"
        : r.kind === "manufacturer-coupon"
        ? "Manufacturer coupon"
        : "Coupon";
      const suffix = applied === r.amount
        ? ""
        : r.kind === "promo-card"
        ? " (remaining balance " + fmt(r.amount - applied) + ")"
        : " (applied " + fmt(applied) + " of " + fmt(r.amount) + ")";
      discounts.push({ description: name + r.code + suffix, amount: fmt(-applied) });
    }
    if (discounts.length > 4) {
      throw new Error("voucherRedemptionLines: " + discounts.length + " discount rows; this component prints at most 4");
    }

    // What a manufacturer coupon on the majority rule may still be set against.
    const afterDiscounts = remaining.slice();

    // Single-purpose vouchers: their VAT was accounted for at sale, so the part
    // they cover leaves the taxable amount and comes back as a payment.
    const spvTenders: Array<{ description: string; minor: number }> = [];
    for (const r of reds) {
      if (r.kind !== "spv" || r.amount === 0) continue;
      const E = r.eligible.reduce((a: number, k: number) => a + remaining[k], 0);
      const Eg = E + roundHalfAway(E * r.rate, 100000);
      let coveredNet: number;
      let coveredGross: number;
      let balance = 0;
      if (r.amount >= Eg) {
        coveredNet = E;
        coveredGross = Eg;
        balance = r.amount - Eg;
      } else {
        coveredNet = Math.min(E, roundHalfAway(r.amount * 100000, 100000 + r.rate));
        coveredGross = r.amount;
      }
      const shares = allocate(coveredNet, r.eligible.map((k: number) => remaining[k]));
      r.eligible.forEach((k: number, j: number) => (remaining[k] -= shares[j]));
      if (coveredGross === 0) continue;
      spvTenders.push({
        description: "Paid by voucher" + r.code + (balance > 0 ? " (remaining balance " + fmt(balance) + ")" : ""),
        minor: coveredGross,
      });
    }

    // Tax per band, rounded once per band, highest rate first.
    const byRate = new Map<number, number>();
    lines.forEach((l: any, k: number) => byRate.set(l.rate, (byRate.get(l.rate) ?? 0) + remaining[k]));
    const bandRows = [...byRate.entries()].filter(([, taxable]) => taxable > 0).sort((a, b) => b[0] - a[0])
      .map(([rate, taxable]) => ({ rate, taxable, tax: roundHalfAway(taxable * rate, 100000) }));
    if (bandRows.length > 4) {
      throw new Error("voucherRedemptionLines: " + bandRows.length + " tax bands; this component prints at most 4");
    }
    const bands = bandRows.map((b) => ({
      description: taxName + " at " + rateLabel(b.rate) + "% on " + fmt(b.taxable),
      amount: fmt(b.tax),
    }));
    const taxMinor = bandRows.reduce((a, b) => a + b.tax, 0);
    const spvPaid = spvTenders.reduce((a, t) => a + t.minor, 0);
    const totalMinor = bandRows.reduce((a, b) => a + b.taxable, 0) + taxMinor + spvPaid;

    // Payments, after the total. Each is capped at what is still due.
    let due = totalMinor;
    const tenders: Array<{ description: string; amount: string }> = [];
    let mpvNote = false;
    let manufacturerNote = false;
    for (const t of spvTenders) {
      due -= t.minor;
      tenders.push({ description: t.description, amount: fmt(-t.minor) });
    }
    for (const r of reds) {
      const pays = r.kind === "mpv" || r.kind === "gift-card" || (r.kind === "manufacturer-coupon" && !reduceBase);
      if (!pays || r.amount === 0) continue;
      if (r.kind === "manufacturer-coupon") {
        const room = r.eligible.reduce((a: number, k: number) => a + afterDiscounts[k], 0);
        const a = Math.min(r.amount, room, due);
        if (a === 0) continue;
        due -= a;
        manufacturerNote = true;
        tenders.push({
          description: "Manufacturer coupon" + r.code + (a < r.amount ? " (applied " + fmt(a) + " of " + fmt(r.amount) + ")" : ""),
          amount: fmt(-a),
        });
        continue;
      }
      const a = Math.min(r.amount, due);
      if (a === 0) continue;
      due -= a;
      if (r.kind === "mpv") mpvNote = true;
      tenders.push({
        description: (r.kind === "mpv" ? "Paid by voucher" : "Paid by gift card") + r.code +
          (r.amount > a ? " (remaining balance " + fmt(r.amount - a) + ")" : ""),
        amount: fmt(-a),
      });
    }
    if (tenders.length > 4) {
      throw new Error("voucherRedemptionLines: " + tenders.length + " payment rows; this component prints at most 4");
    }

    const pad = (rows: Array<{ description: string; amount: string }>) => {
      const out = rows.slice();
      while (out.length < 4) out.push({ description: "", amount: "" });
      return out;
    };

    return {
      subtotal: fmt(subtotalMinor),
      discounts: pad(discounts),
      bands: pad(bands),
      total: fmt(totalMinor),
      tenders: pad(tenders),
      amountDue: fmt(due),
      spvNote: spvTenders.length > 0,
      mpvNote,
      manufacturerNote,
      subtotalMinor,
      taxMinor,
      totalMinor,
      amountDueMinor: due,
    };
  },
});
