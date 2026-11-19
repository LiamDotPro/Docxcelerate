// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Turns one period's sales, returns and subsidiary-rights income into royalty
 * lines and their total. Installed by `dxcl add royalty-earnings-by-format`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "royaltyEarningsByFormat",
  run: (royaltyInput: unknown) => {
    const royalty = royaltyInput as any;

    const currency = royalty?.currency;
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency });
    } catch {
      throw new Error("royalty.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmtMoney = (minor: number) => money.format(minor / 10 ** exponent);
    const counts = new Intl.NumberFormat("en-GB");

    const nonNegInt = (v: unknown, field: string): number => {
      if (typeof v !== "number" || !Number.isInteger(v) || v < 0) {
        throw new Error(field + " must be a non-negative integer");
      }
      return v;
    };
    const posInt = (v: unknown, field: string): number => {
      if (typeof v !== "number" || !Number.isInteger(v) || v < 1) {
        throw new Error(field + " must be a positive integer");
      }
      return v;
    };
    // Basis points, so no percentage is ever carried as a float.
    const toBp = (v: unknown, field: string): number => {
      if (typeof v !== "number" || !Number.isFinite(v) || v < 0 || v > 100) {
        throw new Error(field + " must be a percentage from 0 to 100");
      }
      const bp = Math.round(v * 100);
      if (Math.abs(v * 100 - bp) > 1e-9) throw new Error(field + " must have at most two decimals");
      return bp;
    };
    const pct = (bp: number) => String(bp / 100) + "%";

    // Exact rationals; each line is rounded once, at the end.
    type Q = { n: bigint; d: bigint };
    const gcd = (a: bigint, b: bigint): bigint => {
      a = a < 0n ? -a : a;
      b = b < 0n ? -b : b;
      while (b) [a, b] = [b, a % b];
      return a;
    };
    const q = (n: bigint, d: bigint): Q => {
      const g = gcd(n, d) || 1n;
      return { n: n / g, d: d / g };
    };
    const add = (a: Q, b: Q): Q => q(a.n * b.d + b.n * a.d, a.d * b.d);
    const roundHalfAway = (v: Q): number => {
      const abs = v.n < 0n ? -v.n : v.n;
      const out = (2n * abs + v.d) / (2n * v.d);
      return Number(v.n < 0n ? -out : out);
    };

    // The high-discount clause, validated whether or not anything sells.
    const hd = royalty.highDiscount;
    const treatment: "fraction" | "net" = hd?.treatment ?? "fraction";
    if (hd && treatment !== "fraction" && treatment !== "net") {
      throw new Error('royalty.highDiscount.treatment must be "fraction" or "net"');
    }
    const steps: Array<{ bp: number; num: number; den: number }> = [];
    if (hd) {
      const stepOf = (t: unknown, num: unknown, den: unknown, field: string) => {
        const n = posInt(num, field + ".numerator");
        const d = posInt(den, field + ".denominator");
        if (n > d) throw new Error(field + ".numerator must not exceed its denominator");
        return { bp: toBp(t, field + ".thresholdPercent"), num: n, den: d };
      };
      steps.push(stepOf(hd.thresholdPercent ?? 55, hd.fraction?.numerator ?? 4, hd.fraction?.denominator ?? 5, "royalty.highDiscount" + (hd.fraction ? ".fraction" : "")));
      if (treatment === "fraction") {
        (Array.isArray(hd.deeper) ? hd.deeper : []).forEach((s: any, i: number) => {
          const step = stepOf(s?.thresholdPercent, s?.numerator, s?.denominator, "royalty.highDiscount.deeper[" + i + "]");
          if (step.bp <= steps[steps.length - 1].bp) {
            throw new Error("royalty.highDiscount.deeper[" + i + "].thresholdPercent must ascend");
          }
          steps.push(step);
        });
      }
    }

    const names: any = { hardback: "Hardback", paperback: "Paperback", ebook: "Ebook", audio: "Audiobook" };
    const rawFormats: any[] = royalty.formats === undefined || royalty.formats === null ? [] : royalty.formats;
    if (!Array.isArray(rawFormats)) throw new Error("royalty.formats must be an array");

    type Line = {
      description: string;
      units: string;
      basis: string;
      rate: string;
      amount: string;
      memo: boolean;
      amountMinor: number;
    };
    const lines: Line[] = [];
    const lifeToDate: string[] = [];

    rawFormats.forEach((f: any, fi: number) => {
      const at = "royalty.formats[" + fi + "]";
      if (!Object.prototype.hasOwnProperty.call(names, f?.format)) {
        throw new Error(at + '.format must be "hardback", "paperback", "ebook" or "audio"');
      }
      const name: string = typeof f.label === "string" && f.label !== "" ? f.label : names[f.format];
      if (f.basis !== "list" && f.basis !== "net") throw new Error(at + '.basis must be "list" or "net"');
      const isList = f.basis === "list";
      const list: number | undefined = f.listPriceMinor === undefined
        ? undefined
        : nonNegInt(f.listPriceMinor, at + ".listPriceMinor");
      if (isList && list === undefined) throw new Error(at + ".listPriceMinor is required on a list basis");

      if (!Array.isArray(f.tiers) || f.tiers.length === 0) throw new Error(at + ".tiers must be a non-empty array");
      const tiers = f.tiers.map((t: any, i: number) => ({
        from: posInt(t?.fromUnit, at + ".tiers[" + i + "].fromUnit"),
        bp: toBp(t?.percent, at + ".tiers[" + i + "].percent"),
      }));
      if (tiers[0].from !== 1) throw new Error(at + ".tiers[0].fromUnit must be 1");
      tiers.forEach((t: any, i: number) => {
        if (i > 0 && t.from <= tiers[i - 1].from) throw new Error(at + ".tiers[" + i + "].fromUnit must ascend");
      });
      const tierAt = (copy: number) => {
        let k = 0;
        tiers.forEach((t: any, i: number) => {
          if (t.from <= copy) k = i;
        });
        return k;
      };

      const before = nonNegInt(f.lifeToDateUnitsBefore, at + ".lifeToDateUnitsBefore");
      const rawSales: any[] = f.sales === undefined || f.sales === null ? [] : f.sales;
      if (!Array.isArray(rawSales)) throw new Error(at + ".sales must be an array");
      const sales = rawSales.map((s: any, i: number) => {
        const sf = at + ".sales[" + i + "]";
        const units = nonNegInt(s?.units, sf + ".units");
        const disc = s.discountPercent === undefined ? 0 : toBp(s.discountPercent, sf + ".discountPercent");
        const given = s.netReceiptsMinor === undefined ? undefined : nonNegInt(s.netReceiptsMinor, sf + ".netReceiptsMinor");
        return { units, disc, given, sf };
      });

      // Walk the sales, splitting each row where a tier boundary falls inside it.
      const groups = new Map<number, { tier: number; band: number; royalty: Q; receipts: Q; units: number }>();
      let c = before;
      for (const s of sales) {
        if (s.units === 0) continue;
        // Band: -1 full, 0..n a fraction step, 1000 the net switch.
        let band = -1;
        if (isList && hd) {
          steps.forEach((st, k) => {
            if (s.disc >= st.bp) band = treatment === "net" ? 1000 : k;
          });
        }
        const onNet = !isList || band === 1000;
        let rowReceipts: Q = q(0n, 1n);
        if (onNet) {
          if (s.given !== undefined) rowReceipts = q(BigInt(s.given), 1n);
          else if (list !== undefined) rowReceipts = q(BigInt(s.units) * BigInt(list) * BigInt(10000 - s.disc), 10000n);
          else throw new Error(s.sf + ".netReceiptsMinor is required when the format has no listPriceMinor");
        }
        let pos = c + 1;
        const last = c + s.units;
        while (pos <= last) {
          const ti = tierAt(pos);
          const end = ti === tiers.length - 1 ? last : Math.min(last, tiers[ti + 1].from - 1);
          const cu = BigInt(end - pos + 1);
          const bp = BigInt(tiers[ti].bp);
          let royaltyQ: Q;
          let receiptsQ: Q = q(0n, 1n);
          if (onNet) {
            receiptsQ = q(rowReceipts.n * cu, rowReceipts.d * BigInt(s.units));
            royaltyQ = q(receiptsQ.n * bp, receiptsQ.d * 10000n);
          } else {
            const num = band >= 0 ? BigInt(steps[band].num) : 1n;
            const den = band >= 0 ? BigInt(steps[band].den) : 1n;
            royaltyQ = q(cu * BigInt(list as number) * bp * num, 10000n * den);
          }
          const key = ti * 100000 + band + 1;
          const g = groups.get(key) ?? { tier: ti, band, royalty: q(0n, 1n), receipts: q(0n, 1n), units: 0 };
          g.royalty = add(g.royalty, royaltyQ);
          g.receipts = add(g.receipts, receiptsQ);
          g.units += end - pos + 1;
          groups.set(key, g);
          pos = end + 1;
        }
        c += s.units;
      }
      const after = c;
      lifeToDate.push(name + " " + counts.format(after));

      const sorted = [...groups.entries()].sort((a, b) => a[0] - b[0]).map((e) => e[1]);
      for (const g of sorted) {
        const t = tiers[g.tier];
        const range = g.tier === tiers.length - 1
          ? "copies " + counts.format(t.from) + " and over"
          : "copies " + counts.format(t.from) + "–" + counts.format(tiers[g.tier + 1].from - 1);
        const high = g.band >= 0 && g.band !== 1000 ? steps[g.band] : g.band === 1000 ? steps[0] : undefined;
        const minor = roundHalfAway(g.royalty);
        lines.push({
          description: name + ": " + range + (high ? ", high discount (" + String(high.bp / 100) + "% or more)" : ""),
          units: counts.format(g.units),
          basis: !isList || g.band === 1000
            ? fmtMoney(roundHalfAway(g.receipts)) + " net receipts"
            : fmtMoney(list as number) + " list",
          rate: g.band === 1000
            ? pct(t.bp) + " of net"
            : g.band >= 0
            ? steps[g.band].num + "/" + steps[g.band].den + " of " + pct(t.bp)
            : pct(t.bp),
          amount: fmtMoney(minor),
          memo: false,
          amountMinor: minor,
        });
      }

      // Returns: charged at the top tier reached, never reduced by high discount.
      const ret = f.returns;
      if (ret !== undefined && ret !== null) {
        const ru = nonNegInt(ret.units, at + ".returns.units");
        if (ru > 0) {
          const t = tiers[after === 0 ? 0 : tierAt(after)];
          let exact: Q;
          let basis: string;
          if (isList) {
            exact = q(-BigInt(ru) * BigInt(list as number) * BigInt(t.bp), 10000n);
            basis = fmtMoney(list as number) + " list";
          } else {
            const rr = nonNegInt(ret.netReceiptsMinor, at + ".returns.netReceiptsMinor");
            exact = q(-BigInt(rr) * BigInt(t.bp), 10000n);
            basis = fmtMoney(rr) + " net receipts";
          }
          const minor = roundHalfAway(exact);
          lines.push({
            description: name + ": returns",
            units: counts.format(-ru),
            basis,
            rate: pct(t.bp),
            amount: fmtMoney(minor),
            memo: false,
            amountMinor: minor,
          });
        }
      }
    });

    const rawRights: any[] = royalty.subRights === undefined || royalty.subRights === null ? [] : royalty.subRights;
    if (!Array.isArray(rawRights)) throw new Error("royalty.subRights must be an array");
    let hasMemo = false;
    rawRights.forEach((r: any, i: number) => {
      const at = "royalty.subRights[" + i + "]";
      for (const k of ["right", "licensee"]) {
        if (typeof r?.[k] !== "string" || r[k].trim() === "") throw new Error(at + "." + k + " is required");
      }
      const gross = nonNegInt(r.grossMinor, at + ".grossMinor");
      const tax = r.foreignTaxWithheldMinor === undefined ? 0 : nonNegInt(r.foreignTaxWithheldMinor, at + ".foreignTaxWithheldMinor");
      if (tax > gross) throw new Error(at + ".foreignTaxWithheldMinor must not exceed grossMinor");
      const shareBp = toBp(r.authorSharePercent, at + ".authorSharePercent");
      // The tax is already out of the base, so the memo row below never touches the total.
      const share = roundHalfAway(q(BigInt(gross - tax) * BigInt(shareBp), 10000n));
      lines.push({
        description: r.right + ", " + r.licensee,
        units: "",
        basis: tax > 0 ? fmtMoney(gross) + " less " + fmtMoney(tax) + " tax" : fmtMoney(gross),
        rate: pct(shareBp),
        amount: fmtMoney(share),
        memo: false,
        amountMinor: share,
      });
      if (tax > 0) {
        hasMemo = true;
        lines.push({
          description: "Foreign tax withheld by " + r.licensee + " (" + r.right + "), already deducted",
          units: "",
          basis: fmtMoney(tax),
          rate: "",
          amount: "",
          memo: true,
          amountMinor: 0,
        });
      }
    });

    const none = lines.length === 0;
    const totalMinor = lines.reduce((sum, l) => sum + l.amountMinor, 0);
    return {
      outcome: none ? "none" : "lines",
      negative: totalMinor < 0,
      hasMemo,
      lines,
      total: none ? "" : fmtMoney(totalMinor),
      totalAbs: none ? "" : fmtMoney(Math.abs(totalMinor)),
      totalMinor,
      lifeToDate: lifeToDate.join("; "),
    };
  },
});
