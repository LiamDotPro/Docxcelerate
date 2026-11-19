// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Restates an invoice's VAT bands in the currency the VAT is owed in, checks
 * the rate against the jurisdiction's rule, and words the check. Installed by
 * `dxcl add vat-in-local-currency`.
 *
 * Money is BigInt minor units and the rate is a BigInt over a power of ten, so
 * no float is ever multiplied and a credit note is the exact mirror of its
 * invoice. It is self-contained on purpose: it travels to the engine as source
 * text, so every helper and table lives inside `run`.
 */
export default deriver({
  name: "vatInLocalCurrency",
  run: (input: unknown) => {
    const conv = input as any;
    const fail = (message: string): never => {
      throw new Error("vatInLocalCurrency: " + message);
    };

    const currencyOf = (value: unknown, field: string) => {
      if (typeof value !== "string") fail(field + " is required");
      let money: Intl.NumberFormat;
      try {
        money = new Intl.NumberFormat("en-GB", { style: "currency", currency: value as string });
      } catch {
        fail(field + " must be an ISO 4217 code");
      }
      return { code: value as string, money: money!, exponent: money!.resolvedOptions().maximumFractionDigits as number };
    };

    // UTC day indexes, so the engine's time zone never shifts a day.
    const isoDay = (value: unknown, field: string): number => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        fail(field + " must be an ISO 8601 date");
      }
      const text = (value as string).slice(0, 10);
      const [y, m, d] = text.split("-").map(Number);
      const day = Date.UTC(y, m - 1, d) / 86_400_000;
      if (new Date(day * 86_400_000).toISOString().slice(0, 10) !== text) fail(field + " is not a real date: " + text);
      return day;
    };
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const months = new Intl.DateTimeFormat("en-GB", { month: "long", year: "numeric", timeZone: "UTC" });
    const fmt = (day: number) => dates.format(new Date(day * 86_400_000));
    const fmtMonth = (day: number) => months.format(new Date(day * 86_400_000));
    const monthKey = (day: number) => new Date(day * 86_400_000).toISOString().slice(0, 7);

    // Half away from zero on exact integers, so +x and -x round to mirror images.
    const roundHalfAway = (n: bigint, d: bigint) => {
      const abs = n < 0n ? -n : n;
      const q = (2n * abs + d) / (2n * d);
      return n < 0n ? -q : q;
    };

    const JURISDICTIONS = ["UK", "EU", "DE", "PL", "BG"];
    const jurisdiction = conv?.jurisdiction;
    if (!JURISDICTIONS.includes(jurisdiction)) fail("vatConversion.jurisdiction must be one of " + JURISDICTIONS.join(", "));
    const inv = currencyOf(conv?.invoiceCurrency, "vatConversion.invoiceCurrency");
    const loc = currencyOf(conv?.vatCurrency, "vatConversion.vatCurrency");
    const taxPoint = isoDay(conv?.taxPointDate, "vatConversion.taxPointDate");
    if (jurisdiction === "PL" && conv?.issueDate == null) fail("vatConversion.issueDate is required in PL");
    const issue = conv?.issueDate == null ? null : isoDay(conv.issueDate, "vatConversion.issueDate");

    // The currency each jurisdiction owes VAT in; an owner edits it here.
    const owed: Record<string, string> = {
      UK: "GBP",
      DE: "EUR",
      PL: "PLN",
      BG: taxPoint < Date.UTC(2026, 0, 1) / 86_400_000 ? "BGN" : "EUR",
    };
    if (jurisdiction !== "EU" && loc.code !== owed[jurisdiction]) {
      fail(
        "vatConversion.vatCurrency must be " + owed[jurisdiction] + " in " + jurisdiction + " at this tax point, got " +
          loc.code,
      );
    }

    const same = inv.code === loc.code;
    const text = (n: bigint, c: { money: Intl.NumberFormat; exponent: number }) => {
      const abs = n < 0n ? -n : n;
      const scale = 10n ** BigInt(c.exponent);
      const frac = c.exponent === 0 ? "" : "." + (abs % scale).toString().padStart(c.exponent, "0");
      return c.money.format(Number((n < 0n ? "-" : "") + (abs / scale).toString() + frac));
    };
    if (same) {
      return {
        converted: false,
        sameCurrency: loc.code,
        vatCurrency: loc.code,
        invoiceCurrency: inv.code,
        statement: "",
        bands: [] as any[],
        total: { net: "", vat: "", localNet: "", localVat: "" },
        hasAdvances: false,
        advances: [] as any[],
        warnings: [] as any[],
      };
    }

    const minorOf = (value: unknown, field: string): bigint => {
      const t = String(value).trim();
      const m = /^(-?)(\d+)(?:\.(\d+))?$/.exec(t);
      if (!m || (m[3] ?? "").length > inv.exponent) {
        fail(field + " must be a decimal with at most " + inv.exponent + " decimal places, got " + t);
      }
      const minor = BigInt(m![2] + (m![3] ?? "").padEnd(inv.exponent, "0"));
      return m![1] ? -minor : minor;
    };

    const bandsOf = (list: unknown, path: string) => {
      if (!Array.isArray(list) || list.length === 0) fail(path + " must be a non-empty array");
      const seen = new Set<string>();
      return (list as any[]).map((b, i) => {
        const at = path + "[" + i + "]";
        const t = String(b?.ratePercent).trim();
        const m = /^(\d+)(?:\.(\d{1,4}))?$/.exec(t);
        if (!m) fail(at + ".ratePercent must be a percentage with at most 4 decimal places, got " + t);
        const frac = (m![2] ?? "").replace(/0+$/, "");
        const rate = m![1].replace(/^0+(?=\d)/, "") + (frac ? "." + frac : "");
        if (seen.has(rate)) fail(at + ".ratePercent repeats the " + rate + "% band");
        seen.add(rate);
        return { rate, net: minorOf(b?.net, at + ".net"), vat: minorOf(b?.vat, at + ".vat") };
      });
    };

    const rateOf = (r: any, path: string) => {
      const t = String(r?.value).trim();
      const m = /^(\d+)(?:\.(\d{1,10}))?$/.exec(t);
      if (!m) fail(path + ".value must be a decimal with at most 10 decimal places, got " + t);
      const R = BigInt(m![1] + (m![2] ?? ""));
      if (R === 0n) fail(path + ".value must be greater than zero");
      const scale = 10n ** BigInt((m![2] ?? "").length);
      let multiply: boolean;
      if (r.base === inv.code && r.quote === loc.code) multiply = true;
      else if (r.base === loc.code && r.quote === inv.code) multiply = false;
      else fail(path + " must be a pair of " + inv.code + " and " + loc.code + ", got " + r.base + "/" + r.quote);
      const SOURCES = ["market-selling", "hmrc-period", "ecb", "nbp", "bmf-monthly", "fixed-euro", "agreed"];
      if (!SOURCES.includes(r.source)) fail(path + ".source must be one of " + SOURCES.join(", "));
      const euroBgn = [inv.code, loc.code].sort().join("/") === "BGN/EUR";
      if (r.source === "fixed-euro" && !euroBgn) fail(path + ".source fixed-euro is only for the EUR/BGN pair");
      if (euroBgn) {
        if (r.source !== "fixed-euro") fail(path + ".source must be fixed-euro for the EUR/BGN pair");
        if (r.base !== "EUR" || t !== "1.95583") fail(path + " must be 1 EUR = 1.95583 BGN, never rounded or inverted");
      }
      const day = r.source === "fixed-euro" ? null : isoDay(r.rateDate, path + ".rateDate");
      return { R, scale, multiply: multiply!, text: t, base: r.base as string, quote: r.quote as string, source: r.source as string, day };
    };
    type Rate = ReturnType<typeof rateOf>;

    // One rounding, in the VAT currency's minor units, on exact integers.
    const convert = (a: bigint, r: Rate) => {
      const toScale = 10n ** BigInt(loc.exponent);
      const fromScale = 10n ** BigInt(inv.exponent);
      return r.multiply
        ? roundHalfAway(a * r.R * toScale, r.scale * fromScale)
        : roundHalfAway(a * r.scale * toScale, r.R * fromScale);
    };

    const UK = jurisdiction === "UK";
    const sourceLabel = (r: Rate) => {
      const d = r.day === null ? "" : fmt(r.day);
      const labels: Record<string, string> = {
        "market-selling": UK ? "UK market selling rate of " + d : "selling rate of " + d,
        "hmrc-period": "HMRC period rate for " + (r.day === null ? "" : fmtMonth(r.day)),
        "bmf-monthly": "BMF monthly average rate for " + (r.day === null ? "" : fmtMonth(r.day)),
        ecb: "ECB reference rate of " + d,
        nbp: "NBP average rate of " + d,
        agreed: UK ? "rate agreed with HMRC, dated " + d : "rate agreed with the tax authority, dated " + d,
        "fixed-euro": "the irrevocably fixed euro conversion rate",
      };
      return labels[r.source];
    };
    const rateLine = (r: Rate) => "1 " + r.base + " = " + r.text + " " + r.quote + ", " + sourceLabel(r);

    const rate = rateOf(conv.rate, "vatConversion.rate");
    const bands = bandsOf(conv.bands, "vatConversion.bands");

    // Advances: each band keeps the rate its own chargeable event was declared at.
    if (conv.advances != null && !Array.isArray(conv.advances)) fail("vatConversion.advances must be an array");
    const advances = ((conv.advances ?? []) as any[]).map((a, i) => {
      const at = "vatConversion.advances[" + i + "]";
      if (typeof a?.invoiceNumber !== "string" || !a.invoiceNumber.trim()) fail(at + ".invoiceNumber is required");
      const chargeable = isoDay(a.chargeableDate, at + ".chargeableDate");
      if (chargeable > taxPoint) fail(at + ".chargeableDate is after vatConversion.taxPointDate");
      const abands = bandsOf(a.bands, at + ".bands");
      for (const b of abands) {
        if (!bands.some((x) => x.rate === b.rate)) fail(at + ".bands has a " + b.rate + "% band the invoice does not");
        if (b.net < 0n || b.vat < 0n) fail(at + ".bands must not be negative");
      }
      return { number: a.invoiceNumber.trim() as string, chargeable, bands: abands, rate: rateOf(a.rate, at + ".rate") };
    });
    if (advances.length) {
      bands.forEach((b, i) => {
        if (b.net < 0n || b.vat < 0n) fail("vatConversion.bands[" + i + "] must not be negative when there are advances");
      });
    }
    const left = bands.map((b) => {
      let net = b.net;
      let vat = b.vat;
      for (const a of advances) {
        for (const x of a.bands) {
          if (x.rate === b.rate) {
            net -= x.net;
            vat -= x.vat;
          }
        }
      }
      if (advances.length && (net < 0n || vat < 0n)) fail("the advances exceed the " + b.rate + "% band");
      return { rate: b.rate, net, vat };
    });

    // Warn on the legal choices the data shows; throwing is for facts that make the figures wrong.
    const accepted: Record<string, string[]> = {
      UK: ["market-selling", "hmrc-period", "agreed"],
      EU: ["market-selling", "ecb"],
      DE: ["bmf-monthly", "ecb", "agreed"],
      PL: ["nbp", "ecb"],
      BG: ["fixed-euro", "ecb", "market-selling"],
    };
    const names: Record<string, string> = {
      "market-selling": "the market selling rate",
      "hmrc-period": "an HMRC period rate",
      ecb: "the ECB rate",
      nbp: "the NBP rate",
      "bmf-monthly": "the BMF monthly average",
      "fixed-euro": "the fixed euro rate",
      agreed: "a rate agreed with the tax authority",
    };
    const place: Record<string, string> = { UK: "the UK", EU: "the EU", DE: "Germany", PL: "Poland", BG: "Bulgaria" };
    const easter = (y: number) => {
      const a = y % 19, b = Math.floor(y / 100), c = y % 100;
      const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
      const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
      const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
      const m = Math.floor((a + 11 * h + 22 * l) / 451);
      const month = Math.floor((h + l - 7 * m + 114) / 31), day = ((h + l - 7 * m + 114) % 31) + 1;
      return Date.UTC(y, month - 1, day) / 86_400_000;
    };
    const holiday = (day: number) => {
      const d = new Date(day * 86_400_000);
      const y = d.getUTCFullYear();
      const md = d.toISOString().slice(5, 10);
      const fixed = ["01-01", "01-06", "05-01", "05-03", "08-15", "11-01", "11-11", "12-25", "12-26"];
      if (y >= 2025) fixed.push("12-24");
      const e = easter(y);
      return fixed.includes(md) || day === e + 1 || day === e + 60;
    };
    const working = (day: number) => {
      const wd = new Date(day * 86_400_000).getUTCDay();
      return wd !== 0 && wd !== 6 && !holiday(day);
    };
    const upper = (s: string) => s.charAt(0).toUpperCase() + s.slice(1);

    const warnings: Array<{ text: string }> = [];
    const check = (r: Rate, reference: number, prefix: string, refWord: string) => {
      if (r.day === null) return;
      const out: string[] = [];
      const ok = accepted[jurisdiction];
      if (!ok.includes(r.source)) {
        const list = ok.map((s) => names[s]);
        out.push(
          upper(sourceLabel(r)) + " is not a rate " + place[jurisdiction] + " accepts for VAT; use " +
            (list.length > 1 ? list.slice(0, -1).join(", ") + " or " + list[list.length - 1] : list[0]) + ".",
        );
      }
      if (jurisdiction !== "PL" && r.day > reference) {
        out.push("The exchange rate is dated " + fmt(r.day) + ", after the tax point of " + fmt(reference) + ".");
      }
      if ((r.source === "hmrc-period" || r.source === "bmf-monthly") && monthKey(r.day) !== monthKey(reference)) {
        out.push(
          "A" + (r.source === "hmrc-period" ? "n HMRC period" : " BMF monthly") + " rate applies only to its own month: this one is for " +
            fmtMonth(r.day) + ", the tax point is in " + fmtMonth(reference) + ".",
        );
      }
      if (jurisdiction === "PL") {
        let lwd = reference - 1;
        while (!working(lwd)) lwd--;
        if (r.day !== lwd) {
          const head = "In Poland the rate is the one from the last working day before " + refWord + ", " + fmt(lwd) +
            "; this one is dated " + fmt(r.day) + ", ";
          out.push(
            head + (r.day === reference
              ? "the day itself."
              : r.day > reference
              ? "after " + refWord + "."
              : !working(r.day)
              ? "which is not a working day."
              : "before the last working day."),
          );
        }
      }
      for (const t of out) warnings.push({ text: prefix + t });
    };
    const issuedFirst = issue !== null && issue < taxPoint;
    check(rate, issuedFirst ? issue! : taxPoint, "", issuedFirst ? "the date of issue" : "the tax point");
    for (const a of advances) {
      check(a.rate, a.chargeable, "Advance invoice " + a.number + ": ", "the tax point");
    }

    const sum = (xs: bigint[]) => xs.reduce((s, x) => s + x, 0n);
    const lined = left.map((b) => ({ ...b, localNet: convert(b.net, rate), localVat: convert(b.vat, rate) }));
    const line = (n: bigint, v: bigint, ln: bigint, lv: bigint) => ({
      net: text(n, inv),
      vat: text(v, inv),
      localNet: text(ln, loc),
      localVat: text(lv, loc),
    });
    const mainLine = rateLine(rate);

    return {
      converted: true,
      sameCurrency: "",
      vatCurrency: loc.code,
      invoiceCurrency: inv.code,
      statement: "VAT on this invoice is accounted for in " + loc.code + ". Each rate band is converted separately at " +
        mainLine + ", and the " + loc.code + " totals are the sum of the bands." +
        (advances.length
          ? " The bands are what this invoice charges after the advances below, each of which keeps the rate it was declared at."
          : ""),
      bands: lined.map((b) => ({ rate: b.rate + "%", ...line(b.net, b.vat, b.localNet, b.localVat) })),
      total: line(sum(lined.map((b) => b.net)), sum(lined.map((b) => b.vat)), sum(lined.map((b) => b.localNet)), sum(lined.map((b) => b.localVat))),
      hasAdvances: advances.length > 0,
      advances: advances.map((a) => {
        const c = a.bands.map((b) => ({ ...b, ln: convert(b.net, a.rate), lv: convert(b.vat, a.rate) }));
        return {
          label: "Advance invoice " + a.number + ", chargeable " + fmt(a.chargeable),
          rateLine: rateLine(a.rate),
          ...line(sum(c.map((b) => b.net)), sum(c.map((b) => b.vat)), sum(c.map((b) => b.ln)), sum(c.map((b) => b.lv))),
        };
      }),
      warnings,
    };
  },
});
