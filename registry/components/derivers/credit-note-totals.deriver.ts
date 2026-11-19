// deno-lint-ignore-file no-explicit-any -- the input is request data nobody has checked.
import { deriver } from "docxcelerate/template";

/**
 * What a credit note credits, worked from the invoice it undoes. Installed by
 * `dxcl add credit-note-totals`.
 *
 * Every rate, price and exchange rate comes off the original invoice, never
 * from today's settings. Money is exact: each decimal is parsed from its
 * string into a scaled BigInt, a line is rounded once and VAT once per band,
 * half away from zero, so no float is ever multiplied.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level helper or constant would be a ReferenceError there. Every
 * helper is declared inside `run`. No placeholder: it is cheap, so previews
 * run it and show real figures.
 */
export default deriver({
  name: "creditNoteTotals",
  run: (creditNote: unknown) => {
   
    const note = creditNote as any;

    const fail = (message: string): never => {
      throw new Error("creditNoteTotals: " + message);
    };

    // A number to a scaled BigInt, from its own string, so nothing is rounded
    // by a float. Exponent notation (1e21) and excess decimals are refused.
    const parse = (value: unknown, places: number, field: string): bigint => {
      if (typeof value !== "number" || !Number.isFinite(value)) return fail(field + " must be a finite number");
      const text = String(value);
      if (!/^-?\d+(\.\d+)?$/.test(text)) return fail(field + " must be a plain decimal, got " + text);
      const [whole, fraction = ""] = text.replace("-", "").split(".");
      if (fraction.length > places) return fail(field + " has more than " + places + " decimal places");
      const scaled = BigInt(whole + fraction.padEnd(places, "0"));
      return text.startsWith("-") ? -scaled : scaled;
    };
    const nonNegative = (value: unknown, places: number, field: string): bigint => {
      const scaled = parse(value, places, field);
      return scaled < 0n ? fail(field + " must not be negative") : scaled;
    };
    // Half away from zero; b is positive.
    const divRound = (a: bigint, b: bigint): bigint => {
      const magnitude = (2n * (a < 0n ? -a : a) + b) / (2n * b);
      return a < 0n ? -magnitude : magnitude;
    };
    const min = (a: bigint, b: bigint) => (a < b ? a : b);
    const max = (a: bigint, b: bigint) => (a > b ? a : b);

    const original = note?.original;
    if (original === undefined || original === null || typeof original !== "object") {
      fail("creditNote.original is required");
    }
    const credit = note.credit;
    if (credit === undefined || credit === null || typeof credit !== "object") {
      fail("creditNote.credit is required");
    }
    const mode: string = credit.mode;
    if (mode !== "full" && mode !== "lines" && mode !== "amount") {
      fail('creditNote.credit.mode must be "full", "lines" or "amount"');
    }
    if (typeof original.number !== "string" || original.number.trim() === "") {
      fail("creditNote.original.number is required");
    }
    const number: string = original.number.trim();
    if (typeof original.issueDate !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(original.issueDate)) {
      fail("creditNote.original.issueDate must be an ISO 8601 date");
    }
    const issued = new Date(original.issueDate.slice(0, 10) + "T00:00:00Z");
    if (Number.isNaN(issued.getTime())) fail("creditNote.original.issueDate must be an ISO 8601 date");
    // UTC, so the date does not shift by a day on the engine.
    const issueDate = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" }).format(issued);

    const currency: string = typeof original.currency === "string" ? original.currency.toUpperCase() : "";
    let exponent = 0;
    try {
      exponent = new Intl.NumberFormat("en", { style: "currency", currency }).resolvedOptions().maximumFractionDigits ??
        2;
    } catch (error) {
      if (error instanceof RangeError) fail("creditNote.original.currency is not an ISO 4217 code: " + original.currency);
      throw error;
    }
    const scale = 10n ** BigInt(exponent);
    const money = (value: unknown, field: string) => nonNegative(value, exponent, field);

    const currencyFormat = new Intl.NumberFormat("en-GB", { style: "currency", currency });
    const fmt = (minor: bigint) => currencyFormat.format(Number(minor) / 10 ** exponent);
    const priceFormat = new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency,
      minimumFractionDigits: exponent,
      maximumFractionDigits: 4,
    });
    const quantityFormat = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 4 });
    const sterling = new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" });
    const sterlingRate = new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency: "GBP",
      minimumFractionDigits: 2,
      maximumFractionDigits: 6,
    });

    // Bands, keyed by rate in hundredths of a percent, or "null".
    if (!Array.isArray(original.bands) || original.bands.length === 0) {
      fail("creditNote.original.bands must be a non-empty array");
    }
    const keyOf = (rate: unknown, field: string): { key: string; hundredths: bigint | null } => {
      if (rate === null) return { key: "null", hundredths: null };
      const hundredths = nonNegative(rate, 2, field);
      return { key: String(hundredths), hundredths };
    };
    const bands: Array<{
      key: string;
      rate: bigint | null;
      label: string;
      net: bigint;
      vat: bigint;
      prevNet: bigint;
      prevVat: bigint;
      remNet: bigint;
      remVat: bigint;
      remGross: bigint;
    }> = [];
    original.bands.forEach((band: any, i: number) => {
      const at = "creditNote.original.bands[" + i + "]";
      const { key, hundredths } = keyOf(band?.vatRate, at + ".vatRate");
      if (bands.some((b) => b.key === key)) fail(at + ".vatRate repeats a rate already on the invoice");
      const net = money(band.net, at + ".net");
      const vat = money(band.vat, at + ".vat");
      if (hundredths === null && vat !== 0n) fail(at + ".vat must be 0 on a band with no VAT");
      const label = hundredths === null
        ? (typeof band.label === "string" && band.label.trim() !== "" ? band.label : "No VAT")
        : String(band.vatRate) + "%";
      bands.push({
        key,
        rate: hundredths,
        label,
        net,
        vat,
        prevNet: 0n,
        prevVat: 0n,
        remNet: 0n,
        remVat: 0n,
        remGross: 0n,
      });
    });
    const bandOf = (key: string) => bands.find((b) => b.key === key);

    // Earlier credit notes come off before anything is checked.
    const previous = (note.previousCredits ?? []) as any[];
    if (!Array.isArray(previous)) fail("creditNote.previousCredits must be an array");
    let previousGross = 0n;
    previous.forEach((earlier: any, i: number) => {
      const at = "creditNote.previousCredits[" + i + "]";
      if (!Array.isArray(earlier?.bands)) fail(at + ".bands must be an array");
      earlier.bands.forEach((band: any, j: number) => {
        const field = at + ".bands[" + j + "]";
        const target = bandOf(keyOf(band?.vatRate, field + ".vatRate").key);
        if (target === undefined) return fail(field + ".vatRate is not a rate on the original invoice");
        const net = money(band.net, field + ".net");
        const vat = money(band.vat, field + ".vat");
        target.prevNet += net;
        target.prevVat += vat;
        previousGross += net + vat;
      });
    });
    for (const band of bands) {
      band.remNet = band.net - band.prevNet;
      band.remVat = band.vat - band.prevVat;
      if (band.remNet < 0n || band.remVat < 0n) {
        fail("earlier credits exceed what invoice " + number + " charged in the " + band.label + " band");
      }
      band.remGross = band.remNet + band.remVat;
    }

    const credited = bands.map(() => ({ net: 0n, vat: 0n }));
    const lineRows: Array<{ description: string; quantity: string; unitPrice: string; rate: string; amount: string }> =
      [];

    if (mode === "full") {
      bands.forEach((b, i) => {
        credited[i] = { net: b.remNet, vat: b.remVat };
      });
      if (bands.every((b) => b.remGross === 0n)) fail("nothing is left to credit on invoice " + number);
    } else if (mode === "lines") {
      if (!Array.isArray(original.lines)) fail("creditNote.original.lines is required for a credit of lines");
      if (!Array.isArray(credit.lines) || credit.lines.length === 0) {
        fail("creditNote.credit.lines must list at least one line");
      }
      const originals = new Map<string, { description: string; qty: bigint; unitNet: bigint; key: string }>();
      original.lines.forEach((line: any, i: number) => {
        const at = "creditNote.original.lines[" + i + "]";
        if (typeof line?.id !== "string" || line.id === "") fail(at + ".id is required");
        if (originals.has(line.id)) fail(at + ".id repeats " + line.id);
        if (typeof line.description !== "string") fail(at + ".description is required");
        const qty = parse(line.qty, 4, at + ".qty");
        if (qty <= 0n) fail(at + ".qty must be more than 0");
        const { key } = keyOf(line.vatRate, at + ".vatRate");
        if (bandOf(key) === undefined) fail(at + ".vatRate is not a rate on the original invoice's bands");
        originals.set(line.id, {
          description: line.description,
          qty,
          unitNet: nonNegative(line.unitNet, 4, at + ".unitNet"),
          key,
        });
      });
      const returned = new Map<string, bigint>();
      const bandNet = new Map<string, bigint>();
      credit.lines.forEach((line: any, i: number) => {
        const at = "creditNote.credit.lines[" + i + "]";
        const source = originals.get(line?.id);
        if (source === undefined) return fail(at + ".id is not a line on the original invoice");
        if (returned.has(line.id)) fail(at + ".id repeats " + line.id);
        const qty = parse(line.qty, 4, at + ".qty");
        if (qty <= 0n) fail(at + ".qty must be more than 0");
        if (qty > source.qty) fail(at + ".qty is more than the " + quantityFormat.format(Number(source.qty) / 1e4) + " invoiced");
        returned.set(line.id, qty);
        // Quantity and price are both in ten-thousandths; rounded once, here.
        const net = divRound(qty * source.unitNet * scale, 10n ** 8n);
        bandNet.set(source.key, (bandNet.get(source.key) ?? 0n) + net);
        lineRows.push({
          description: source.description,
          quantity: quantityFormat.format(Number(qty) / 1e4),
          unitPrice: priceFormat.format(Number(source.unitNet) / 1e4),
          rate: bandOf(source.key)!.label,
          amount: fmt(net),
        });
      });
      bands.forEach((b, i) => {
        const net = bandNet.get(b.key);
        if (net === undefined) return;
        // Every line at this rate returned whole: copy what is left, so a full
        // return at one rate cannot drift a penny from the original.
        const whole = b.prevNet === 0n && b.prevVat === 0n &&
          [...originals].every(([id, line]) => line.key !== b.key || returned.get(id) === line.qty);
        credited[i] = whole
          ? { net: b.remNet, vat: b.remVat }
          : { net, vat: b.rate === null ? 0n : divRound(net * b.rate, 10000n) };
      });
    } else {
      const amount = money(credit.amount, "creditNote.credit.amount");
      if (amount <= 0n) fail("creditNote.credit.amount must be more than 0");
      const eligible = bands.map((b, i) => ({ b, i })).filter(({ b }) => b.remGross > 0n);
      const pool = eligible.reduce((sum, { b }) => sum + b.remGross, 0n);
      if (amount > pool) {
        fail("creditNote.credit.amount " + fmt(amount) + " is more than the " + fmt(pool) + " still creditable");
      }
      // Largest remainder: floors first, then one minor unit each to the biggest
      // remainders, the earlier band winning a tie, so the shares add up exactly.
      const shares = eligible.map(({ b, i }) => ({
        i,
        share: amount * b.remGross / pool,
        remainder: amount * b.remGross % pool,
      }));
      const leftover = Number(amount - shares.reduce((sum, s) => sum + s.share, 0n));
      [...shares].sort((x, y) => (x.remainder === y.remainder ? x.i - y.i : x.remainder > y.remainder ? -1 : 1))
        .slice(0, leftover).forEach((s) => s.share += 1n);
      for (const { i, share } of shares) {
        const b = bands[i];
        if (share === b.remGross) {
          credited[i] = { net: b.remNet, vat: b.remVat };
          continue;
        }
        // Back the VAT out of the gross, then keep both halves inside what the
        // band has left: that interval is never empty.
        const backed = b.rate === null ? 0n : divRound(share * b.rate, 10000n + b.rate);
        const vat = max(max(0n, share - b.remNet), min(min(b.remVat, share), backed));
        credited[i] = { net: share - vat, vat };
      }
    }

    bands.forEach((b, i) => {
      if (credited[i].net > b.remNet || credited[i].vat > b.remVat) {
        fail(
          "the " + b.label + " band can take at most " + fmt(b.remNet) + " net and " + fmt(b.remVat) +
            " VAT after earlier credits; this credit is " + fmt(credited[i].net) + " net and " +
            fmt(credited[i].vat) + " VAT.",
        );
      }
    });
    const netMinor = credited.reduce((sum, c) => sum + c.net, 0n);
    const vatMinor = credited.reduce((sum, c) => sum + c.vat, 0n);
    const totalMinor = netMinor + vatMinor;
    if (totalMinor === 0n) fail("nothing to credit");

    // Sterling: only when the invoice carried UK VAT in another currency. One
    // conversion of the total, because summing converted bands would not agree.
    let vatGBPMinor: bigint | null = null;
    if (currency !== "GBP" && original.fxToGBP !== undefined && original.fxToGBP !== null) {
      const fx = parse(original.fxToGBP, 6, "creditNote.original.fxToGBP");
      if (fx <= 0n) fail("creditNote.original.fxToGBP must be more than 0");
      const firstFull = mode === "full" && previous.length === 0 && original.vatGBP !== undefined &&
        original.vatGBP !== null;
      vatGBPMinor = firstFull
        ? nonNegative(original.vatGBP, 2, "creditNote.original.vatGBP")
        : divRound(vatMinor * fx * 100n, scale * 10n ** 6n);
    }

    const basis = mode === "full"
      ? previous.length === 0
        ? "This credit note cancels invoice " + number + " in full."
        : "This credit note cancels what remained of invoice " + number + " after earlier credits."
      : mode === "lines"
      ? "This credit note credits the items listed below from invoice " + number + "."
      : "This credit note is a credit of " + fmt(totalMinor) + " against invoice " + number +
        ", spread across its VAT rates in proportion to what was charged at each.";
    const converted = vatGBPMinor === null
      ? ""
      : " VAT is converted to sterling at that invoice's rate of " +
        sterlingRate.format(Number(parse(original.fxToGBP, 6, "creditNote.original.fxToGBP")) / 1e6) + " to 1 " +
        currency + ".";

    const shown = bands.map((b, i) => ({ b, c: credited[i] })).filter(({ c }) => c.net !== 0n || c.vat !== 0n);

    return {
      basis,
      original: "Credited against invoice " + number + " of " + issueDate + ", at the VAT rates that invoice charged." +
        converted,
      previouslyCredited: previous.length === 0 ? null : "Earlier credit notes against this invoice: " + fmt(previousGross) + ".",
      hasLines: mode === "lines",
      lines: lineRows,
      multipleBands: shown.length > 1,
      bands: shown.map(({ b, c }) => ({
        label: b.label,
        net: fmt(c.net),
        vat: fmt(c.vat),
        netMinor: Number(c.net),
        vatMinor: Number(c.vat),
      })),
      net: fmt(netMinor),
      vat: fmt(vatMinor),
      total: fmt(totalMinor),
      netMinor: Number(netMinor),
      vatMinor: Number(vatMinor),
      totalMinor: Number(totalMinor),
      vatGBP: vatGBPMinor === null ? null : sterling.format(Number(vatGBPMinor) / 100),
      vatGBPMinor: vatGBPMinor === null ? null : Number(vatGBPMinor),
    };
  },
});
