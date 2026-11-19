import { deriver } from "docxcelerate/template";

/**
 * Decides everything a Japanese qualified invoice has to get right: the
 * heading, the checked registration number, the grouping by rate, the tax once
 * per rate, the totals and the transitional note. Installed by
 * `dxcl add japan-qualified-invoice`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The step
 * table, the patterns and the formatters are all declared inside `run`.
 */
export default deriver({
  name: "japanQualifiedInvoice",
  run: (jpInvoice: unknown) => {
    const fail = (message: string): never => {
      throw new Error("japanQualifiedInvoice: " + message);
    };
    // deno-lint-ignore no-explicit-any
    const inv = jpInvoice as any;
    if (inv === null || typeof inv !== "object") fail("jpInvoice must be an object");

    const yen = (n: number) => n.toLocaleString("ja-JP") + "円";
    const longDate = new Intl.DateTimeFormat("ja-JP", { dateStyle: "long", timeZone: "UTC" });
    const formatDate = (iso: string) => longDate.format(new Date(iso + "T00:00:00Z"));

    // Dates.
    if (typeof inv.transactionDate !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(inv.transactionDate)) {
      fail("jpInvoice.transactionDate must be an ISO 8601 date");
    }
    const date: string = inv.transactionDate.slice(0, 10);
    const parsed = new Date(date + "T00:00:00Z");
    if (isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== date) {
      fail("jpInvoice.transactionDate is not a real calendar date");
    }
    if (date < "2023-10-01") fail("jpInvoice.transactionDate is before 2023-10-01, when the system began");

    // Lines.
    const rawLines = inv.lines;
    if (!Array.isArray(rawLines) || rawLines.length < 1) fail("jpInvoice.lines must have at least one item");
    if (rawLines.length > 10) fail("jpInvoice.lines has more than 10 items, the limit of this component");
    for (const [i, line] of rawLines.entries()) {
      const at = "jpInvoice.lines[" + i + "]";
      if (typeof line?.description !== "string" || line.description.trim() === "") {
        fail(at + ".description is required");
      }
      if (!Number.isSafeInteger(line.amountYen) || Math.abs(line.amountYen) > 1e12) {
        fail(at + ".amountYen must be a whole number of yen no larger than 1,000,000,000,000");
      }
      if (line.rate !== "10" && line.rate !== "8") fail(at + '.rate must be "10" or "8"');
    }

    // Registration.
    const rawNumber = typeof inv.registrationNumber === "string" ? inv.registrationNumber.trim() : "";
    const registered = rawNumber !== "";
    const registration: string | null = registered ? rawNumber.replace(/[\s-]/g, "").toUpperCase() : null;
    if (registration !== null) {
      if (!/^T\d{13}$/.test(registration)) fail("jpInvoice.registrationNumber must be T and 13 digits");
      const digits = registration.slice(2).split("").reverse().map(Number);
      const sum = digits.reduce((acc, d, i) => acc + d * (i % 2 === 0 ? 1 : 2), 0);
      if (Number(registration[1]) !== 9 - (sum % 9)) fail("jpInvoice.registrationNumber check digit is wrong");
    }

    // Kind.
    const simplified = inv.simplified === true;
    if (!registered && simplified) fail("jpInvoice.simplified needs a registrationNumber");
    if (registered && !simplified && (typeof inv.buyerName !== "string" || inv.buyerName.trim() === "")) {
      fail("jpInvoice.buyerName is required on a qualified invoice");
    }
    if (registered) {
      if (!["floor", "ceil", "half-up"].includes(inv.rounding)) {
        fail('jpInvoice.rounding must be "floor", "ceil" or "half-up"');
      }
      if (!["exclusive", "inclusive"].includes(inv.basis)) {
        fail('jpInvoice.basis must be "exclusive" or "inclusive"');
      }
    }
    const inclusive = registered && inv.basis === "inclusive";

    // Group: 10% first, then 8%, each in input order.
    // deno-lint-ignore no-explicit-any
    const band = (rate: string) => rawLines.filter((l: any) => l.rate === rate);
    const standardLines = band("10");
    const reducedLines = band("8");
    const items = [
      // deno-lint-ignore no-explicit-any
      ...standardLines.map((l: any) => ({ description: l.description.trim(), amount: yen(l.amountYen) })),
      // deno-lint-ignore no-explicit-any
      ...reducedLines.map((l: any) => ({ description: l.description.trim() + " ※", amount: yen(l.amountYen) })),
    ];
    while (items.length < 10) items.push({ description: "", amount: "" });

    // Tax once per band, in integers, never per line.
    // deno-lint-ignore no-explicit-any
    const sum = (ls: any[]) => ls.reduce((acc: number, l: any) => acc + l.amountYen, 0);
    const taxOf = (total: number, r: number): number => {
      const n = total * r;
      const d = inclusive ? 100 + r : 100;
      const rem = n % d;
      const q = (n - rem) / d;
      if (inv.rounding === "floor") return q;
      if (inv.rounding === "ceil") return q + (rem > 0 ? 1 : 0);
      return q + (2 * rem >= d ? 1 : 0);
    };
    const summarise = (ls: unknown[], r: number) => {
      const total = ls.length === 0 ? null : sum(ls);
      if (total !== null && total < 0) {
        fail(r + "% lines total below zero: a return needs a 適格返還請求書, which this component does not produce");
      }
      const tax = total === null || !registered ? null : taxOf(total, r);
      const basisLabel = !registered ? "" : inclusive ? "（税込）" : "（税抜）";
      return {
        total,
        tax,
        out: {
          label: r + "%対象" + basisLabel,
          total: total === null ? null : yen(total),
          taxLabel: (inclusive ? "内消費税（" : "消費税（") + r + "%）",
          tax: tax === null ? null : yen(tax),
        },
      };
    };
    const standard = summarise(standardLines, 10);
    const reduced = summarise(reducedLines, 8);
    const bands = [standard, reduced];
    const net = (standard.total ?? 0) + (reduced.total ?? 0);
    const grand = registered && !inclusive ? net + (standard.tax ?? 0) + (reduced.tax ?? 0) : net;

    // Transitional deduction share, from the date, for a seller not registered.
    const steps = [
      { to: "2026-09-30", share: 80, cap: "10億円" },
      { to: "2028-09-30", share: 70, cap: "1億円" },
      { to: "2030-09-30", share: 50, cap: "1億円" },
      { to: "2031-09-30", share: 30, cap: "1億円" },
    ];
    const step = registered ? undefined : steps.find((s) => s.to >= date);
    const note = step
      ? "経過措置により、買手は本取引の仕入税額相当額の" + step.share + "%を控除できます（" + formatDate(step.to) +
        "までの取引。同一の売手からの課税仕入れが年" + step.cap + "を超える部分は対象外）。"
      : null;

    return {
      heading: !registered ? "請求書" : simplified ? "適格簡易請求書" : "適格請求書",
      registration,
      date: formatDate(date),
      amountHeader: !registered ? "金額" : inclusive ? "金額（税込）" : "金額（税抜）",
      items,
      standard: bands[0].out,
      reduced: bands[1].out,
      totalLabel: registered ? "合計（税込）" : "合計",
      total: yen(grand),
      legend: reducedLines.length > 0 ? "※は軽減税率対象" : null,
      note,
    };
  },
});
