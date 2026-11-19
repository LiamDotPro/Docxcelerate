import { deriver } from "docxcelerate/template";

/**
 * Decides the whole sales-tax block of a Canadian invoice: which taxes the
 * province of supply charges, the rate in force on the tax point date, each
 * tax on its own base, the registration numbers worth printing, and what is
 * missing for the customer to claim input tax credits. Installed by
 * `dxcl add canadian-sales-tax`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The
 * province map, the rate table and every helper are declared inside `run`.
 */
export default deriver({
  name: "canadianSalesTax",
  run: (salesTax: unknown, seller: unknown, customer: unknown) => {
    // deno-lint-ignore no-explicit-any
    const st = salesTax as any;
    // deno-lint-ignore no-explicit-any
    const sel = seller as any;
    // deno-lint-ignore no-explicit-any
    const cus = customer as any;

    // Which taxes each province of supply charges: [federal, provincial].
    const structure: Record<string, [string, string | null]> = {
      ON: ["HST", null], NS: ["HST", null], NB: ["HST", null], NL: ["HST", null], PE: ["HST", null],
      BC: ["GST", "PST"], SK: ["GST", "PST"], MB: ["GST", "RST"], QC: ["GST", "QST"],
      AB: ["GST", null], NT: ["GST", null], NU: ["GST", null], YT: ["GST", null],
    };
    const province = st?.supplyProvince;
    if (typeof province !== "string" || !Object.hasOwn(structure, province)) {
      throw new Error("salesTax.supplyProvince must be one of " + Object.keys(structure).join(", "));
    }
    if (typeof st.taxPointDate !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(st.taxPointDate)) {
      throw new Error("salesTax.taxPointDate must be an ISO 8601 date");
    }
    const date: string = st.taxPointDate.slice(0, 10);
    if (date < "2013-04-01") {
      throw new Error("salesTax.taxPointDate must be on or after 2013-04-01");
    }
    if (!Array.isArray(st.lines)) throw new Error("salesTax.lines must be an array");

    // These are the numbers an owner edits. Older rows (NB/NL 13%, PE 14%,
    // SK 5%, MB 8%) are left out on purpose: a date they would cover throws
    // rather than quietly taxing at today's rate.
    const builtIn = [
      { tax: "GST", rate: 5, from: "2008-01-01" },
      { tax: "HST", province: "ON", rate: 13, from: "2010-07-01" },
      { tax: "HST", province: "NS", rate: 15, from: "2010-07-01", to: "2025-03-31" },
      { tax: "HST", province: "NS", rate: 14, from: "2025-04-01" },
      { tax: "HST", province: "NB", rate: 15, from: "2016-07-01" },
      { tax: "HST", province: "NL", rate: 15, from: "2016-07-01" },
      { tax: "HST", province: "PE", rate: 15, from: "2016-10-01" },
      { tax: "PST", province: "BC", rate: 7, from: "2013-04-01" },
      { tax: "PST", province: "SK", rate: 6, from: "2017-03-23" },
      { tax: "RST", province: "MB", rate: 7, from: "2019-07-01" },
      { tax: "QST", province: "QC", rate: 9.975, from: "2013-01-01" },
    ];
    const rates: Array<{ tax: string; province?: string; rate: number; from: string; to?: string }> =
      Array.isArray(st.rates) && st.rates.length > 0 ? st.rates : builtIn;
    for (const row of rates) {
      if (
        typeof row?.rate !== "number" || !Number.isFinite(row.rate) || row.rate < 0 ||
        Math.abs(row.rate * 1000 - Math.round(row.rate * 1000)) > 1e-6
      ) {
        throw new Error(
          "salesTax.rates: " + row?.tax + " " + (row?.province ?? "") + " rate " + row?.rate + " from " + row?.from +
            " must be a non-negative number with at most three decimal places",
        );
      }
    }
    const pickRate = (tax: string): number => {
      const hits = rates.filter((r) =>
        r.tax === tax && (tax === "GST" || r.province === province) &&
        r.from <= date && (r.to === undefined || date <= r.to)
      );
      const where = tax === "GST" ? "" : " for " + province;
      if (hits.length === 0) {
        throw new Error("No " + tax + " rate" + where + " on " + date + "; add a row to salesTax.rates");
      }
      if (hits.length > 1) {
        throw new Error("Overlapping " + tax + " rows" + where + " on " + date + "; fix salesTax.rates");
      }
      return hits[0].rate;
    };

    const [federalName, provincialName] = structure[province];
    const federalOn = sel?.gstRegistered === true;
    const provincialOn = provincialName !== null && sel?.provincialRegistered === true;
    const federalRate = federalOn ? pickRate(federalName) : 0;
    const provincialRate = provincialOn ? pickRate(provincialName as string) : 0;

    const gstStatuses = ["taxable", "zero-rated", "exempt"];
    const provStatuses = ["taxable", "exempt"];
    let subtotal = 0;
    let federalBase = 0;
    let provincialBase = 0;
    const notFederal: string[] = [];
    const notProvincial: string[] = [];
    let taxableFederal = 0;
    let taxableProvincial = 0;
    let undescribed = false;
    st.lines.forEach((line: { description?: unknown; amount?: unknown; gstCategory?: unknown; provincialCategory?: unknown }, i: number) => {
      const at = "salesTax.lines[" + i + "]";
      if (typeof line?.amount !== "number" || !Number.isFinite(line.amount)) {
        throw new Error(at + ".amount must be a finite number");
      }
      const gst = line.gstCategory ?? "taxable";
      if (!gstStatuses.includes(gst as string)) {
        throw new Error(at + ".gstCategory must be one of " + gstStatuses.join(", "));
      }
      if (line.provincialCategory !== undefined && !provStatuses.includes(line.provincialCategory as string)) {
        throw new Error(at + ".provincialCategory must be one of " + provStatuses.join(", "));
      }
      const prov = line.provincialCategory ?? (gst === "taxable" ? "taxable" : "exempt");
      const cents = Math.round(line.amount * 100);
      const blank = typeof line.description !== "string" || line.description.trim() === "";
      const name = blank ? "line " + (i + 1) : (line.description as string).trim();
      if (blank) undescribed = true;
      subtotal += cents;
      if (gst === "taxable") {
        federalBase += cents;
        taxableFederal++;
      } else notFederal.push(name);
      if (prov === "taxable") {
        provincialBase += cents;
        taxableProvincial++;
      } else notProvincial.push(name);
    });

    // Half away from zero in integers, so a 9.975% half cent cannot float down.
    const taxOn = (base: number, rate: number) =>
      Math.sign(base) * Math.floor((Math.abs(base) * Math.round(rate * 1000) + 50000) / 100000);
    const federalTax = federalOn ? taxOn(federalBase, federalRate) : 0;
    const provincialTax = provincialOn ? taxOn(provincialBase, provincialRate) : 0;
    const total = subtotal + federalTax + provincialTax;

    const currency = new Intl.NumberFormat("en-CA", { style: "currency", currency: "CAD" });
    // + 0 turns a negative zero into zero, which would print as "-$0.00".
    const money = (cents: number) => currency.format(cents / 100 + 0);

    const normalise = (value: unknown) => typeof value === "string" ? value.replace(/[\s-]/g, "").toUpperCase() : "";
    const blank = (value: unknown) => typeof value !== "string" || value.trim() === "";
    const bnRaw = normalise(sel?.bn);
    const bnMatch = /^(\d{9})RT(\d{4})$/.exec(bnRaw);
    let luhn = 0;
    if (bnMatch) {
      [...bnMatch[1]].forEach((ch, i) => {
        let d = Number(ch);
        if (i % 2 === 1) {
          d *= 2;
          if (d > 9) d -= 9;
        }
        luhn += d;
      });
    }
    const bnValid = bnMatch !== null && bnMatch[2] !== "0000" && luhn % 10 === 0;
    const qstRaw = normalise(sel?.qst);
    const qstOn = provincialOn && provincialName === "QST";
    const qstValid = /^\d{10}TQ(\d{4})$/.test(qstRaw) && !qstRaw.endsWith("0000");

    const notes: string[] = [];
    for (
      const [charged, taxName, names, taxable] of [
        [federalOn, federalName, notFederal, taxableFederal],
        [provincialOn, provincialName, notProvincial, taxableProvincial],
      ] as const
    ) {
      if (charged && taxable > 0 && names.length > 0) notes.push("No " + taxName + " on: " + names.join("; ") + ".");
    }

    const abs = Math.abs(total);
    const tier = abs < 10000 ? 0 : abs < 50000 ? 1 : 2;
    const warnings: string[] = [];
    if (federalOn && !blank(sel.bn) && !bnValid) {
      warnings.push(
        sel.bn.trim() +
          " is not a valid GST/HST registration number: it must be a nine-digit business number that passes its check digit, then RT and a four-digit account.",
      );
    }
    if (qstOn && !blank(sel.qst) && !qstValid) {
      warnings.push(
        sel.qst.trim() + " is not a valid QST registration number: it must be ten digits, then TQ and four digits.",
      );
    }
    if (provincialName !== null && sel?.provincialRegistered === undefined) {
      warnings.push(
        "It is not stated whether the seller is registered for " + provincialName + ", so none has been charged.",
      );
    }
    const missing: string[] = [];
    if (blank(sel?.name)) missing.push("seller name");
    if (tier >= 1 && federalOn && blank(sel.bn)) missing.push("GST/HST registration number");
    if (tier >= 1 && qstOn && blank(sel.qst)) missing.push("QST registration number");
    if (tier >= 2) {
      if (blank(cus?.name)) missing.push("customer name");
      if (blank(st.paymentTerms)) missing.push("payment terms");
      if (undescribed) missing.push("a description for every line");
    }
    if (missing.length > 0) {
      warnings.push(
        "Missing for an invoice of " + ["under $100", "$100 or more", "$500 or more"][tier] + ": " +
          missing.join(", ") + ". Until it is added the customer cannot claim input tax credits on this invoice.",
      );
    }

    return {
      subtotal: money(subtotal),
      federalLabel: federalOn ? federalName + " " + federalRate + "% on " + money(federalBase) : null,
      federalAmount: federalOn ? money(federalTax) : null,
      provincialLabel: provincialOn ? provincialName + " " + provincialRate + "% on " + money(provincialBase) : null,
      provincialAmount: provincialOn ? money(provincialTax) : null,
      total: money(total),
      bn: federalOn && bnValid ? bnRaw : null,
      qst: qstOn && qstValid ? qstRaw : null,
      smallSupplierNote: federalOn
        ? null
        : "Not registered for GST/HST (small supplier), so no " + federalName + " is charged.",
      statusNote: notes.length > 0 ? notes.join(" ") : null,
      warning: warnings.length > 0 ? warnings.join(" ") : null,
    };
  },
});
