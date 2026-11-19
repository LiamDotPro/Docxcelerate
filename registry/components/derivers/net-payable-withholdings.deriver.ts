import { deriver } from "docxcelerate/template";

/**
 * Settles an invoice total into what the customer transfers. Installed by
 * `dxcl add net-payable-withholdings`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level helper or constant would be a ReferenceError there. Every
 * parser, rate and label lives inside `run`.
 */
export default deriver({
  name: "netPayableWithholdings",
  run: (invoice: unknown, supplier: unknown, customer: unknown) => {
    // deno-lint-ignore no-explicit-any
    const inv = (invoice ?? {}) as any;
    // deno-lint-ignore no-explicit-any
    const sup = (supplier ?? {}) as any;
    // deno-lint-ignore no-explicit-any
    const cus = (customer ?? {}) as any;

    const fail = (field: string, rule: string, got: unknown): never => {
      throw new Error(
        "netPayableWithholdings: " + field + " " + rule + "; got " + (got === undefined ? "undefined" : JSON.stringify(got)) + ".",
      );
    };
    // A decimal with at most two places, to exact integer hundredths (cents, or
    // basis points for a percentage), from the string and never by a float multiply.
    const hundredths = (field: string, value: unknown): number => {
      const s = typeof value === "number" ? String(value) : value;
      if (typeof s !== "string" || !/^\d+(\.\d{1,2})?$/.test(s)) {
        return fail(field, "must be a non-negative decimal with at most two places", value);
      }
      const [whole, frac = ""] = s.split(".");
      return Number(whole) * 100 + Number((frac + "00").slice(0, 2));
    };
    // Half up and exact: every operand is a non-negative integer.
    const pct = (cents: number, bp: number) => Math.floor((cents * bp + 5000) / 10000);
    const given = (v: unknown) => v !== undefined && v !== null && v !== "";

    // Validation, before any figure is computed.
    const country: string = inv.country;
    if (country !== "ES" && country !== "IT") fail("invoice.country", "must be ES or IT", inv.country);
    const types = ["business", "consumer", "public-administration"];
    if (!types.includes(cus.type)) fail("customer.type", "must be one of " + types.join(", "), cus.type);
    const isES = country === "ES";
    const consumer = cus.type === "consumer";

    const B = hundredths("invoice.base", inv.base);
    const E = given(inv.exemptAmount) ? hundredths("invoice.exemptAmount", inv.exemptAmount) : 0;

    const other = (field: string, value: unknown, theirs: "ES" | "IT") => {
      if (country === theirs && (value === true || (typeof value !== "boolean" && given(value)))) {
        fail(field, "applies to " + (theirs === "ES" ? "IT" : "ES") + " only", value);
      }
    };
    other("customer.reRegime", cus.reRegime, "IT");
    other("supplier.newProfessional", sup.newProfessional, "IT");
    other("supplier.forfettario", sup.forfettario, "ES");
    other("supplier.rivalsaInps", sup.rivalsaInps, "ES");
    other("supplier.cassaPercent", sup.cassaPercent, "ES");
    other("invoice.splitPayment", inv.splitPayment, "ES");
    other("invoice.stampRecharged", inv.stampRecharged, "ES");

    const forfettario = sup.forfettario === true;
    const professional = sup.professional === true;
    const vatSet = isES ? ["21", "10", "4", "0"] : ["22", "10", "5", "4", "0"];
    let vat = 0;
    if (inv.vatPercent === undefined || inv.vatPercent === null) {
      if (!forfettario) fail("invoice.vatPercent", "is required", inv.vatPercent);
    } else {
      if (!vatSet.includes(String(inv.vatPercent))) {
        fail("invoice.vatPercent", "must be one of " + vatSet.join(", "), inv.vatPercent);
      }
      vat = Number(inv.vatPercent) * 100;
      if (forfettario && vat !== 0) fail("invoice.vatPercent", "must be missing or 0 for a forfettario supplier", inv.vatPercent);
    }

    if (sup.newProfessional === true && !professional) {
      fail("supplier.newProfessional", "requires supplier.professional", sup.newProfessional);
    }
    if (sup.rivalsaInps === true && given(sup.cassaPercent)) {
      fail("supplier.rivalsaInps", "cannot be combined with supplier.cassaPercent", sup.rivalsaInps);
    }
    let cassaBp = 0;
    if (given(sup.cassaPercent)) {
      cassaBp = hundredths("supplier.cassaPercent", sup.cassaPercent);
      if (cassaBp <= 0 || cassaBp > 1000) fail("supplier.cassaPercent", "must be above 0 and at most 10", sup.cassaPercent);
    }

    let agent = false;
    if (!isES) {
      const ritenutaPossible = professional && !forfettario && !consumer;
      if (ritenutaPossible && typeof cus.withholdingAgent !== "boolean") {
        fail("customer.withholdingAgent", "must be true or false when a ritenuta could apply", cus.withholdingAgent);
      }
      if (consumer && cus.withholdingAgent === true) {
        fail("customer.withholdingAgent", "cannot be true for a consumer", cus.withholdingAgent);
      }
      agent = ritenutaPossible && cus.withholdingAgent === true;
    }
    const withheld = isES ? professional && !consumer : agent;

    const split = inv.splitPayment === true;
    if (split) {
      if (consumer) fail("invoice.splitPayment", "is not available to a consumer", inv.splitPayment);
      if (forfettario || vat === 0) fail("invoice.splitPayment", "needs IVA to be charged", inv.splitPayment);
      if (withheld) fail("invoice.splitPayment", "cannot be combined with a ritenuta (art. 12 DL 87/2018)", inv.splitPayment);
    }

    // Figures, in legal order, each on the rows already rounded above it.
    let rivalsa = 0, cassa = 0, T = B, V: number, re = 0, stamp = 0, stampDue = false;
    if (isES) {
      V = pct(B, vat);
      if (cus.reRegime === true && vat > 0) {
        re = pct(B, vat === 2100 ? 520 : vat === 1000 ? 140 : 50);
      }
    } else {
      if (sup.rivalsaInps === true) rivalsa = pct(B, 400);
      if (cassaBp > 0) cassa = pct(B, cassaBp);
      T = B + rivalsa + cassa;
      V = forfettario ? 0 : pct(T, vat);
      // Strictly over 77.47 EUR carrying no IVA.
      stampDue = E + (V === 0 ? T : 0) > 7747;
      if (stampDue && inv.stampRecharged === true) stamp = 200;
    }
    const total = (isES ? B : T) + V + re + E + stamp;
    let withholding = 0;
    if (withheld) {
      withholding = isES ? pct(B, sup.newProfessional === true ? 700 : 1500) : pct(B + rivalsa, 2000);
    }
    const splitPayment = split ? V : 0;
    const net = total - withholding - splitPayment;

    const nf = new Intl.NumberFormat(isES ? "es-ES" : "it-IT", { style: "currency", currency: "EUR" });
    const fmt = (cents: number) => nf.format(cents / 100);
    const rate = (bp: number) => String(bp / 100).replace(".", ",");

    // These are the strings an owner edits: the legal wording, in the country's
    // language whatever the document locale.
    const text = isES
      ? {
        title: "Total a pagar",
        base: "Base imponible",
        vat: "IVA " + rate(vat) + "% (" + (vat === 2100 ? "art. 90" : vat === 1000 ? "art. 91.Uno" : "art. 91.Dos") + " Ley 37/1992)",
        exempt: "Suplidos (art. 78.Tres.3 Ley 37/1992)",
        total: "Total factura",
        withholding: "Retención IRPF " + (sup.newProfessional === true ? "7" : "15") + "% sobre " + fmt(B) + " (art. 101.5 Ley 35/2006)",
        net: "Total a pagar",
      }
      : {
        title: "Netto a pagare",
        base: professional ? "Compenso" : "Corrispettivo",
        vat: "IVA " + rate(vat) + "% su " + fmt(T) + " (art. 16 DPR 633/1972)",
        exempt: "Spese anticipate in nome e per conto del cliente (art. 15 DPR 633/1972)",
        total: "Totale documento",
        withholding: "Ritenuta d'acconto 20% su " + fmt(B + rivalsa) + " (art. 25 DPR 600/1973)",
        net: "Netto a pagare",
      };
    const rows = (on: boolean, label: string, cents: number) => ({
      label: on ? label : null,
      value: on ? fmt(cents) : null,
    });
    const vatRow = rows(vat > 0 && !forfettario, text.vat, V);
    const reRow = rows(
      isES && cus.reRegime === true && vat > 0,
      "Recargo de equivalencia " + rate(vat === 2100 ? 520 : vat === 1000 ? 140 : 50) + "% (art. 161 Ley 37/1992)",
      re,
    );
    const rivalsaRow = rows(sup.rivalsaInps === true, "Rivalsa INPS 4% (art. 1, c. 212, L. 662/1996)", rivalsa);
    const cassaRow = rows(
      cassaBp > 0,
      "Contributo integrativo " + (given(sup.cassaName) ? sup.cassaName + " " : "") + rate(cassaBp) + "%",
      cassa,
    );
    const exemptRow = rows(E > 0, text.exempt, E);
    const stampRow = rows(stamp > 0, "Imposta di bollo (DPR 642/1972)", stamp);
    const withholdingRow = rows(withheld, text.withholding, 0 - withholding);
    const splitRow = rows(
      split,
      "IVA versata dal committente, scissione dei pagamenti (art. 17-ter DPR 633/1972)",
      0 - splitPayment,
    );

    return {
      title: text.title,
      baseLabel: text.base,
      base: fmt(B),
      rivalsaLabel: rivalsaRow.label,
      rivalsa: rivalsaRow.value,
      cassaLabel: cassaRow.label,
      cassa: cassaRow.value,
      vatLabel: vatRow.label,
      vat: vatRow.value,
      reLabel: reRow.label,
      re: reRow.value,
      exemptLabel: exemptRow.label,
      exempt: exemptRow.value,
      stampLabel: stampRow.label,
      stamp: stampRow.value,
      totalLabel: text.total,
      total: fmt(total),
      withholdingLabel: withholdingRow.label,
      withholding: withholdingRow.value,
      splitPaymentLabel: splitRow.label,
      splitPayment: splitRow.value,
      netLabel: text.net,
      net: fmt(net),
      withholdingNote: !isES && forfettario && professional && !consumer
        ? "Compenso non soggetto a ritenuta d'acconto (art. 1, c. 67, L. 190/2014)."
        : null,
      stampNote: stampDue && stamp === 0
        ? "Imposta di bollo di " + fmt(200) + " assolta dall'emittente (DPR 642/1972)."
        : null,
      cents: {
        base: B,
        rivalsa,
        cassa,
        vat: V,
        re,
        exempt: E,
        stamp,
        total,
        withholding,
        splitPayment,
        net,
      },
    };
  },
});
