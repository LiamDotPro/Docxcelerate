import { deriver } from "docxcelerate/template";

/**
 * What a UK construction subcontractor is paid on one invoice, and every line
 * of the working, as strings ready to print.
 *
 * Self-contained on purpose: a deriver is shipped to the engine as source
 * text, so the formatter, the helpers and every constant live inside `run`.
 * Anything outside it is undefined where it runs. It has no placeholder,
 * because it is cheap and previews should show real figures.
 *
 * Money is whole pence from the first line, so no sum can drift. CIS is
 * floored, because HMRC prescribes no rounding and flooring never
 * over-deducts. VAT is rounded to the nearest penny with half going up, in
 * integers: `Math.round(gross * 0.2)` misrounds at .5 in floating point.
 */
export default deriver({
  name: "cisInvoiceDeduction",
  run: (invoice: unknown, supplier: unknown, customer: unknown) => {
    const fail = (message: string): never => {
      throw new Error(`cisInvoiceDeduction: ${message}`);
    };
    const inv = (invoice ?? {}) as Record<string, unknown>;
    const sup = (supplier ?? {}) as Record<string, unknown>;
    const cust = (customer ?? {}) as Record<string, unknown>;
    const money = (field: string, value: unknown, required: boolean): number => {
      if (value == null && !required) return 0;
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
        return fail(`invoice.${field} must be a finite number of pounds, 0 or more; got ${String(value)}.`);
      }

      return Math.round(value * 100);
    };
    const gbp = new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" });
    const pounds = (pence: number) => gbp.format(pence / 100);

    // 1. Validate. Unclear data fails the document: a guessed treatment gives
    //    a disputed figure either way.
    const grossP = money("grossExVat", inv.grossExVat, true);
    const materialsP = money("materialsCost", inv.materialsCost, true);
    const plantP = money("plantHire", inv.plantHire, false);
    const consumablesP = money("consumables", inv.consumables, false);

    if (typeof sup.vatRegistered !== "boolean") fail("supplier.vatRegistered must be true or false.");
    if (sup.cisStatus !== "gross" && sup.cisStatus !== "net" && sup.cisStatus !== "unverified") {
      fail(`supplier.cisStatus must be "gross", "net" or "unverified"; got ${String(sup.cisStatus)}.`);
    }

    const share = inv.reverseChargeShare;

    if (share != null && (typeof share !== "number" || !Number.isFinite(share) || share < 0 || share > 1)) {
      fail(`invoice.reverseChargeShare must be a number from 0 to 1; got ${String(share)}.`);
    }

    const notified = cust.endUserNotified;

    if (notified != null && notified !== "" && notified !== "end-user" && notified !== "intermediary") {
      fail(`customer.endUserNotified must be "end-user", "intermediary" or missing; got ${String(notified)}.`);
    }

    const vatRate = inv.vatRate;
    const registered = sup.vatRegistered === true;

    if (registered && vatRate !== 0.2 && vatRate !== 0.05 && vatRate !== 0) {
      fail(`invoice.vatRate must be 0.2, 0.05 or 0; got ${String(vatRate)}.`);
    }

    if (registered && vatRate !== 0 && (typeof cust.vatRegistered !== "boolean" || typeof cust.cisRegistered !== "boolean")) {
      fail("customer.vatRegistered and customer.cisRegistered must both be true or false when VAT is charged.");
    }

    // 2. The costs together cannot exceed the gross.
    const costsP = materialsP + plantP + consumablesP;

    if (costsP > grossP) fail("materials, plant hire and consumables exceed the gross.");

    // 3. VAT treatment: the first rule that matches wins. Silence from the
    //    customer is not end-user status, so it stays in the reverse charge.
    const treatment = !registered
      ? "not-registered"
      : vatRate === 0
      ? "zero-rated"
      : cust.vatRegistered === true && cust.cisRegistered === true && (notified == null || notified === "") &&
          ((share as number | undefined) ?? 1) > 0.05
      ? "reverse-charge"
      : "standard";

    // 4. Any markup on materials is already inside the gross, so it stays in labour.
    const labourP = grossP - costsP;

    // 5. CIS comes off labour only. Integer operands, so the floor is exact.
    const pct = sup.cisStatus === "gross" ? 0 : sup.cisStatus === "net" ? 20 : 30;
    const deductionP = Math.floor(labourP * pct / 100);

    // 6. VAT is on the whole gross, never on labour, and CIS never reduces it.
    const vatPct = vatRate === 0.05 ? 5 : 20;
    const vatP = Math.floor((grossP * vatPct + 50) / 100);
    const vatChargedP = treatment === "standard" ? vatP : 0;

    let vatLabel = "";
    let vatValue = "";

    if (treatment === "standard") {
      vatLabel = `VAT at ${vatPct}%`;
      vatValue = pounds(vatP);
    } else if (treatment === "zero-rated") {
      vatLabel = "VAT at 0% (zero-rated)";
      vatValue = pounds(0);
    } else if (treatment === "reverse-charge") {
      vatLabel = "VAT (domestic reverse charge)";
      // Never "£0.00": a zero reads as zero-rating, a different treatment.
      vatValue = "Not charged: customer accounts to HMRC";
    }

    return {
      vatTreatment: treatment,
      gross: pounds(grossP),
      materials: pounds(materialsP),
      plantHire: plantP ? pounds(plantP) : "",
      consumables: consumablesP ? pounds(consumablesP) : "",
      labour: pounds(labourP),
      cisLabel: pct === 0
        ? "CIS deduction (gross payment status)"
        : pct === 30
        ? "CIS deduction at 30% (unverified)"
        : "CIS deduction at 20%",
      cisDeduction: pounds(deductionP),
      vatLabel,
      vatValue,
      // 7. Integer pence, so adding VAT before taking CIS off cannot change the result.
      payable: pounds(grossP + vatChargedP - deductionP),
      notionalVat: treatment === "reverse-charge" ? vatP / 100 : null,
      warning: pct === 30
        ? "CIS has been deducted at the higher rate of 30% because HMRC could not verify this subcontractor as registered. If they have registered since, verify them again before paying: the deduction falls to 20%, or to nothing under gross payment status."
        : "",
    };
  },
});
