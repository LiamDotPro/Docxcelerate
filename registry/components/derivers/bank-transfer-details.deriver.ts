import { deriver } from "docxcelerate/template";

/**
 * Decides which bank details a payer sees, checks them, and fits the payment
 * reference to the rail. Installed by `dxcl add bank-transfer-details`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The country
 * list, the character sets and every helper are declared inside `run`.
 */
export default deriver({
  name: "bankTransferDetails",
  run: (payee: unknown, payer: unknown, invoice: unknown) => {
    // deno-lint-ignore no-explicit-any
    const pay = payee as any;
    // deno-lint-ignore no-explicit-any
    const per = payer as any;
    // deno-lint-ignore no-explicit-any
    const inv = invoice as any;

    if (typeof pay?.accountName !== "string" || pay.accountName.trim() === "") {
      throw new Error("payee.accountName is required");
    }
    if (typeof inv?.reference !== "string" || inv.reference.trim() === "") {
      throw new Error("invoice.reference is required");
    }
    if (typeof inv.currency !== "string" || !/^[A-Za-z]{3}$/.test(inv.currency.trim())) {
      throw new Error("invoice.currency must be a three-letter ISO 4217 code");
    }

    const eea = [
      "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "IE", "IT", "LV",
      "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "IS", "LI", "NO",
    ];
    // Bacs and Faster Payments; the SWIFT basic set is shared by SEPA, MT103 and ACH.
    const ukSet = /^[A-Z0-9 .\-/&]+$/;
    const swiftSet = /^[A-Za-z0-9/\-?:().,'+ ]+$/;
    const limits: Record<string, { max: number; set: RegExp; text: string; upper: boolean }> = {
      uk: { max: 18, set: ukSet, text: "A-Z, 0-9, space and . - / &", upper: true },
      sepa: { max: 140, set: swiftSet, text: "A-Z, a-z, 0-9, space and / - ? : ( ) . , ' +", upper: false },
      international: { max: 140, set: swiftSet, text: "A-Z, a-z, 0-9, space and / - ? : ( ) . , ' +", upper: false },
      // 80 is ACH CCD+ addenda, tighter than Fedwire's 4x35, so one reference fits both.
      us: { max: 80, set: swiftSet, text: "A-Z, a-z, 0-9, space and / - ? : ( ) . , ' +", upper: false },
    };

    let country = String(per?.country ?? "").trim().toUpperCase();
    if (country === "UK") country = "GB";
    const currency = inv.currency.trim().toUpperCase();

    // Spaces and hyphens ignored; null unless all digits and an allowed length.
    const digits = (s: unknown, min: number, max: number): string | null => {
      if (typeof s !== "string") return null;
      const d = s.replace(/[\s-]/g, "");
      return /^\d+$/.test(d) && d.length >= min && d.length <= max ? d : null;
    };
    const ibanOf = (s: unknown): string | null => {
      if (typeof s !== "string") return null;
      const v = s.replace(/\s+/g, "").toUpperCase();
      if (!/^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/.test(v)) return null;
      const moved = v.slice(4) + v.slice(0, 4);
      let rem = 0;
      for (const ch of moved) {
        const code = ch.charCodeAt(0);
        const value = code >= 65 ? String(code - 55) : ch;
        for (const d of value) rem = (rem * 10 + Number(d)) % 97;
      }
      return rem === 1 ? v.replace(/(.{4})(?=.)/g, "$1 ") : null;
    };
    const bicOf = (s: unknown): string | null => {
      if (typeof s !== "string") return null;
      const v = s.replace(/\s+/g, "").toUpperCase();
      return /^[A-Z]{4}[A-Z]{2}[A-Z0-9]{2}([A-Z0-9]{3})?$/.test(v) ? v : null;
    };
    const abaOf = (s: unknown): string | null => {
      const d = digits(s, 9, 9);
      if (d === null) return null;
      const n = d.split("").map(Number);
      const sum = 3 * (n[0] + n[3] + n[6]) + 7 * (n[1] + n[4] + n[7]) + (n[2] + n[5] + n[8]);
      return sum % 10 === 0 ? d : null;
    };

    const sortCode = digits(pay.uk?.sortCode, 6, 6);
    const ukAccount = digits(pay.uk?.accountNumber, 8, 8);
    const iban = ibanOf(pay.iban);
    const bic = bicOf(pay.bic);
    const ach = abaOf(pay.us?.achRouting);
    const usAccount = digits(pay.us?.accountNumber, 4, 17);

    let rail = "none";
    if (country === "GB" && currency === "GBP" && sortCode !== null && ukAccount !== null) rail = "uk";
    else if (eea.includes(country) && currency === "EUR" && iban !== null) rail = "sepa";
    else if (country === "US" && currency === "USD" && ach !== null && usAccount !== null) rail = "us";
    else if (iban !== null && bic !== null) rail = "international";

    const lines = {
      warning: null as string | null,
      method: "We cannot print our bank details on this invoice. Please call us, on a number you already have for us, to arrange payment.",
      accountName: null as string | null,
      sortCode: null as string | null,
      routingLabel: null as string | null,
      routing: null as string | null,
      wireRouting: null as string | null,
      accountNumber: null as string | null,
      iban: null as string | null,
      bic: null as string | null,
      reference: null as string | null,
    };
    if (rail === "none") return lines;

    lines.accountName = pay.accountName.trim();
    if (rail === "uk") {
      lines.method =
        "Pay by bank transfer (Faster Payments or Bacs). Enter the account name exactly as it is shown here: your bank checks it against the account before you pay.";
      lines.sortCode = (sortCode as string).replace(/(\d\d)(\d\d)(\d\d)/, "$1-$2-$3");
      lines.accountNumber = ukAccount;
    } else if (rail === "sepa") {
      lines.method = "Pay by SEPA credit transfer, in euros.";
      lines.iban = iban;
      lines.bic = bic;
    } else if (rail === "us") {
      const wire = abaOf(pay.us.wireRouting);
      const given = pay.us.wireRouting !== undefined && pay.us.wireRouting !== null && pay.us.wireRouting !== "";
      const separate = wire !== null && wire !== ach;
      // Never imply the ACH number works for wires when the caller said it does not.
      lines.method = given && wire === null
        ? "Pay by ACH transfer, in US dollars."
        : "Pay by ACH or wire transfer, in US dollars.";
      lines.routingLabel = given && wire === null || separate ? "ACH routing number" : "Routing number (ACH and wire)";
      lines.routing = ach;
      lines.wireRouting = separate ? wire : null;
      lines.accountNumber = usAccount;
    } else {
      lines.method = "Pay by international (SWIFT) transfer, in " + currency + ".";
      lines.iban = iban;
      lines.bic = bic;
    }

    // Banks cut a long reference short without saying so, so one that does not
    // fit is replaced by the customer reference, never truncated or cleaned.
    const rule = limits[rail];
    const fit = (s: unknown): string | null => {
      if (typeof s !== "string") return null;
      let v = s.trim().replace(/\s+/g, " ");
      if (v === "") return null;
      if (rule.upper) v = v.toUpperCase();
      return v.length <= rule.max && rule.set.test(v) ? v : null;
    };
    const reference = fit(inv.reference) ?? fit(inv.customerReference);
    if (reference === null) {
      throw new Error(
        "invoice.reference does not fit a " + rail + " payment reference (" + rule.max + " characters from " +
          rule.text + "), and no invoice.customerReference that does was given",
      );
    }
    lines.reference = reference;

    const changed = pay.detailsChangedOn;
    if (typeof changed === "string" && changed !== "") {
      const then = /^\d{4}-\d{2}-\d{2}/.test(changed) ? new Date(changed.slice(0, 10) + "T00:00:00Z") : null;
      const tail = " Before you pay, call us on a number you already have for us, not one from this invoice or an email, and check the details below with us.";
      // UTC, so the date does not shift by a day on the engine. A date that
      // does not parse still warns: a malformed date never suppresses it.
      lines.warning = then !== null && !Number.isNaN(then.getTime())
        ? "Our bank details changed on " +
          new Intl.DateTimeFormat(country === "US" ? "en-US" : "en-GB", { dateStyle: "long", timeZone: "UTC" }).format(then) +
          "." + tail
        : "Our bank details have changed since our last invoice." + tail;
    }
    return lines;
  },
});
