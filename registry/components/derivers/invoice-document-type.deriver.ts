import { deriver } from "docxcelerate/template";

/**
 * Decides an invoice's heading, statement and dates. Installed by
 * `dxcl add invoice-document-type`.
 *
 * It is self-contained on purpose:
 * it travels to the engine as source text, so a module-level constant or helper
 * would be a ReferenceError there. Thresholds are declared inside it.
 */
export default deriver({
  name: "invoiceDocumentType",
  run: (invoice: unknown, seller: unknown) => {
    // deno-lint-ignore no-explicit-any
    const inv = invoice as any;
    // deno-lint-ignore no-explicit-any
    const sel = seller as any;

    if (typeof inv?.number !== "string" || inv.number.trim() === "") {
      throw new Error("invoice.number is required");
    }
    if (typeof inv.issueDate !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(inv.issueDate)) {
      throw new Error("invoice.issueDate must be an ISO 8601 date");
    }
    const kinds = ["invoice", "proforma", "self-billed"];
    if (!kinds.includes(inv.kind)) {
      throw new Error("invoice.kind must be one of " + kinds.join(", "));
    }
    const jurisdictions = ["UK", "EU", "DE", "FR", "US"];
    if (!jurisdictions.includes(inv.jurisdiction)) {
      throw new Error("invoice.jurisdiction must be one of " + jurisdictions.join(", "));
    }

    const kind: string = inv.kind;
    const jurisdiction: string = inv.jurisdiction;
    const vatApplies = jurisdiction !== "US" && sel?.vatRegistered === true;
    const vatNumber: string = typeof sel?.vatNumber === "string" ? sel.vatNumber.trim() : "";
    if (vatApplies && vatNumber === "") {
      throw new Error("seller.vatNumber is required when seller.vatRegistered is true");
    }

    // VAT-inclusive ceilings for a simplified invoice. These are the numbers an
    // owner edits. UK: HMRC, 250 GBP. DE: UStDV section 33 Kleinbetragsrechnung,
    // 250 EUR gross. EU and FR: Directive 2006/112/EC Art 220a (content in
    // Art 226b), 100 EUR; France may differ, check it.
    const limit = jurisdiction === "UK" || jurisdiction === "DE" ? 250 : 100;
    const gross = inv.grossTotal;
    // Rounded to 2 dp so the decision agrees with the printed total.
    const simplified = kind === "invoice" && vatApplies && sel.simplifiedAllowed === true &&
      typeof gross === "number" && Number.isFinite(gross) &&
      Math.round(gross * 100) / 100 <= limit;

    const heading = kind === "proforma"
      ? "Pro forma invoice"
      : kind === "self-billed"
      ? "Self-billed invoice"
      : !vatApplies
      ? "Invoice"
      : simplified
      ? "Simplified VAT invoice"
      : "VAT invoice";

    let statement: string | null = null;
    if (vatApplies && kind === "proforma") {
      statement = "This is not a VAT invoice.";
    } else if (vatApplies && kind === "self-billed") {
      const statements: Record<string, string> = {
        UK: "The VAT shown is your output tax due to HMRC.",
        // UStG section 14(4) no. 10 requires the word "Gutschrift".
        DE: "Gutschrift (self-billing).",
        // CGI requires "Autofacturation".
        FR: "Autofacturation (self-billing).",
        // Directive Art 226(10a).
        EU: "Self-billing.",
      };
      statement = statements[jurisdiction] ?? null;
    }

    const day = (value: unknown, field: string): string | null => {
      if (value === undefined || value === null || value === "") return null;
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    // UTC, so the date does not shift by a day on the engine.
    const formatter = new Intl.DateTimeFormat(jurisdiction === "US" ? "en-US" : "en-GB", {
      dateStyle: "long",
      timeZone: "UTC",
    });
    const fmt = (d: string) => formatter.format(new Date(d + "T00:00:00Z"));

    const issueDay = day(inv.issueDate, "invoice.issueDate") as string;
    const de = jurisdiction === "DE";
    // Label and value apart, for an arrangement that sets them in two columns;
    // the line is the two joined, for one that runs them down the page.
    let supplyLabel: string | null = null;
    let supplyValue: string | null = null;
    let supplyLine: string | null = null;

    if (inv.supplyPeriod !== undefined && inv.supplyPeriod !== null) {
      const from = day(inv.supplyPeriod.from, "invoice.supplyPeriod.from");
      const to = day(inv.supplyPeriod.to, "invoice.supplyPeriod.to");
      if (from === null) throw new Error("invoice.supplyPeriod.from is required");
      if (to === null) throw new Error("invoice.supplyPeriod.to is required");
      if (to < from) throw new Error("invoice.supplyPeriod.to is before invoice.supplyPeriod.from");
      supplyLabel = de ? "Leistungszeitraum" : "Supply period";
      supplyValue = fmt(from) + (de ? " bis " : " to ") + fmt(to);
    } else {
      const supplyDay = day(inv.supplyDate, "invoice.supplyDate");
      if (supplyDay !== null && supplyDay !== issueDay) {
        supplyLabel = de ? "Leistungsdatum" : "Date of supply";
        supplyValue = fmt(supplyDay);
      } else if (de) {
        supplyLabel = "Leistungsdatum";
        supplyValue = "entspricht Rechnungsdatum";
        supplyLine = "Leistungsdatum entspricht Rechnungsdatum.";
      }
    }
    if (supplyLine === null && supplyLabel !== null) {
      supplyLine = supplyLabel + ": " + supplyValue;
    }

    return {
      heading,
      statement,
      issueDate: fmt(issueDay),
      supplyLine,
      supplyLabel,
      supplyValue,
      vatNumber: vatApplies ? vatNumber : null,
    };
  },
});
