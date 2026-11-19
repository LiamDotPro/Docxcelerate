// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Compares each provider's bill with its own line on a good faith estimate and
 * decides what the patient is told. Installed by
 * `dxcl add good-faith-estimate-comparison`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper, formatter and the 40000-cent threshold are declared inside `run`.
 */
export default deriver({
  name: "goodFaithEstimateComparison",
  run: (patientBill: unknown) => {
    const bill = patientBill as any;
    const emptyRow = () => ({ name: "", estimate: "", billed: "", difference: "" });
    const none = () => ({
      outcome: "none" as "none" | "within" | "over" | "dispute",
      estimateDate: "",
      rows: [emptyRow(), emptyRow(), emptyRow(), emptyRow(), emptyRow(), emptyRow()],
      disputeSubject: "",
      initialBillDate: "",
      deadline: "",
      stateName: "",
      stateContact: "",
    });

    if (bill === null || typeof bill !== "object") throw new Error("patientBill is required");
    const coverage = bill.coverage;
    if (coverage !== "uninsured" && coverage !== "self-pay" && coverage !== "insured" && coverage !== "federal-program") {
      throw new Error('patientBill.coverage must be "uninsured", "self-pay", "insured" or "federal-program"');
    }
    if (coverage === "insured" || coverage === "federal-program") return none();
    if (bill.estimate === undefined || bill.estimate === null) return none();
    if (!Array.isArray(bill.providers) || bill.providers.length === 0) return none();

    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    const centsOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
        throw new Error(field + " must be a non-negative integer of cents");
      }
      return value;
    };
    const dates = new Intl.DateTimeFormat("en-US", { dateStyle: "long", timeZone: "UTC" });
    const fmtDate = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));
    const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
    const fmtMoney = (cents: number) => money.format(cents / 100);

    const estimateIso = isoDay(bill.estimate.date, "patientBill.estimate.date");
    const billIso = isoDay(bill.initialBillDate, "patientBill.initialBillDate");
    if (bill.providers.length > 6) {
      throw new Error("patientBill.providers has " + bill.providers.length + " entries; this component prints at most 6");
    }

    let stateName = "";
    let stateContact = "";
    if (bill.stateProcess !== undefined && bill.stateProcess !== null) {
      const sp = bill.stateProcess;
      if (typeof sp.name !== "string" || sp.name.trim() === "") throw new Error("patientBill.stateProcess.name is required");
      if (typeof sp.contact !== "string" || sp.contact.trim() === "") {
        throw new Error("patientBill.stateProcess.contact is required");
      }
      stateName = sp.name.trim();
      stateContact = sp.contact.trim();
    }

    const seen = new Set<string>();
    const results = bill.providers.map((p: any, i: number) => {
      const f = "patientBill.providers[" + i + "]";
      if (typeof p?.name !== "string" || p.name.trim() === "") throw new Error(f + ".name is required");
      const key = p.name.trim();
      if (seen.has(key)) throw new Error(f + ".name is a duplicate of an earlier provider");
      seen.add(key);
      const onEstimate = p.estimateCents !== undefined && p.estimateCents !== null;
      const est = onEstimate ? centsOf(p.estimateCents, f + ".estimateCents") : 0;
      if (!Array.isArray(p.billedLines)) throw new Error(f + ".billedLines must be an array");
      let billed = 0;
      p.billedLines.forEach((l: any, j: number) => {
        billed += centsOf(l?.amountCents, f + ".billedLines[" + j + "].amountCents");
      });
      const diff = billed - est;
      return {
        name: key,
        status: diff >= 40000 ? "dispute" : diff > 0 ? "over" : "within",
        row: {
          name: p.name.trim(),
          estimate: onEstimate ? fmtMoney(est) : "Not on your estimate",
          billed: fmtMoney(billed),
          difference: diff > 0 ? fmtMoney(diff) + " more" : diff < 0 ? fmtMoney(-diff) + " less" : "No difference",
        },
      };
    });

    const disputed = results.filter((r: any) => r.status === "dispute").map((r: any) => r.name as string);
    const outcome = disputed.length > 0 ? "dispute" : results.some((r: any) => r.status === "over") ? "over" : "within";
    const rows = results.map((r: any) => r.row);
    while (rows.length < 6) rows.push(emptyRow());

    const deadlineDate = (() => {
      const [y, m, d] = billIso.split("-").map(Number);
      return new Date(Date.UTC(y, m - 1, d + 120));
    })();

    return {
      outcome: outcome as "none" | "within" | "over" | "dispute",
      estimateDate: fmtDate(estimateIso),
      rows,
      disputeSubject: disputed.length === 0
        ? ""
        : disputed.length === 1
        ? "the bill from " + disputed[0]
        : "the bills from " + new Intl.ListFormat("en-US", { type: "conjunction" }).format(disputed),
      initialBillDate: fmtDate(billIso),
      deadline: outcome === "dispute" ? dates.format(deadlineDate) : "",
      stateName,
      stateContact,
    };
  },
});
