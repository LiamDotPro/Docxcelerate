// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Decides, line by line, whether a US sale may go out exempt from sales tax,
 * under which of the buyer's certificates, and which rule failed for the rest.
 * Installed by `dxcl add exemption-certificate-check`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "exemptionCertificateCheck",
  run: (raw: unknown) => {
    const ex = raw as any;
    const SLOTS = 16;

    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    // UTC day indexes, so a day does not shift with the engine's time zone.
    const dayIndex = (iso: string) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d) / 86_400_000;
    };
    const dates = new Intl.DateTimeFormat("en-US", { dateStyle: "long", timeZone: "UTC" });
    const fmtDate = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));
    const optDay = (value: unknown, field: string) =>
      value === undefined || value === null ? undefined : isoDay(value, field);
    const stateOf = (value: unknown, field: string): string => {
      const s = typeof value === "string" ? value.trim().toUpperCase() : "";
      if (!/^[A-Z]{2}$/.test(s)) throw new Error(field + " must be a two-letter state code");
      return s;
    };
    const codes = (value: unknown, field: string): string[] => {
      if (!Array.isArray(value)) throw new Error(field + " must be an array");
      return value.map((v, i) => stateOf(v, field + "[" + i + "]"));
    };

    const invoiceDate = isoDay(ex?.invoiceDate, "exemption.invoiceDate");
    const checkedOn = optDay(ex?.checkedOn, "exemption.checkedOn") ?? invoiceDate;
    if (dayIndex(checkedOn) < dayIndex(invoiceDate)) {
      throw new Error("exemption.checkedOn is before exemption.invoiceDate");
    }
    const st = stateOf(ex?.shipToState, "exemption.shipToState");

    const currency = ex?.currency ?? "USD";
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-US", { style: "currency", currency });
    } catch {
      throw new Error("exemption.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmtMoney = (minor: number) => money.format(minor / 10 ** exponent);

    const sstAsOf = isoDay(ex?.sstMembers?.asOf, "exemption.sstMembers.asOf");
    const sst = codes(ex?.sstMembers?.states, "exemption.sstMembers.states");

    let mtc: { asOf: string; states: Array<{ code: string; resaleOnly: boolean }> } | undefined;
    if (ex?.mtcAcceptingStates !== undefined && ex?.mtcAcceptingStates !== null) {
      const m = ex.mtcAcceptingStates;
      const asOf = isoDay(m.asOf, "exemption.mtcAcceptingStates.asOf");
      if (!Array.isArray(m.states)) throw new Error("exemption.mtcAcceptingStates.states must be an array");
      mtc = {
        asOf,
        states: m.states.map((s: any, i: number) => ({
          code: stateOf(s?.code, "exemption.mtcAcceptingStates.states[" + i + "].code"),
          // Wrongly allowing a non-resale exemption is the costly direction.
          resaleOnly: s?.resaleOnly !== false,
        })),
      };
    }

    const LABELS: Record<string, string> = {
      A: "Federal government",
      B: "State or local government",
      C: "Tribal government",
      D: "Foreign diplomat",
      E: "Charitable organization",
      F: "Religious organization",
      G: "Resale",
      H: "Agricultural production",
      I: "Industrial production or manufacturing",
      J: "Direct pay permit",
      K: "Direct mail",
      L: "Other",
      M: "Educational organization",
    };
    const USE: Record<string, string> = { G: "resale", H: "agricultural", I: "industrial", K: "direct-mail" };

    const rawCerts: unknown[] = ex?.certificates === undefined || ex?.certificates === null ? [] : ex.certificates;
    if (!Array.isArray(rawCerts)) throw new Error("exemption.certificates must be an array");
    const certs = rawCerts.map((c: any, i: number) => {
      const f = "exemption.certificates[" + i + "]";
      if (c?.form !== "SST" && c?.form !== "MTC" && c?.form !== "state") {
        throw new Error(f + ".form must be SST, MTC or state");
      }
      if (typeof c.reason !== "string" || !Object.prototype.hasOwnProperty.call(LABELS, c.reason)) {
        throw new Error(f + ".reason must be a letter from A to M");
      }
      const state = c.state === undefined || c.state === null ? undefined : stateOf(c.state, f + ".state");
      const states = c.states === undefined || c.states === null ? undefined : codes(c.states, f + ".states");
      const covers: string[] = c.form === "state" ? (state ? [state] : []) : (states ?? (state ? [state] : []));
      return {
        form: c.form as string,
        reason: c.reason as string,
        number: typeof c.number === "string" && c.number.trim() !== "" ? c.number.trim() : undefined,
        covers,
        formLabel: c.form === "SST"
          ? "Streamlined certificate (F0003)"
          : c.form === "MTC"
          ? "MTC uniform certificate"
          : state + " state certificate",
        expires: optDay(c.expires, f + ".expires"),
        receivedOn: optDay(c.receivedOn, f + ".receivedOn"),
        pending: c.pending === true,
      };
    });

    const rawLines: unknown[] = ex?.lines === undefined || ex?.lines === null ? [] : ex.lines;
    if (!Array.isArray(rawLines)) throw new Error("exemption.lines must be an array");
    if (rawLines.length > SLOTS) {
      throw new Error("exemptionCertificateCheck: " + rawLines.length + " lines; this component prints at most " + SLOTS);
    }
    const items = rawLines.map((l: any, i: number) => {
      const f = "exemption.lines[" + i + "]";
      if (typeof l?.description !== "string" || l.description.trim() === "") {
        throw new Error(f + ".description is required");
      }
      if (!Number.isSafeInteger(l.amountMinor)) {
        throw new Error(f + ".amountMinor must be a safe integer of minor units");
      }
      return {
        description: l.description as string,
        amountMinor: l.amountMinor as number,
        category: typeof l.category === "string" ? l.category : undefined,
      };
    });

    const empty = () => ({ description: "", amount: "", status: "" as "Exempt" | "Taxable" | "", reason: "" });
    if (items.length === 0) {
      return {
        hasLines: false,
        lines: [] as any[],
        rows: Array.from({ length: SLOTS }, empty),
        summary: "",
        basis: "",
        exemptMinor: 0,
        taxableMinor: 0,
      };
    }

    const invoiceDay = dayIndex(invoiceDate);
    const deadlineDay = invoiceDay + 90;
    const deadline = new Date(deadlineDay * 86_400_000).toISOString().slice(0, 10);
    const stIsSst = sst.includes(st);
    const mtcEntry = mtc?.states.find((s) => s.code === st);

    // One certificate against one line: the first failing check stops it, and
    // its step number lets the caller report the furthest one reached.
    const check = (c: (typeof certs)[number], category: string | undefined) => {
      const label = LABELS[c.reason] + " (" + c.reason + ")";
      const fail = (step: number, why: string) => ({ pass: "" as const, step, why });
      if (c.reason === "L") return fail(1, "Other (L) needs a manual review and is never accepted automatically");
      const use = USE[c.reason];
      if (use && category !== use) {
        return fail(
          1,
          label + " covers only lines categorised " + use + "; this line " +
            (category === undefined ? "has no category" : "is " + category),
        );
      }
      if (c.covers.length === 0) return fail(2, "Certificate names no state");
      if (!c.covers.includes(st)) {
        return fail(2, "Certificate covers " + c.covers.join(", ") + ", not the ship-to state " + st);
      }
      if (c.form === "SST" && !stIsSst) {
        return fail(
          3,
          st + " is not a Streamlined Sales Tax full member as of " + fmtDate(sstAsOf) +
            ", so Form F0003 is not accepted there",
        );
      }
      if (c.form === "MTC") {
        if (!mtc) return fail(3, "No list of states accepting the MTC uniform certificate was supplied");
        if (!mtcEntry) return fail(3, st + " does not accept the MTC uniform certificate");
        if (mtcEntry.resaleOnly && c.reason !== "G") {
          return fail(
            3,
            st + " accepts the MTC uniform certificate for resale only, and this one claims " + label,
          );
        }
      }
      if (!c.pending && c.expires !== undefined && dayIndex(c.expires) < invoiceDay) {
        return fail(4, "Certificate expired on " + fmtDate(c.expires));
      }
      if (c.pending) {
        if (!stIsSst) {
          return fail(5, "A pending certificate is honoured only in Streamlined member states, and " + st + " is not one");
        }
        if (dayIndex(checkedOn) > deadlineDay) {
          return fail(
            5,
            "Exemption certificate pending since " + fmtDate(invoiceDate) + " was not received by " + fmtDate(deadline),
          );
        }
        return { pass: "pending" as const, step: 5, why: "" };
      }
      if (c.receivedOn !== undefined && dayIndex(c.receivedOn) > invoiceDay) {
        if (!stIsSst) {
          return fail(
            5,
            "Certificate received on " + fmtDate(c.receivedOn) + ", after the sale; " + st +
              " gives no 90-day window here",
          );
        }
        if (dayIndex(c.receivedOn) > deadlineDay) {
          return fail(
            5,
            "Certificate received on " + fmtDate(c.receivedOn) + ", after the 90-day deadline of " + fmtDate(deadline),
          );
        }
      }
      return { pass: "full" as const, step: 5, why: "" };
    };

    const lines = items.map((l) => {
      const results = certs.map((c) => ({ c, r: check(c, l.category) }));
      const full = results.find((x) => x.r.pass === "full");
      const pend = results.find((x) => x.r.pass === "pending");
      const base: any = { description: l.description, amountMinor: l.amountMinor };
      if (l.category !== undefined) base.category = l.category;
      if (full) {
        const c = full.c;
        return {
          ...base,
          taxable: false,
          pending: false,
          exemptReason: LABELS[c.reason] + " (" + c.reason + "), " + c.formLabel + (c.number ? " " + c.number : "") +
            (c.reason === "J" ? "; the buyer pays the tax to the state directly" : ""),
        };
      }
      if (pend) {
        return {
          ...base,
          taxable: false,
          pending: true,
          exemptReason: "Exemption certificate pending, required by " + fmtDate(deadline),
        };
      }
      let best: (typeof results)[number] | undefined;
      for (const x of results) if (!best || x.r.step > best.r.step) best = x;
      return {
        ...base,
        taxable: true,
        pending: false,
        failedRule: best ? best.r.why : "No exemption certificate on file",
      };
    });

    const rows = Array.from({ length: SLOTS }, (_, i) => {
      const l = lines[i];
      if (!l) return empty();
      return {
        description: l.description as string,
        amount: fmtMoney(l.amountMinor),
        status: (l.taxable ? "Taxable" : "Exempt") as "Exempt" | "Taxable",
        reason: (l.exemptReason ?? l.failedRule) as string,
      };
    });

    const exemptLines = lines.filter((l) => !l.taxable);
    const taxableLines = lines.filter((l) => l.taxable);
    const exemptMinor = exemptLines.reduce((s, l) => s + l.amountMinor, 0);
    const taxableMinor = taxableLines.reduce((s, l) => s + l.amountMinor, 0);
    const pending = exemptLines.filter((l) => l.pending).length;
    const n = lines.length;

    return {
      hasLines: true,
      // Same shape as us-sales-tax-breakdown's salesTax.lines.
      lines: lines.map(({ pending: _p, ...l }) => l) as Array<{
        description: string;
        amountMinor: number;
        taxable: boolean;
        category?: string;
        exemptReason?: string;
        failedRule?: string;
      }>,
      rows,
      summary: "Exempt: " + exemptLines.length + " of " + n + " lines, " + fmtMoney(exemptMinor) +
        (pending > 0 ? " (" + pending + " pending a certificate)" : "") +
        ". Taxable: " + taxableLines.length + " of " + n + " lines, " + fmtMoney(taxableMinor) + ".",
      basis: "Checked on " + fmtDate(checkedOn) + " for delivery to " + st +
        ", against Streamlined Sales Tax full members as of " + fmtDate(sstAsOf) +
        (mtc ? " and states accepting the MTC uniform certificate as of " + fmtDate(mtc.asOf) : "") + ".",
      exemptMinor,
      taxableMinor,
    };
  },
});
