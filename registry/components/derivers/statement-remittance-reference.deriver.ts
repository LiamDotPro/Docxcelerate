// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Nets a statement's credits against its open invoices, picks the one payment
 * reference the payer's rail can carry, and lays out the tear-off advice.
 * Installed by `dxcl add statement-remittance-reference`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The rail
 * table, the character sets, the mod-97 function and the collator are all
 * declared inside `run`.
 */
export default deriver({
  name: "statementRemittanceReference",
  run: (statementInput: unknown, itemsInput: unknown) => {
    const st = statementInput as any;

    // --- the statement -----------------------------------------------------
    const isoDay = (value: unknown, field: string): string => {
      const s = typeof value === "string" ? value.slice(0, 10) : "";
      const d = new Date(s + "T00:00:00Z");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || Number.isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== s) {
        throw new Error(field + " must be an ISO 8601 date (YYYY-MM-DD)");
      }
      return s;
    };

    // Every rail's limits in one place. `kind` says how its reference is built.
    const free = { kind: "free", fold: true, maxLines: 1, sep: " ", lead: false };
    const sepaSet = /[^a-zA-Z0-9/\-?:().,'+ ]/g;
    const achSet = /[^A-Z0-9\-./,&()' ]/g;
    const rails: Record<string, any> = {
      bacs: { ...free, currency: "GBP", set: /[^A-Z0-9.&/\- ]/g, width: 18 },
      fps: { ...free, currency: "GBP", set: /[^A-Z0-9.&/\- ]/g, width: 18 },
      sepa: { ...free, currency: "EUR", fold: false, set: sepaSet, width: 140, sep: "," },
      "sepa-rf": { ...free, kind: "rf", currency: "EUR", set: /[^A-Z0-9]/g, width: 25 },
      "ach-ccd": { ...free, kind: "ach", currency: "USD", set: achSet, width: 80 },
      "ach-ctx": { ...free, kind: "ach", currency: "USD", set: achSet, width: 80 },
      swift: { ...free, currency: null, fold: false, set: sepaSet, width: 35, maxLines: 4, sep: ",", lead: true },
      nz: { ...free, currency: "NZD", set: /[^A-Z0-9\-/.& ]/g, width: 12, maxLines: 2 },
      kid: { kind: "digits", currency: "NOK", label: "KID" },
      "se-ocr": { kind: "digits", currency: "SEK", label: "OCR" },
    };
    const rail: string = st?.rail;
    if (typeof rail !== "string" || !Object.prototype.hasOwnProperty.call(rails, rail)) {
      throw new Error("statement.rail must be one of " + Object.keys(rails).join(", "));
    }
    const cfg = rails[rail];

    const currency = typeof st.currency === "string" ? st.currency.toUpperCase() : "";
    let money: Intl.NumberFormat;
    try {
      // The constructor accepts any well-formed code, so ask Intl which it knows.
      if (!Intl.supportedValuesOf("currency").includes(currency)) throw new Error();
      money = new Intl.NumberFormat(rail.startsWith("ach") ? "en-US" : "en-GB", { style: "currency", currency });
    } catch {
      throw new Error("statement.currency must be an ISO 4217 code");
    }
    if (cfg.currency !== null && currency !== cfg.currency) {
      throw new Error("statement.currency must be " + cfg.currency + " on " + rail);
    }
    const d: number = money.resolvedOptions().maximumFractionDigits as number;
    const fmt = (minor: number) => money.format(minor / 10 ** d);

    const statementDay = isoDay(st.date, "statement.date");
    if (typeof st.customerNumber !== "string" || st.customerNumber.trim() === "") {
      throw new Error("statement.customerNumber is required");
    }
    const accountNumber: string = st.customerNumber.trim();
    const acceptsRf = rail === "sepa-rf" || (["bacs", "fps", "sepa", "swift"].includes(rail) && st.acceptsRf === true);
    const yymmdd = statementDay.slice(2, 4) + statementDay.slice(5, 7) + statementDay.slice(8, 10);

    const locale = rail.startsWith("ach") ? "en-US" : "en-GB";
    const dateFmt = (style: "long" | "medium") => {
      const f = new Intl.DateTimeFormat(locale, { dateStyle: style, timeZone: "UTC" });
      return (day: string) => f.format(new Date(day + "T00:00:00Z"));
    };
    const longDate = dateFmt("long");
    const mediumDate = dateFmt("medium");
    const list = new Intl.ListFormat("en", { type: "conjunction" });
    const collator = new Intl.Collator("en", { numeric: true });

    // --- the items ---------------------------------------------------------
    const raw = itemsInput === undefined || itemsInput === null ? [] : itemsInput;
    if (!Array.isArray(raw)) throw new Error("items must be an array");
    const types = ["invoice", "credit-note", "payment"];
    const seen = new Set<string>();
    const parsed = raw.map((it: any, i: number) => {
      const at = "items[" + i + "]";
      if (!types.includes(it?.type)) throw new Error(at + ".type must be one of " + types.join(", "));
      if (typeof it.number !== "string" || it.number.trim() === "") throw new Error(at + ".number is required");
      const key = it.type + "\u0000" + it.number.trim();
      if (seen.has(key)) throw new Error(at + ".number duplicates another " + it.type + ", " + it.number.trim());
      seen.add(key);
      const date = isoDay(it.date, at + ".date");
      // A credit has no due date; it is ordered by its own date.
      const hasDue = it.type === "invoice" && it.dueDate !== undefined && it.dueDate !== null && it.dueDate !== "";
      const due = hasDue ? isoDay(it.dueDate, at + ".dueDate") : date;
      if (typeof it.amount !== "number" || !Number.isFinite(it.amount)) {
        throw new Error(at + ".amount must be a finite number");
      }
      if (it.amount < 0) throw new Error(at + ".amount must be zero or more");
      // toPrecision first, so 1.005 is 100.5 minor units and rounds to 101.
      const minor = Math.round(Number((it.amount * 10 ** d).toPrecision(15)));
      const reference = typeof it.reference === "string" && it.reference.trim() !== "" ? it.reference : null;
      return { i, type: it.type as string, number: it.number as string, date, due, minor, reference, disputed: it.disputed === true };
    }).filter((it) => it.minor !== 0);

    // --- netting -----------------------------------------------------------
    const order = (a: any, b: any) =>
      a.due < b.due ? -1 : a.due > b.due ? 1 : a.date < b.date ? -1 : a.date > b.date ? 1 : collator.compare(a.number, b.number);
    const disputed = parsed.filter((it) => it.disputed).sort(order);
    const credits = parsed.filter((it) => !it.disputed && it.type !== "invoice").sort(order);
    const invoices = parsed.filter((it) => !it.disputed && it.type === "invoice").sort(order);

    let pool = credits.reduce((sum, c) => sum + c.minor, 0);
    const cleared: any[] = [];
    const toPay: any[] = [];
    for (const inv of invoices) {
      const used = Math.min(pool, inv.minor);
      pool -= used;
      const remaining = inv.minor - used;
      if (remaining === 0) cleared.push(inv);
      else toPay.push({ ...inv, remaining, used });
    }
    const net = toPay.reduce((sum, inv) => sum + inv.remaining, 0) - pool;
    const outcome = net > 0 ? "pay" : net === 0 ? "clear" : "credit";
    const invoiceTotal = invoices.reduce((sum, inv) => sum + inv.minor, 0);
    const creditTotal = credits.reduce((sum, c) => sum + c.minor, 0);
    if (invoiceTotal - creditTotal !== net) throw new Error("netting does not balance");

    // --- notes -------------------------------------------------------------
    let creditsNote: string | null = null;
    const part = toPay.find((p) => p.used > 0);
    if (cleared.length > 0 || part) {
      const many = credits.length > 1;
      const named = list.format(
        credits.map((c) => (c.type === "payment" ? "payment " : "credit note ") + c.number + " (" + fmt(c.minor) + ")"),
      );
      const clauses: string[] = [];
      if (cleared.length > 0) clauses.push((many ? "They clear " : "It clears ") + list.format(cleared.map((c) => c.number)));
      if (part) clauses.push((cleared.length > 0 ? (many ? "leave " : "leaves ") : (many ? "They leave " : "It leaves ")) + fmt(part.remaining) + " to pay on " + part.number);
      creditsNote = named.charAt(0).toUpperCase() + named.slice(1) + (many ? " are" : " is") +
        " set against your oldest invoices first. " + clauses.join(", and ") + ".";
    }
    let disputedNote: string | null = null;
    if (disputed.length > 0) {
      const many = disputed.length > 1;
      disputedNote = list.format(disputed.map((x) => x.number + " (" + fmt(x.minor) + ")")) +
        (many ? " are in dispute and are not" : " is in dispute and is not") + " included in the amount to pay.";
    }

    const result = {
      outcome: outcome as "pay" | "clear" | "credit",
      net: outcome === "clear" ? "" : fmt(Math.abs(net)),
      accountNumber,
      statementDate: longDate(statementDay),
      referenceRule: null as "item" | "list" | "rf" | "account" | "per-item" | null,
      referenceLabel: null as string | null,
      referenceLines: [] as string[],
      perItemNote: null as string | null,
      creditsNote,
      disputedNote,
      rows: [] as Array<{ item: string; due: string; amount: string; running: string }>,
      tickInstruction: false,
    };
    if (outcome !== "pay") return result;

    // --- the reference -----------------------------------------------------
    const fold = (text: string): string => {
      let s = text.normalize("NFD").replace(/[̀-ͯ]/g, "");
      if (cfg.fold) s = s.toUpperCase();
      return s.replace(cfg.set, " ").replace(/ +/g, " ").trim();
    };
    const mod97 = (s: string): number => {
      let r = 0;
      for (const ch of s) {
        const digits = /[A-Z]/.test(ch) ? String(ch.charCodeAt(0) - 55) : ch;
        for (const c of digits) r = (r * 10 + Number(c)) % 97;
      }
      return r;
    };
    const validRf = (ref: string) => /^RF\d{2}[A-Z0-9]{1,21}$/.test(ref) && mod97(ref.slice(4) + ref.slice(0, 4)) === 1;

    // Fits already-folded tokens into the rail's lines. Breaks only between
    // tokens, so a number is never split; a comma stays on the line it ended.
    // Greedy; a token fits only if its trailing comma will too (when more follow).
    const place = (tokens: string[]): string[] | null => {
      if (tokens.length === 0 || tokens.some((t) => t === "")) return null;
      const lines: string[] = [];
      let line = tokens[0];
      for (const [k, t] of tokens.slice(1).entries()) {
        const comma = cfg.sep === "," && k < tokens.length - 2 ? 1 : 0;
        if ((line + cfg.sep + t).length + comma <= cfg.width) {
          line += cfg.sep + t;
        } else {
          lines.push(cfg.sep === "," ? line + "," : line);
          line = t;
        }
      }
      lines.push(line);
      if (lines.length > cfg.maxLines || lines.some((l) => l.length > cfg.width || (cfg.lead && /^[:-]/.test(l)))) return null;
      return lines;
    };

    const label: Record<string, string> = {
      bacs: "payment reference",
      fps: "payment reference",
      sepa: "remittance information",
      "sepa-rf": "structured creditor reference (RF)",
      "ach-ccd": "CCD+ addenda",
      "ach-ctx": "CTX addenda records, one per line",
      swift: "remittance information (MT103 field 70)",
      nz: "particulars, code and reference",
      kid: "KID",
      "se-ocr": "OCR reference",
    };
    const tickRows = (extra: (x: any) => string) => {
      let running = 0;
      const rows = toPay.map((x) => {
        running += x.remaining;
        const part = x.used > 0
          ? " (part-paid by credit: " + fmt(x.remaining) + " of " + fmt(x.minor) + " remains)"
          : "";
        return { item: x.number + extra(x) + part, due: mediumDate(x.due), amount: fmt(x.remaining), running: fmt(running) };
      });
      if (running !== net) throw new Error("netting does not balance");
      return rows;
    };

    let rule: "item" | "list" | "rf" | "account" | "per-item";
    let lines: string[] = [];
    let perItem = false;

    if (cfg.kind === "digits") {
      const refs = new Map<number, string>();
      for (const x of toPay) {
        const digits = (x.reference ?? "").replace(/ /g, "");
        if (!/^\d{2,25}$/.test(digits)) {
          throw new Error("items[" + x.i + "].reference must be 2 to 25 digits on " + rail + ", for " + x.number);
        }
        refs.set(x.i, digits);
      }
      perItem = toPay.length > 1;
      rule = perItem ? "per-item" : "item";
      lines = perItem ? [] : [refs.get(toPay[0].i) as string];
      result.rows = tickRows((x) => ", " + cfg.label + " " + refs.get(x.i));
      if (perItem) {
        result.perItemNote = cfg.label + " payments carry one reference each, so please make a separate payment for each item " +
          "in the advice below, quoting the " + cfg.label + " beside it.";
      }
    } else {
      const custFolded = fold(accountNumber);
      if (rail === "nz" && (custFolded === "" || custFolded.length > 12)) {
        throw new Error("statement.customerNumber does not fit the Particulars field on nz (12 characters)");
      }
      const segment = (x: any) => {
        const n = fold(x.number);
        return n === "" ? null : "RMR*IV*" + n + "*PI*" + (x.remaining / 10 ** d).toFixed(d) + "\\";
      };
      const tried: string[] = [];
      const itemStep = (): string[] | null => {
        const x = toPay[0];
        if (cfg.kind === "rf") {
          const ref = (x.reference ?? "").replace(/ /g, "").toUpperCase();
          return validRf(ref) ? [ref] : null;
        }
        if (cfg.kind === "ach") {
          const s = segment(x);
          return s !== null && s.length <= 80 ? [s] : null;
        }
        return place([fold(x.reference ?? x.number)]);
      };
      const listStep = (): string[] | null => {
        if (cfg.kind === "ach") {
          const segs = toPay.map(segment);
          if (segs.some((s) => s === null)) return null;
          const all = segs as string[];
          if (rail === "ach-ctx") return all.every((s) => s.length <= 80) ? all : null;
          return all.join("").length <= 80 ? [all.join("")] : null;
        }
        if (cfg.kind === "rf") return null;
        const nums = toPay.map((x) => fold(x.number));
        if (nums.some((n) => n === "") || new Set(nums).size !== nums.length) return null;
        let prefix = nums.reduce((p, n) => {
          let k = 0;
          while (k < p.length && k < n.length && p[k] === n[k]) k++;
          return p.slice(0, k);
        });
        while (prefix !== "" && /\d$/.test(prefix)) prefix = prefix.slice(0, -1);
        if (nums.some((n) => n.length === prefix.length)) prefix = "";
        const rest = nums.map((n) => n.slice(prefix.length));
        if (rest.some((r) => r.includes(cfg.sep))) return null;
        return place([prefix + rest[0], ...rest.slice(1)]);
      };
      const rfStep = (): string[] | null => {
        if (!acceptsRf) return null;
        const cust = accountNumber.normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z0-9]/g, "");
        let base = cust + yymmdd;
        if (base.length > 21) base = cust;
        if (base === "" || base.length > 21) return null;
        const ref = "RF" + String(98 - mod97(base + "RF00")).padStart(2, "0") + base;
        return place([ref]);
      };
      const accountStep = (): string[] | null => {
        if (cfg.kind === "rf") return null;
        const candidates = rail === "nz"
          ? ["STMT " + yymmdd]
          : custFolded === ""
          ? []
          : [custFolded + " STMT " + yymmdd, custFolded + " " + yymmdd, custFolded];
        for (const c of candidates) {
          const placed = place([c]);
          if (placed) return placed;
        }
        return null;
      };
      const steps: Array<[typeof rule, () => string[] | null]> = [
        toPay.length === 1 ? ["item", itemStep] : ["list", listStep],
        ["rf", rfStep],
        ["account", accountStep],
      ];
      let found: string[] | null = null;
      let won: typeof rule = "account";
      for (const [name, step] of steps) {
        tried.push(name);
        found = step();
        if (found) {
          won = name;
          break;
        }
      }
      if (!found) {
        throw new Error("No reference for this statement fits " + rail + ": tried " + tried.join(", ") + ", and nothing fits without truncating");
      }
      rule = won;
      lines = rail === "nz"
        ? ["Particulars: " + custFolded, ...(found.length === 2 ? ["Code: " + found[0]] : []), "Reference: " + found[found.length - 1]]
        : found;
      result.rows = tickRows(() => "");
    }

    result.referenceRule = rule;
    result.referenceLabel = perItem ? null : label[rail];
    result.referenceLines = lines;
    result.rows = [
      ...result.rows,
      ...disputed.map((x) => ({
        item: x.number + " (in dispute, not included)",
        due: mediumDate(x.due),
        amount: fmt(x.minor),
        running: "",
      })),
    ];
    result.tickInstruction = result.rows.length >= 2;
    return result;
  },
});
