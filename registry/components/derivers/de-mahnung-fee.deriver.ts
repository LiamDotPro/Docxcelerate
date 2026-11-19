// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Decides, for each invoice of a German Mahnung, since when the debtor is in
 * Verzug, why, and what this letter may charge. Installed by
 * `dxcl add de-mahnung-fee`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper, the Easter computation and both ceilings are declared inside `run`.
 */
export default deriver({
  name: "deMahnungFee",
  run: (input: unknown) => {
    const mahnung = input as any;
    type Route = "contract" | "event" | "invoice-30" | "goods-30" | "earlier-mahnung" | "this-mahnung";

    // § 288 Abs. 5 BGB, in cents.
    const PAUSCHALE = 4000;
    // The component's own conservative bound for "modest" postage and print
    // under § 309 Nr. 5 a) BGB. It is not a figure in the statute.
    const CONSUMER_FEE_CEILING = 250;

    const dates = new Intl.DateTimeFormat("de-DE", { day: "2-digit", month: "2-digit", year: "numeric", timeZone: "UTC" });
    const money = new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR" });
    const fmtDay = (day: number) => dates.format(new Date(day * 86_400_000));
    const fmtMoney = (cents: number) => money.format(cents / 100);
    const isoOf = (day: number) => new Date(day * 86_400_000).toISOString().slice(0, 10);

    // UTC day indexes, so a day does not shift with the engine's time zone.
    const isoDay = (value: unknown, field: string): string => {
      const text = typeof value === "string" ? value.slice(0, 10) : "";
      const d = new Date(text + "T00:00:00Z");
      // getTime() first: toISOString throws a bare RangeError on "2026-13-01".
      const ok = /^\d{4}-\d{2}-\d{2}$/.test(text) && !Number.isNaN(d.getTime()) && d.toISOString().slice(0, 10) === text;
      if (!ok) throw new Error(field + " must be an ISO 8601 date");
      return text;
    };
    const dayIndex = (iso: string) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d) / 86_400_000;
    };
    // Through toPrecision so 2.005 is the 200.5 it was written as, not the
    // 200.49999999999997 a float multiplication leaves.
    const centsOf = (value: unknown, field: string, positive = false): number => {
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0 || (positive && value === 0)) {
        throw new Error(field + (positive ? " must be a finite number above 0" : " must be a finite number of 0 or more"));
      }
      return Math.round(Number((value * 100).toPrecision(15)));
    };

    // ---- validation, before any arithmetic ----
    if (mahnung?.debtorType !== "consumer" && mahnung?.debtorType !== "business") {
      throw new Error('mahnung.debtorType must be "consumer" or "business"');
    }
    const business = mahnung.debtorType === "business";
    const dateIso = isoDay(mahnung.date, "mahnung.date");
    const today = dayIndex(dateIso);
    const feeCents = mahnung.fee === undefined || mahnung.fee === null ? 0 : centsOf(mahnung.fee, "mahnung.fee");
    if (!business && feeCents > CONSUMER_FEE_CEILING) {
      throw new Error("mahnung.fee must not exceed " + fmtMoney(CONSUMER_FEE_CEILING) + " for a consumer");
    }
    const rawHolidays: unknown[] = mahnung.holidays === undefined || mahnung.holidays === null ? [] : mahnung.holidays;
    if (!Array.isArray(rawHolidays)) throw new Error("mahnung.holidays must be an array of ISO 8601 dates");
    const extra = new Set(rawHolidays.map((h, i) => dayIndex(isoDay(h, "mahnung.holidays[" + i + "]"))));
    if (!Array.isArray(mahnung.invoices) || mahnung.invoices.length === 0) {
      throw new Error("mahnung.invoices must be a non-empty array");
    }

    // ---- § 193 BGB ----
    const nationwide = new Map<number, Set<number>>();
    const holidaysOf = (year: number) => {
      let set = nationwide.get(year);
      if (!set) {
        // Anonymous Gregorian algorithm (Meeus/Jones/Butcher).
        const a = year % 19, b = Math.floor(year / 100), c = year % 100;
        const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
        const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
        const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
        const m = Math.floor((a + 11 * h + 22 * l) / 451);
        const month = Math.floor((h + l - 7 * m + 114) / 31);
        const easter = Date.UTC(year, month - 1, ((h + l - 7 * m + 114) % 31) + 1) / 86_400_000;
        const fixed = (mo: number, dd: number) => Date.UTC(year, mo - 1, dd) / 86_400_000;
        set = new Set([
          fixed(1, 1), easter - 2, easter + 1, fixed(5, 1), easter + 39, easter + 50,
          fixed(10, 3), fixed(12, 25), fixed(12, 26),
        ]);
        nationwide.set(year, set);
      }
      return set;
    };
    const nonWorking = (day: number) => {
      const when = new Date(day * 86_400_000);
      const weekday = when.getUTCDay();
      return weekday === 0 || weekday === 6 || extra.has(day) || holidaysOf(when.getUTCFullYear()).has(day);
    };
    const shift = (day: number) => {
      while (nonWorking(day)) day += 1;
      return day;
    };

    // ---- per invoice ----
    const parsed = (mahnung.invoices as any[]).map((inv, i) => {
      const at = "mahnung.invoices[" + i + "]";
      if (typeof inv?.number !== "string" || inv.number.trim() === "") throw new Error(at + ".number is required");
      const amountCents = centsOf(inv.amount, at + ".amount", true);
      const dueIso = isoDay(inv.dueDate, at + ".dueDate");
      const optional = (value: unknown, field: string) => value === undefined || value === null ? undefined : isoDay(value, at + "." + field);
      const received = optional(inv.invoiceReceivedDate, "invoiceReceivedDate");
      const goods = optional(inv.goodsReceivedDate, "goodsReceivedDate");
      const contract = optional(inv.contractPaymentDate, "contractPaymentDate");
      if (contract !== undefined && contract < dueIso) {
        throw new Error(at + ".contractPaymentDate must not be before " + at + ".dueDate");
      }
      const due = shift(dayIndex(dueIso));
      if (today <= due) {
        throw new Error("mahnung.date must be after " + at + ".dueDate, which is due through " + fmtDay(due));
      }
      const earlier: Array<{ day: number; fee: number }> = [];
      const rawEarlier: unknown[] = inv.mahnungen === undefined || inv.mahnungen === null ? [] : inv.mahnungen;
      if (!Array.isArray(rawEarlier)) throw new Error(at + ".mahnungen must be an array");
      rawEarlier.forEach((m: any, j) => {
        const field = at + ".mahnungen[" + j + "]";
        const day = dayIndex(isoDay(m?.date, field + ".date"));
        if (day >= today) throw new Error(field + ".date must be before mahnung.date");
        earlier.push({ day, fee: m.feeCharged === undefined || m.feeCharged === null ? 0 : centsOf(m.feeCharged, field + ".feeCharged") });
      });

      // Each candidate is the first day of Verzug by that route; ties go to the
      // lower rank, so statutory Verzug beats a Mahnung on the same day.
      const candidates: Array<{ day: number; rank: number; route: Route; original?: number; last?: number }> = [];
      if (contract !== undefined) {
        const original = dayIndex(contract);
        const last = shift(original);
        candidates.push({ day: last + 1, rank: 0, route: inv.contractPaymentFromEvent === true ? "event" : "contract", original, last });
      }
      const thirty = (basis: string, route: Route) => {
        const start = Math.max(dayIndex(dueIso), dayIndex(basis)); // § 187 Abs. 1: the start day is not counted
        const last = shift(start + 30);
        candidates.push({ day: last + 1, rank: 1, route, original: start, last });
      };
      if (received !== undefined) {
        if (business || inv.verzugWarningOnInvoice === true) thirty(received, "invoice-30");
      } else if (business && goods !== undefined) thirty(goods, "goods-30");
      const firstEarlier = earlier.filter((m) => m.day > due).sort((x, y) => x.day - y.day)[0];
      if (firstEarlier) candidates.push({ day: firstEarlier.day, rank: 2, route: "earlier-mahnung" });
      candidates.push({ day: today, rank: 3, route: "this-mahnung" });
      candidates.sort((x, y) => x.day - y.day || x.rank - y.rank);

      return {
        number: inv.number as string,
        amountCents,
        dueIso,
        winner: candidates[0],
        before: earlier.reduce((sum, m) => sum + m.fee, 0),
      };
    });

    // Print order: oldest due date first, input order on a tie (sort is stable).
    const ordered = [...parsed].sort((x, y) => (x.dueIso < y.dueIso ? -1 : x.dueIso > y.dueIso ? 1 : 0));

    // The letter's one fee goes to the chargeable invoice in Verzug longest.
    let carrier: (typeof ordered)[number] | undefined;
    if (feeCents > 0) {
      for (const inv of ordered) {
        if (inv.winner.route === "this-mahnung") continue;
        if (!carrier || inv.winner.day < carrier.winner.day) carrier = inv;
      }
    }

    const invoices = ordered.map((inv) => {
      const { winner, before } = inv;
      const route = winner.route;
      const since = fmtDay(winner.day);
      const now = inv === carrier ? feeCents : 0;
      const after = before + now;
      const excessAfter = Math.max(0, after - PAUSCHALE);
      const feeNowMinor = business ? excessAfter - Math.max(0, before - PAUSCHALE) : now;
      const costsMinor = business ? PAUSCHALE + excessAfter : after;

      let verzugLine: string;
      if (route === "contract" || route === "event") {
        const term = route === "event" ? "Der aus dem vereinbarten Ereignis berechnete Zahlungstermin " : "Der vertraglich bestimmte Zahlungstermin ";
        const moved = winner.last !== winner.original
          ? "; da er auf einen Samstag, Sonntag oder Feiertag fiel, lief er nach § 193 BGB bis " + fmtDay(winner.last as number)
          : "";
        verzugLine = "In Verzug seit " + since + ": " + term + fmtDay(winner.original as number) + " ist verstrichen" + moved +
          " (§ 286 Abs. 2 Nr. " + (route === "event" ? "2" : "1") + " BGB).";
      } else if (route === "invoice-30" || route === "goods-30") {
        const moved = winner.last !== (winner.original as number) + 30
          ? "; das Fristende fiel auf einen Samstag, Sonntag oder Feiertag und verschob sich nach § 193 BGB auf " + fmtDay(winner.last as number)
          : "";
        verzugLine = "In Verzug seit " + since + ": 30 Tage nach Fälligkeit und " +
          (route === "invoice-30" ? "Zugang der Rechnung am " : "Empfang der Leistung am ") + fmtDay(winner.original as number) + moved +
          (route === "invoice-30" ? " (§ 286 Abs. 3 BGB)." : " (§ 286 Abs. 3 Satz 2 BGB).");
      } else if (route === "earlier-mahnung") {
        verzugLine = "In Verzug seit " + since + ": durch unsere Mahnung vom " + since + " (§ 286 Abs. 1 BGB).";
      } else {
        verzugLine = "Mit Zugang dieses Schreibens geraten Sie in Verzug (§ 286 Abs. 1 BGB).";
      }

      let feeLine: string;
      if (route === "this-mahnung") {
        feeLine = "Für dieses Schreiben berechnen wir keine Mahngebühr, weil erst dieses Schreiben den Verzug begründet.";
      } else if (carrier && inv !== carrier) {
        feeLine = "Die Kosten dieses Schreibens sind bei Rechnung " + carrier.number + " berechnet.";
      } else if (now === 0) {
        feeLine = "Für dieses Schreiben berechnen wir keine Mahngebühr.";
      } else {
        feeLine = "Mahngebühr für dieses Schreiben: " + fmtMoney(now) + " (Porto und Druck).";
        if (business) {
          if (after <= PAUSCHALE) feeLine += " Sie wird auf die Verzugspauschale angerechnet und nicht zusätzlich berechnet.";
          else if (before < PAUSCHALE) feeLine += " Davon berechnen wir " + fmtMoney(feeNowMinor) + " zusätzlich; der Rest ist durch die Verzugspauschale abgegolten.";
          else feeLine += " Sie wird zusätzlich berechnet, weil die Mahngebühren die Verzugspauschale bereits übersteigen.";
        }
      }

      let pauschaleLine = "";
      if (business) {
        pauschaleLine = route === "this-mahnung"
          ? "Mit Eintritt des Verzugs entsteht die Verzugspauschale nach § 288 Abs. 5 BGB von " + fmtMoney(PAUSCHALE) + "."
          : "Verzugspauschale nach § 288 Abs. 5 BGB: " + fmtMoney(PAUSCHALE) + ".";
        if (after > 0) {
          pauschaleLine += " Mahngebühren von insgesamt " + fmtMoney(after) + " sind darauf angerechnet (§ 288 Abs. 5 Satz 3 BGB)";
          pauschaleLine += after > PAUSCHALE ? "; berechnet wird nur der übersteigende Betrag von " + fmtMoney(after - PAUSCHALE) + "." : ".";
        }
      }

      return {
        number: inv.number,
        heading: "Rechnung " + inv.number + " über " + fmtMoney(inv.amountCents) + ", fällig am " + fmtDay(dayIndex(inv.dueIso)),
        verzugLine,
        feeLine,
        pauschaleLine,
        route,
        verzugSince: isoOf(winner.day),
        feeNowMinor,
        costsMinor,
      };
    });

    return {
      business,
      proofLine: !business && feeCents > 0 && carrier
        ? "Ihnen bleibt der Nachweis gestattet, dass uns ein Schaden nicht entstanden oder wesentlich niedriger ist als die berechnete Mahngebühr."
        : "",
      invoices,
    };
  },
});
