// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

interface Slot {
  number: string;
  dueDate: string;
  overdue: string;
  amount: string;
  line: string;
}

interface Reminder {
  outcome: "remind" | "none";
  level: 0 | 1 | 2 | 3;
  title: string;
  opening: string;
  invoices: Slot[];
  overflow: string;
  total: string;
  request: string;
  contact: string;
  consequence: string;
  paidSince: string;
  promptFacts: string;
  deadline: string;
  sent: Array<{ number: string; level: 1 | 2 | 3 }>;
}

/**
 * Decides which overdue invoices a payment reminder covers, its level, title,
 * deadline and what it may say happens next. Installed by
 * `dxcl add dunning-reminder`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper and number is declared inside `run`.
 */
export default deriver({
  name: "dunningReminder",
  run: (reminder: unknown): Reminder => {
    const rem = reminder as any;

    // The numbers an owner edits. The floors are as researched, not verified
    // against current statute text; check them before relying on one.
    const defaults = { graceDays: 7, minIntervalDays: 14, deadlineDays: { first: 14, second: 10, final: 7 } };
    const floors = {
      // BW 6:96(6): 14 days after receipt.
      nlDays: 14,
      // shortcut: post assumed received two days after sending, upgrade if tracked delivery dates are available
      nlPostReceiptDays: 2,
      // WER art. XIX.2: 14 days from the third working day after sending on paper, or the day after sending by email.
      beDays: 14,
      beWorkingDays: 3,
      // Pre-Action Protocol for Debt Claims.
      ukDays: 30,
    };
    const maxRows = 8;

    // UTC day indexes, so a day does not shift with the engine's time zone.
    const dayIndex = (iso: string) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d) / 86_400_000;
    };
    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date (YYYY-MM-DD)");
      }
      const iso = value.slice(0, 10);
      if (new Date(dayIndex(iso) * 86_400_000).toISOString().slice(0, 10) !== iso) {
        throw new Error(field + " is not a real date");
      }
      return iso;
    };
    const count = (value: unknown, field: string, min: number, fallback: number): number => {
      if (value === undefined || value === null) return fallback;
      if (typeof value !== "number" || !Number.isInteger(value) || value < min) {
        throw new Error(field + " must be an integer of at least " + min);
      }
      return value;
    };
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const fmtDay = (index: number) => dates.format(new Date(index * 86_400_000));

    if (typeof rem !== "object" || rem === null) throw new Error("reminder is required");

    // 1. Validate everything first, so bad data throws even when nothing is sent.
    const sendIso = isoDay(rem.sendDate, "reminder.sendDate");
    const send = dayIndex(sendIso);
    const checkedIso = isoDay(rem.paymentsCheckedTo, "reminder.paymentsCheckedTo");
    if (dayIndex(checkedIso) > send) throw new Error("reminder.paymentsCheckedTo is after reminder.sendDate");

    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency: rem.currency });
    } catch {
      throw new Error("reminder.currency must be an ISO 4217 code");
    }
    const digits = money.resolvedOptions().maximumFractionDigits as number;
    const fmtMoney = (minor: number) => money.format(minor / 10 ** digits);

    const graceDays = count(rem.graceDays, "reminder.graceDays", 0, defaults.graceDays);
    const minInterval = count(rem.minIntervalDays, "reminder.minIntervalDays", 1, defaults.minIntervalDays);
    const given = rem.deadlineDays ?? {};
    if (typeof given !== "object") throw new Error("reminder.deadlineDays must be an object");
    const deadlineDays = [
      0,
      count(given.first, "reminder.deadlineDays.first", 1, defaults.deadlineDays.first),
      count(given.second, "reminder.deadlineDays.second", 1, defaults.deadlineDays.second),
      count(given.final, "reminder.deadlineDays.final", 1, defaults.deadlineDays.final),
    ];

    const floorNames = ["NL-consumer", "BE-consumer", "UK-pre-action"];
    if (rem.legalFloor !== undefined && !floorNames.includes(rem.legalFloor)) {
      throw new Error("reminder.legalFloor must be one of " + floorNames.join(", "));
    }
    if (rem.channel !== undefined && rem.channel !== "post" && rem.channel !== "email") {
      throw new Error("reminder.channel must be post or email");
    }
    const email = rem.channel === "email";

    const consequenceNames = ["suspension", "agency", "court"];
    if (rem.intendedConsequences !== undefined && !Array.isArray(rem.intendedConsequences)) {
      throw new Error("reminder.intendedConsequences must be an array");
    }
    const intended: string[] = rem.intendedConsequences ?? [];
    intended.forEach((c, i) => {
      if (!consequenceNames.includes(c)) {
        throw new Error("reminder.intendedConsequences[" + i + "] must be one of " + consequenceNames.join(", "));
      }
    });

    let sinceIso = "";
    if (rem.customerSince !== undefined && rem.customerSince !== null && rem.customerSince !== "") {
      sinceIso = isoDay(rem.customerSince, "reminder.customerSince");
      if (dayIndex(sinceIso) > send) throw new Error("reminder.customerSince is after reminder.sendDate");
    }

    if (rem.invoices !== undefined && rem.invoices !== null && !Array.isArray(rem.invoices)) {
      throw new Error("reminder.invoices must be an array");
    }
    const all: any[] = rem.invoices ?? [];
    const included: Array<{ number: string; dueIso: string; minor: number; overdue: number; level: number }> = [];
    all.forEach((inv, i) => {
      const at = "reminder.invoices[" + i + "]";
      if (typeof inv?.number !== "string" || inv.number.trim() === "") throw new Error(at + ".number is required");
      const dueIso = isoDay(inv.dueDate, at + ".dueDate");
      if (typeof inv.openAmount !== "number" || !Number.isFinite(inv.openAmount)) {
        throw new Error(at + ".openAmount must be a finite number");
      }
      const lastLevel = inv.lastLevel ?? 0;
      if (!Number.isInteger(lastLevel) || lastLevel < 0 || lastLevel > 3) {
        throw new Error(at + ".lastLevel must be an integer from 0 to 3");
      }
      let sinceLast = Infinity;
      if (lastLevel >= 1) {
        if (inv.lastReminderDate === undefined || inv.lastReminderDate === null || inv.lastReminderDate === "") {
          throw new Error(at + ".lastReminderDate is required when lastLevel is " + lastLevel);
        }
        const lastIso = isoDay(inv.lastReminderDate, at + ".lastReminderDate");
        sinceLast = send - dayIndex(lastIso);
        if (sinceLast < 0) throw new Error(at + ".lastReminderDate is after reminder.sendDate");
      }

      // 2. Minor units, so every sum is an integer.
      const minor = Math.round(inv.openAmount * 10 ** digits);
      // 3. Eligibility.
      const overdue = send - dayIndex(dueIso);
      if (
        overdue > 0 && overdue >= graceDays && minor > 0 && inv.disputed !== true && inv.onHold !== true &&
        lastLevel < 3 && (lastLevel === 0 || sinceLast >= minInterval)
      ) {
        included.push({ number: inv.number, dueIso, minor, overdue, level: lastLevel + 1 });
      }
    });

    const emptySlot = (): Slot => ({ number: "", dueDate: "", overdue: "", amount: "", line: "" });
    const slots = Array.from({ length: maxRows }, emptySlot);

    // 4. Nothing to chase prints nothing.
    if (included.length === 0) {
      return {
        outcome: "none", level: 0, title: "", opening: "", invoices: slots, overflow: "", total: "", request: "",
        contact: "", consequence: "", paidSince: "", promptFacts: "", deadline: "", sent: [],
      };
    }

    // 5. The letter is as firm as its firmest invoice.
    const level = Math.max(...included.map((i) => i.level)) as 1 | 2 | 3;
    const mixed = included.some((i) => i.level < level);

    // 6. Oldest first, ties by number.
    included.sort((a, b) =>
      a.dueIso < b.dueIso ? -1 : a.dueIso > b.dueIso ? 1 : a.number < b.number ? -1 : a.number > b.number ? 1 : 0
    );
    const one = included.length === 1;
    const sum = (rows: typeof included) => rows.reduce((t, i) => t + i.minor, 0);

    // 7. Slots.
    included.slice(0, maxRows).forEach((inv, i) => {
      const overdue = inv.overdue === 1 ? "1 day" : inv.overdue + " days";
      const dueDate = fmtDay(dayIndex(inv.dueIso));
      const amount = fmtMoney(inv.minor);
      slots[i] = {
        number: inv.number,
        dueDate,
        overdue,
        amount,
        line: "Invoice " + inv.number + ", due " + dueDate + ", " + overdue + " overdue: " + amount,
      };
    });

    // 8. Overflow.
    const rest = included.slice(maxRows);
    const overflow = rest.length === 0 ? "" : rest.length + " further invoice" + (rest.length === 1 ? "" : "s") +
      ", totalling " + fmtMoney(sum(rest)) + ", " + (rest.length === 1 ? "is" : "are") +
      " also overdue and included in the total.";

    // 9, 10. Title and opening.
    const title = ["", "Payment reminder", "Second reminder", "Final reminder"][level];
    const some = mixed ? "some of " : "";
    const below = one ? "the invoice below" : "the invoices below";
    const opening = level === 1
      ? "Our records show that " + (one ? "the invoice below is" : "the invoices below are") + " past " +
        (one ? "its" : "their") + " due date. This may simply have been overlooked, so this is a friendly reminder."
      : level === 2
      ? "This is our second reminder about " + some + below + ". We have not yet received payment."
      : "This is our final reminder about " + some + below + ". Despite earlier reminders, " +
        (one ? "it remains" : "they remain") + " unpaid.";

    // 11. Deadline: never before the floor, never on a weekend.
    const base = send + deadlineDays[level];
    let floor = -Infinity;
    if (rem.legalFloor === "NL-consumer") {
      floor = send + (email ? 0 : floors.nlPostReceiptDays) + floors.nlDays;
    } else if (rem.legalFloor === "BE-consumer") {
      if (email) {
        floor = send + 1 + floors.beDays;
      } else {
        let day = send;
        for (let worked = 0; worked < floors.beWorkingDays;) {
          day += 1;
          const weekday = new Date(day * 86_400_000).getUTCDay();
          if (weekday !== 0 && weekday !== 6) worked += 1;
        }
        floor = day + floors.beDays;
      }
    } else if (rem.legalFloor === "UK-pre-action") {
      floor = send + floors.ukDays;
    }
    let due = Math.max(base, floor);
    const weekday = new Date(due * 86_400_000).getUTCDay();
    if (weekday === 6) due += 2;
    else if (weekday === 0) due += 1;
    const deadline = fmtDay(due);

    const total = fmtMoney(sum(included));
    // 12, 13. Request and contact.
    const request = level === 3
      ? "Please make sure " + total + " reaches us by " + deadline + "."
      : "Please pay " + total + " by " + deadline + ".";
    const contact = level === 1 ? "" : "If something is wrong with " + (one ? "the invoice" : "any of these invoices") +
      ", or you cannot pay in full by then, please contact us before " + deadline + " so we can agree a way forward.";

    // 14. Consequence: only what is intended, and never for a first-time invoice.
    let consequence = "";
    if (level === 3 && intended.length > 0) {
      const phrases: Record<string, string> = {
        suspension: "suspend your service",
        agency: "pass the debt to a collection agency",
        court: "start a court claim to recover it",
      };
      const list = new Intl.ListFormat("en-GB", { type: "conjunction" });
      const what = list.format(consequenceNames.filter((c) => intended.includes(c)).map((c) => phrases[c]));
      const final = included.filter((i) => i.level === 3);
      consequence = final.length === included.length
        ? "If we have not received payment by " + deadline + ", we intend to " + what + "."
        : "If invoice" + (final.length === 1 ? " " : "s ") + list.format(final.map((i) => i.number)) +
          (final.length === 1 ? " is" : " are") + " not paid by " + deadline + ", we intend to " + what + ".";
    }

    // 15. Payments cross in the post.
    const paidSince = "This reminder takes account of payments received up to and including " +
      fmtDay(dayIndex(checkedIso)) + ". If you have paid since then, thank you, and please ignore this reminder.";

    // 16. Facts for the generated opening.
    let promptFacts = "This is the " + ["", "first", "second", "final"][level] + " reminder. The oldest invoice in it is " +
      Math.max(...included.map((i) => i.overdue)) + " days overdue.";
    if (sinceIso) {
      const [y1, m1, d1] = sinceIso.split("-").map(Number);
      const [y2, m2, d2] = sendIso.split("-").map(Number);
      const months = (y2 - y1) * 12 + (m2 - m1) - (d2 < d1 ? 1 : 0);
      const years = Math.floor(months / 12);
      promptFacts += " The customer has been with us for " + (months >= 12
        ? years + (years === 1 ? " year" : " years")
        : months >= 1
        ? months + (months === 1 ? " month" : " months")
        : "less than a month") + ".";
    }

    return {
      outcome: "remind", level, title, opening, invoices: slots, overflow, total, request, contact, consequence,
      paidSince, promptFacts, deadline, sent: included.map((i) => ({ number: i.number, level: i.level as 1 | 2 | 3 })),
    };
  },
});
