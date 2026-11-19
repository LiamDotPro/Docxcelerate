// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Applies a period's time entries to prepaid hour blocks, closes the blocks the
 * period ends, and writes the ledger and every note a retainer invoice prints.
 * Installed by `dxcl add prepaid-hours-drawdown`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 *
 * Everything is integer minutes and integer minor units. Known limit: with an
 * increment such as 1 or 10 minutes, a displayed two-decimal figure can differ
 * by 0.01 from the sum of the displayed figures above it. The arithmetic is
 * exact; only the printing rounds.
 */
export default deriver({
  name: "prepaidHoursDrawdown",
  run: (retainer: unknown) => {
    const rt = retainer as any;

    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency: rt?.currency });
    } catch {
      throw new Error("retainer.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmtMoney = (minor: number) => money.format(minor / 10 ** exponent);
    const hoursFmt = new Intl.NumberFormat("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const fmtH = (minutes: number) => hoursFmt.format(minutes / 60);
    const fmtSigned = (minutes: number) =>
      minutes > 0 ? "+" + fmtH(minutes) : minutes < 0 ? "−" + fmtH(-minutes) : fmtH(0);
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const fmtD = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));
    const listFmt = new Intl.ListFormat("en-GB");
    // Every input is >= 0, so Math.round is round-half-up.
    const round = (x: number) => Math.round(x);
    const cmp = (a: any, b: any) => (a < b ? -1 : a > b ? 1 : 0);

    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    const dayIndex = (iso: string) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d) / 86_400_000;
    };
    const minutesOf = (hours: unknown, field: string): number => {
      if (typeof hours !== "number" || !Number.isFinite(hours) || hours < 0) {
        throw new Error(field + " must be a finite number of hours of 0 or more");
      }
      const m = Math.round(hours * 60);
      if (Math.abs(hours * 60 - m) >= 1e-6) throw new Error(field + " must be a whole number of minutes");
      return m;
    };
    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
        throw new Error(field + " must be a non-negative integer of minor units");
      }
      return value;
    };

    const periodEnd = isoDay(rt?.periodEnd, "retainer.periodEnd");
    const increment = rt.increment;
    if (typeof increment !== "number" || !Number.isInteger(increment) || increment <= 0) {
      throw new Error("retainer.increment must be a positive integer number of minutes");
    }
    const overageRate = rt.overageRate === undefined ? undefined : minorOf(rt.overageRate, "retainer.overageRate");
    if (rt.lowBalanceHours !== undefined) {
      if (typeof rt.lowBalanceHours !== "number" || !Number.isFinite(rt.lowBalanceHours) || rt.lowBalanceHours < 0) {
        throw new Error("retainer.lowBalanceHours must be a finite number of 0 or more");
      }
    }
    let expiringDays = 30;
    if (rt.expiringDays !== undefined) {
      if (typeof rt.expiringDays !== "number" || !Number.isInteger(rt.expiringDays) || rt.expiringDays < 0) {
        throw new Error("retainer.expiringDays must be a non-negative integer");
      }
      expiringDays = rt.expiringDays;
    }

    // Blocks.
    const rawBlocks: any[] = Array.isArray(rt.blocks) ? rt.blocks : [];
    const seen = new Set<string>();
    const blocks = rawBlocks.map((b: any, i: number) => {
      const f = "retainer.blocks[" + i + "]";
      if (typeof b?.id !== "string" || b.id.trim() === "") throw new Error(f + ".id is required");
      if (seen.has(b.id)) throw new Error(f + ".id is duplicated");
      seen.add(b.id);
      const hours = minutesOf(b.hours, f + ".hours");
      const opening = b.opening === undefined ? hours : minutesOf(b.opening, f + ".opening");
      const rate = minorOf(b.rate, f + ".rate");
      const start = isoDay(b.start, f + ".start");
      const expiry = isoDay(b.expiry, f + ".expiry");
      if (expiry < start) throw new Error(f + ".expiry is before " + f + ".start");
      let cap: number | undefined;
      if (b.rolloverCap !== undefined) {
        const c = b.rolloverCap;
        const hasH = c?.hours !== undefined;
        const hasP = c?.percent !== undefined;
        if (hasH === hasP) throw new Error(f + ".rolloverCap needs exactly one of hours or percent");
        if (hasH) {
          cap = minutesOf(c.hours, f + ".rolloverCap.hours");
        } else {
          if (typeof c.percent !== "number" || !(c.percent >= 0 && c.percent <= 100)) {
            throw new Error(f + ".rolloverCap.percent must be from 0 to 100");
          }
          cap = Math.floor((hours * c.percent) / 100 + 1e-9);
        }
      }
      const rolledExpiry = b.rolledExpiry === undefined ? undefined : isoDay(b.rolledExpiry, f + ".rolledExpiry");
      const roles: string[] | undefined = Array.isArray(b.allowedRoles) && b.allowedRoles.length > 0
        ? b.allowedRoles
        : undefined;
      return {
        i,
        id: b.id as string,
        name: typeof b.name === "string" && b.name !== "" ? (b.name as string) : (b.id as string),
        opening,
        rate,
        start,
        expiry,
        cap,
        rolledExpiry,
        roles,
        from: b.rolledInFrom as string | undefined,
        // Set below: the block that receives this block's rolled hours.
        receiver: undefined as any,
      };
    });
    const byId = new Map<string, any>(blocks.map((b: any) => [b.id, b]));
    for (const b of blocks) {
      if (b.from === undefined) continue;
      const f = "retainer.blocks[" + b.i + "].rolledInFrom";
      const giver = byId.get(b.from);
      if (!giver) throw new Error(f + " names no block");
      if (giver === b) throw new Error(f + " names the block itself");
      if (giver.receiver) throw new Error(f + " names a block another block already rolls into");
      giver.receiver = b;
    }
    for (const g of blocks) {
      const f = "retainer.blocks[" + g.i + "]";
      if (g.rolledExpiry !== undefined && g.receiver) {
        if (g.rolledExpiry < g.receiver.start || g.rolledExpiry > g.receiver.expiry) {
          throw new Error(f + ".rolledExpiry must fall within the receiving block's dates");
        }
      }
      if (g.cap !== undefined && g.expiry <= periodEnd && !g.receiver) {
        throw new Error(f + " has a rollover cap and ends this period, but no block names it in rolledInFrom");
      }
    }

    // Entries.
    const rawEntries: any[] = Array.isArray(rt.entries) ? rt.entries : [];
    const entries = rawEntries.map((e: any, i: number) => {
      const f = "retainer.entries[" + i + "]";
      const date = isoDay(e?.date, f + ".date");
      if (date > periodEnd) throw new Error(f + ".date is after retainer.periodEnd");
      if (typeof e.timekeeper !== "string" || e.timekeeper.trim() === "") throw new Error(f + ".timekeeper is required");
      if (typeof e.minutes !== "number" || !Number.isInteger(e.minutes) || e.minutes < 0) {
        throw new Error(f + ".minutes must be a non-negative integer");
      }
      return {
        i,
        date,
        timekeeper: e.timekeeper as string,
        role: typeof e.role === "string" ? (e.role as string) : undefined,
        billed: Math.ceil(e.minutes / increment) * increment,
        standardRate: e.standardRate,
      };
    }).filter((e: any) => e.billed > 0).sort((a: any, b: any) => cmp(a.date, b.date) || a.i - b.i);

    // Pools.
    const mk = (b: any, over: any = {}) => ({
      b,
      balance: b.opening,
      expiry: b.expiry,
      start: b.start,
      roles: b.roles as string[] | undefined,
      rate: b.rate,
      rolled: false,
      closed: false,
      used: 0,
      forfeit: 0,
      value: 0,
      out: 0,
      giver: undefined as any,
      ...over,
    });
    const pools: any[] = blocks.map((b: any) => mk(b));
    const burn = (a: any, b: any) =>
      cmp(a.expiry, b.expiry) || cmp(a.start, b.start) || (a.rolled ? 0 : 1) - (b.rolled ? 0 : 1) || a.b.i - b.b.i;

    const rolledNotes: string[] = [];
    const lapsed: Array<{ minutes: number; label: string }> = [];
    const closePool = (p: any) => {
      p.closed = true;
      const unused = p.balance;
      p.balance = 0;
      let forfeit = unused;
      if (!p.rolled && p.b.cap !== undefined && unused > 0) {
        const rolled = Math.min(unused, p.b.cap);
        forfeit = unused - rolled;
        if (rolled > 0) {
          const r = p.b.receiver;
          p.out = rolled;
          pools.push(mk(r, {
            balance: rolled,
            expiry: p.b.rolledExpiry ?? r.expiry,
            start: r.start,
            roles: p.roles,
            rate: p.rate,
            rolled: true,
            giver: p.b,
            rolledIn: rolled,
          }));
          rolledNotes.push(
            fmtH(rolled) + " hours rolled from " + p.b.name + " into " + r.name + " (capped at " + fmtH(p.b.cap) +
              " hours) and expire on " + fmtD(p.b.rolledExpiry ?? r.expiry),
          );
        }
      }
      if (forfeit > 0) {
        p.forfeit = forfeit;
        p.value = round((forfeit * p.rate) / 60);
        lapsed.push({
          minutes: forfeit,
          label: p.rolled ? "rolled hours from " + p.giver.name + " lapsed on " + fmtD(p.expiry) : p.b.name,
        });
      }
    };
    // A rolled pool made by one close may itself be due, so look again each time.
    const closeDue = (due: (p: any) => boolean) => {
      for (;;) {
        const next = pools.filter((p) => !p.closed && due(p)).sort(burn)[0];
        if (!next) return;
        closePool(next);
      }
    };

    // Drawing.
    const out: any[] = [];
    const overs: Array<{ minutes: number; rate: number; amount: number }> = [];
    const splits: string[] = [];
    const excludedRoles: string[] = [];
    for (const e of entries) {
      closeDue((p) => p.expiry < e.date);
      const live = pools.filter((p) => !p.closed && p.balance > 0 && p.start <= e.date && e.date <= p.expiry).sort(burn);
      const eligible = live.filter((p) => !p.roles || (e.role !== undefined && p.roles.includes(e.role)));
      let remaining = e.billed;
      const parts: any[] = [];
      for (const p of eligible) {
        if (remaining === 0) break;
        const take = Math.min(remaining, p.balance);
        p.balance -= take;
        p.used += take;
        remaining -= take;
        parts.push({ blockId: p.b.id, rolled: p.rolled, minutes: take, amount: 0, name: p.b.name });
      }
      if (remaining > 0) {
        const rate = overageRate ?? e.standardRate;
        if (rate === undefined) {
          throw new Error("retainer.entries[" + e.i + "]: overage needs retainer.overageRate or a standardRate");
        }
        minorOf(rate, "retainer.entries[" + e.i + "].standardRate");
        const amount = round((remaining * rate) / 60);
        overs.push({ minutes: remaining, rate, amount });
        parts.push({ blockId: null, rolled: false, minutes: remaining, amount, name: "" });
        if (live.length > eligible.length && e.role !== undefined && !excludedRoles.includes(e.role)) {
          excludedRoles.push(e.role);
        }
      }
      if (parts.length > 1) {
        const clauses = parts.map((p) => fmtH(p.minutes) + " hours " + (p.blockId === null ? "as overage" : "from " + p.name));
        const joined = clauses.length === 2
          ? clauses[0] + " and " + clauses[1]
          : clauses.slice(0, -1).join(", ") + " and " + clauses[clauses.length - 1];
        splits.push(
          "the entry of " + fmtD(e.date) + " by " + e.timekeeper + " (" + fmtH(e.billed) + " hours) drew " + joined,
        );
      }
      out.push({
        index: e.i,
        date: e.date,
        billedMinutes: e.billed,
        parts: parts.map((p) => ({ blockId: p.blockId, rolled: p.rolled, minutes: p.minutes, amount: p.amount })),
      });
    }
    closeDue((p) => p.expiry <= periodEnd);

    // Ledger, in burn order by block.
    const rows = [...blocks].sort((a: any, b: any) => cmp(a.expiry, b.expiry) || cmp(a.start, b.start) || a.i - b.i).map((b: any) => {
      const mine = pools.filter((p) => p.b === b);
      const own = mine.find((p) => !p.rolled);
      const rolledPool = mine.find((p) => p.rolled);
      const rolled = (rolledPool?.rolledIn ?? 0) - own.out;
      const row = {
        id: b.id,
        name: b.name,
        opening: b.opening,
        used: own.used + (rolledPool?.used ?? 0),
        rolled,
        forfeited: own.forfeit + (rolledPool?.forfeit ?? 0),
        closing: own.balance + (rolledPool?.balance ?? 0),
        expires: own.closed
          ? "Ended " + fmtD(b.expiry)
          : rolledPool && rolledPool.expiry !== b.expiry
          ? fmtD(b.expiry) + "; rolled hours " + fmtD(rolledPool.expiry)
          : fmtD(b.expiry),
      };
      return row;
    });
    const sum = (list: any[], k: string) => list.reduce((t, r) => t + r[k], 0);
    const cells = (r: any, name: string, expires: string) => ({
      name,
      opening: fmtH(r.opening),
      used: fmtH(r.used),
      rolled: fmtSigned(r.rolled),
      forfeited: fmtH(r.forfeited),
      closing: fmtH(r.closing),
      expires,
    });
    const blank = { name: "", opening: "", used: "", rolled: "", forfeited: "", closing: "", expires: "" };
    const slots: any[] = rows.length <= 4
      ? rows.map((r) => cells(r, r.name, r.expires))
      : [
        ...rows.slice(0, 3).map((r) => cells(r, r.name, r.expires)),
        cells(
          { opening: sum(rows.slice(3), "opening"), used: sum(rows.slice(3), "used"), rolled: sum(rows.slice(3), "rolled"), forfeited: sum(rows.slice(3), "forfeited"), closing: sum(rows.slice(3), "closing") },
          rows.length - 3 + " other blocks",
          "Various",
        ),
      ];
    while (slots.length < 4) slots.push({ ...blank });

    // Balance and expiry warnings.
    const open = pools.filter((p) => !p.closed).sort(burn);
    const remaining = open.reduce((t, p) => t + p.balance, 0);
    const balanceArm: "exhausted" | "low" | "none" = blocks.length > 0 && remaining === 0
      ? "exhausted"
      : blocks.length > 0 && rt.lowBalanceHours !== undefined && remaining <= rt.lowBalanceHours * 60
      ? "low"
      : "none";
    const periodLabel = fmtD(periodEnd);
    const balanceNote = balanceArm === "exhausted"
      ? "No prepaid hours remain after " + periodLabel + ". Further time is billed as overage until more hours are bought."
      : balanceArm === "low"
      ? fmtH(remaining) + " hours remain across your prepaid blocks after " + periodLabel + "."
      : "";

    const expiringNote = expiringDays === 0 ? "" : open
      .filter((p) => p.balance > 0 && dayIndex(p.expiry) <= dayIndex(periodEnd) + expiringDays)
      .map((p) => {
        const lapse = !p.rolled && p.b.cap !== undefined && p.b.receiver ? p.balance - p.b.cap : 0;
        return fmtH(p.balance) + " hours in " + (p.rolled ? p.b.name + " (rolled hours)" : p.b.name) + " expire on " +
          fmtD(p.expiry) +
          (lapse > 0 ? "; up to " + fmtH(p.b.cap) + " can roll over and " + fmtH(lapse) + " would lapse if unused" : "") + ".";
      }).join(" ");

    // Notes.
    const overMinutes = overs.reduce((t, o) => t + o.minutes, 0);
    const overAmount = overs.reduce((t, o) => t + o.amount, 0);
    let overageNote = "";
    if (overMinutes > 0) {
      const groups = new Map<number, { minutes: number; amount: number }>();
      for (const o of overs) {
        const g = groups.get(o.rate) ?? { minutes: 0, amount: 0 };
        g.minutes += o.minutes;
        g.amount += o.amount;
        groups.set(o.rate, g);
      }
      const clauses = [...groups.entries()].sort((a, b) => b[0] - a[0]).map(([rate, g]) =>
        fmtH(g.minutes) + " hours at " + fmtMoney(rate) + " an hour (" + fmtMoney(g.amount) + ")"
      );
      const joined = clauses.length === 1
        ? clauses[0]
        : clauses.slice(0, -1).join(", ") + " and " + clauses[clauses.length - 1];
      const roleList = listFmt.format(excludedRoles);
      overageNote = fmtH(overMinutes) + " hours were not covered by prepaid hours and are billed as overage: " + joined +
        ", " + fmtMoney(overAmount) + " in all." +
        (roleList ? " " + roleList[0].toUpperCase() + roleList.slice(1) + " time is not covered by any block it could draw on." : "");
    }
    const forfeitedMinutes = lapsed.reduce((t, l) => t + l.minutes, 0);
    const forfeitedValue = pools.reduce((t, p) => t + p.value, 0);
    const forfeitedNote = forfeitedMinutes === 0 ? "" : fmtH(forfeitedMinutes) + " hours unused at the end of " +
      listFmt.format(lapsed.map((l) => l.label)) + " were above the rollover cap and have lapsed. Their value, " +
      fmtMoney(forfeitedValue) + " at the block rate, is not credited or refunded." +
      (rt.vatOnPurchase === true
        ? " The VAT charged when the hours were bought is not adjusted: the prepayment was taxed when it was received."
        : "");

    return {
      show: blocks.length > 0,
      periodEnd: periodLabel,
      increment: increment === 1 ? "1 minute" : increment + " minutes",
      blocks: slots,
      showTotal: blocks.length >= 2,
      total: {
        opening: fmtH(sum(rows, "opening")),
        used: fmtH(sum(rows, "used")),
        rolled: fmtSigned(sum(rows, "rolled")),
        forfeited: fmtH(sum(rows, "forfeited")),
        closing: fmtH(sum(rows, "closing")),
      },
      overageNote,
      splitNote: splits.length === 0
        ? ""
        : "Where an entry ran past the hours it could use in a block, it was split at that point: " + splits.join("; ") + ".",
      rolledNote: rolledNotes.length === 0 ? "" : rolledNotes.join("; ") + ".",
      forfeitedNote,
      expiringNote,
      balanceArm,
      balanceNote,
      entries: out,
      overageMinutes: overMinutes,
      overageAmount: overAmount,
      forfeitedMinutes,
      forfeitedValue,
      closing: rows.map((r) => ({ id: r.id, minutes: r.closing })),
      // Raw minutes per row, for the ledger invariant in tests.
      ledger: rows,
    };
  },
});
