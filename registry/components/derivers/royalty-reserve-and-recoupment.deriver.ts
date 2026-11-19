// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Works a royalty statement down from what was earned to what is paid, and
 * decides what each advance and the total mean. Installed by
 * `dxcl add royalty-reserve-and-recoupment`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "royaltyReserveAndRecoupment",
  run: (royalty: unknown) => {
    const roy = royalty as any;
    const fail = (message: string): never => {
      throw new Error(message);
    };

    // Currency first: the exponent decides how every amount is read and printed.
    // Intl accepts any three letters, so a code it does not list is refused here.
    const currency = roy?.currency;
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency });
      const known = (Intl as any).supportedValuesOf?.("currency") as string[] | undefined;
      if (known && !known.includes(String(currency).toUpperCase())) throw new Error("unknown");
    } catch {
      throw new Error("royalty.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmt = (minor: number) => money.format(minor / 10 ** exponent);
    // A deduction is printed negative, but zero is never "-US$0.00".
    const neg = (minor: number) => fmt(minor === 0 ? 0 : -minor);
    const list = new Intl.ListFormat("en-GB", { style: "long", type: "conjunction" });
    const percent = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 2 });

    const nonNegative = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
        fail(field + " must be a non-negative integer of minor units");
      }
      return value as number;
    };
    const optionalMinor = (value: unknown, field: string): number =>
      value === undefined || value === null ? 0 : nonNegative(value, field);
    const titleOf = (value: unknown, field: string): string => {
      if (typeof value !== "string" || value.trim() === "") fail(field + " is required");
      return value as string;
    };
    const listOf = (value: unknown): any[] => (Array.isArray(value) ? value : []);

    const periodIndex = roy.periodIndex;
    if (typeof periodIndex !== "number" || !Number.isInteger(periodIndex) || periodIndex < 1) {
      fail("royalty.periodIndex must be an integer of 1 or more");
    }
    const periodLabel = typeof roy.periodLabel === "string" && roy.periodLabel.trim() !== ""
      ? roy.periodLabel
      : "period " + periodIndex;

    const reservePercent = roy.reservePercent ?? 0;
    if (typeof reservePercent !== "number" || !Number.isFinite(reservePercent) || reservePercent < 0 || reservePercent > 100) {
      fail("royalty.reservePercent must be a number from 0 to 100");
    }
    const bp = Math.round(reservePercent * 100);

    const minimum = roy.minimumPaymentMinor === undefined || roy.minimumPaymentMinor === null
      ? undefined
      : nonNegative(roy.minimumPaymentMinor, "royalty.minimumPaymentMinor");
    const broughtForward = optionalMinor(roy.carriedForwardMinor, "royalty.carriedForwardMinor");

    // Titles in order of first appearance, with their summed figures.
    const titles: string[] = [];
    const figures = new Map<string, { P: number; D: number; R: number; reserves: any[] }>();
    const figuresOf = (title: string) => {
      let f = figures.get(title);
      if (!f) {
        f = { P: 0, D: 0, R: 0, reserves: [] };
        figures.set(title, f);
        titles.push(title);
      }
      return f;
    };
    listOf(roy.earned).forEach((e, i) => {
      const f = figuresOf(titleOf(e?.title, "royalty.earned[" + i + "].title"));
      f.P += optionalMinor(e.physicalMinor, "royalty.earned[" + i + "].physicalMinor");
      f.D += optionalMinor(e.digitalMinor, "royalty.earned[" + i + "].digitalMinor");
    });
    const open = listOf(roy.openReserves);
    open.forEach((o, i) => {
      const at = "royalty.openReserves[" + i + "]";
      const f = figuresOf(titleOf(o?.title, at + ".title"));
      if (typeof o.setInPeriod !== "number" || !Number.isInteger(o.setInPeriod)) {
        fail(at + ".setInPeriod must be an integer");
      }
      if (o.setInPeriod >= periodIndex) fail(at + ".setInPeriod must be before royalty.periodIndex");
      f.reserves.push({
        setIn: o.setInPeriod as number,
        amount: nonNegative(o.amountMinor, at + ".amountMinor"),
        label: typeof o.label === "string" && o.label.trim() !== "" ? o.label : "period " + o.setInPeriod,
      });
    });
    listOf(roy.returnsCharged).forEach((r, i) => {
      const at = "royalty.returnsCharged[" + i + "]";
      figuresOf(titleOf(r?.title, at + ".title")).R += nonNegative(r.amountMinor, at + ".amountMinor");
    });
    const advanceInputs = listOf(roy.advances);
    const advanceOf = new Map<string, number>();
    advanceInputs.forEach((a, i) => {
      const at = "royalty.advances[" + i + "]";
      if (!Array.isArray(a?.titles) || a.titles.length === 0) fail(at + ".titles must not be empty");
      a.titles.forEach((t: unknown, j: number) => {
        figuresOf(titleOf(t, at + ".titles[" + j + "]"));
        if (advanceOf.has(t as string)) fail(at + ".titles[" + j + "] is already in another advance");
        advanceOf.set(t as string, i);
      });
      nonNegative(a.unrecoupedOpeningMinor, at + ".unrecoupedOpeningMinor");
    });
    if (titles.length === 0) fail("royalty has no titles");

    const anyPhysical = titles.some((t) => figures.get(t)!.P > 0);
    let liquidation = 0;
    if (open.length > 0 || (bp > 0 && anyPhysical)) {
      liquidation = roy.reserveLiquidationPeriods;
      if (typeof liquidation !== "number" || !Number.isInteger(liquidation) || liquidation < 1) {
        fail("royalty.reserveLiquidationPeriods must be an integer of 1 or more when a reserve exists");
      }
    }

    // Per-title ledger: returns are charged oldest first, then aged reserves
    // are released, so an aged reserve absorbs returns before its remainder goes.
    const notes: string[] = [];
    const ledgers = titles.map((title) => {
      const f = figures.get(title)!;
      const reserves = f.reserves.slice().sort((a, b) => a.setIn - b.setIn);
      const start = reserves.reduce((s, r) => s + r.amount, 0);
      const sentences: string[] = [];
      let left = f.R;
      let applied = 0;
      for (const r of reserves) {
        const take = Math.min(left, r.amount);
        applied += take;
        r.amount -= take;
        left -= take;
        if (take > 0) {
          sentences.push("On " + title + ", returns of " + fmt(take) + " were charged against the reserve from " + r.label + ".");
        }
      }
      const excess = left;
      let released = 0;
      for (const r of reserves) {
        if (periodIndex - r.setIn >= liquidation && r.amount > 0) {
          released += r.amount;
          sentences.push(
            "On " + title + ", the remaining " + fmt(r.amount) + " of the reserve from " + r.label +
              " was released, having been held for " + (periodIndex - r.setIn) + " periods.",
          );
          r.amount = 0;
        }
      }
      if (excess > 0) {
        sentences.push(
          "On " + title + ", returns of " + fmt(excess) +
            " were more than the reserves held for them and were charged against this period's royalties.",
        );
      }
      // Half up on integers, once per title, so the rows add up to the total.
      if (!Number.isSafeInteger(f.P * bp)) fail("royalty.earned physical royalties are too large to reserve");
      const held = Math.floor((f.P * bp + 5000) / 10000);
      const end = start - applied - released + held;
      notes.push(...sentences);
      return { title, P: f.P, D: f.D, start, applied, released, held, excess, end, net: f.P + f.D - held - excess + released };
    });

    const reserveLedgers = ledgers.filter((l) => l.start > 0 || l.held > 0);
    if (reserveLedgers.length > 6) {
      fail("royaltyReserveAndRecoupment: " + reserveLedgers.length + " titles carry a reserve; this component prints at most 6");
    }

    // One accounting unit per advance, then each unadvanced title alone.
    const netOf = new Map(ledgers.map((l) => [l.title, l.net]));
    const rawUnits: Array<{ titles: string[]; U: number; noAdvance: boolean }> = [
      ...advanceInputs.map((a) => ({ titles: a.titles as string[], U: a.unrecoupedOpeningMinor as number, noAdvance: false })),
      ...titles.filter((t) => !advanceOf.has(t)).map((t) => ({ titles: [t], U: 0, noAdvance: true })),
    ];
    if (rawUnits.length > 6) {
      fail("royaltyReserveAndRecoupment: " + rawUnits.length + " accounts; this component prints at most 6");
    }
    const units = rawUnits.map((u) => {
      const net = u.titles.reduce((s, t) => s + netOf.get(t)!, 0);
      const closing = Math.max(0, u.U - net);
      const payable = Math.max(0, net - u.U);
      const state = u.U === 0 && net < 0
        ? "deficit"
        : u.U > 0 && net < u.U
        ? "unrecouped"
        : u.U > 0
        ? "earned-out-now"
        : !u.noAdvance
        ? "earned-out"
        : "no-advance";
      const names = list.format(u.titles);
      const status = state === "unrecouped"
        ? "The advance on " + names + " has " + fmt(closing) +
          " still to earn from future royalties, so nothing is payable on it yet. This is not a debt: an advance is recovered only from royalties, never repaid."
        : state === "earned-out-now"
        ? payable > 0
          ? "The advance on " + names + " earned out this period: " + fmt(u.U) + " of this period's royalties completed it, and the remaining " + fmt(payable) + " is payable."
          : "The advance on " + names + " earned out exactly this period, so nothing more is payable on it yet; royalties from the next period are payable in full."
        : state === "earned-out"
        ? net > 0
          ? "The advance on " + names + " has already earned out, so the " + fmt(net) + " earned this period is payable in full."
          : "The advance on " + names + " has already earned out, and nothing was earned on it this period."
        : state === "no-advance"
        ? net > 0
          ? "There is no advance on " + names + ", so the " + fmt(net) + " earned this period is payable in full."
          : "There is no advance on " + names + ", and nothing was earned on it this period."
        : "Returns on " + names + " this period were " + fmt(-net) + " more than its royalties. That is set against future royalties on it, and is not a debt.";
      return { titles: u.titles, names, status, state, U: u.U, noAdvance: u.noAdvance, net, closing, payable };
    });

    const sum = (xs: number[]) => xs.reduce((s, x) => s + x, 0);
    const physicalTotal = sum(ledgers.map((l) => l.P));
    const digitalTotal = sum(ledgers.map((l) => l.D));
    const heldTotal = sum(ledgers.map((l) => l.held));
    const releasedTotal = sum(ledgers.map((l) => l.released));
    const excessTotal = sum(ledgers.map((l) => l.excess));
    const netTotal = physicalTotal + digitalTotal - heldTotal + releasedTotal - excessTotal;
    if (netTotal !== sum(ledgers.map((l) => l.net))) fail("royaltyReserveAndRecoupment: settlement does not reconcile");
    const payablesTotal = sum(units.map((u) => u.payable));
    const totalDue = payablesTotal + broughtForward;

    const outcome = totalDue === 0 ? "nothing" : minimum !== undefined && totalDue < minimum ? "below-minimum" : "payable";
    const payableNow = outcome === "payable" ? totalDue : 0;
    const carriedForward = outcome === "below-minimum" ? totalDue : 0;

    const blankRow = { title: "", start: "", returns: "", released: "", held: "", end: "" };
    const reserveRows = reserveLedgers.map((l) => ({
      title: l.title,
      start: fmt(l.start),
      returns: fmt(l.applied),
      released: fmt(l.released),
      held: fmt(l.held),
      end: fmt(l.end),
    }));
    while (reserveRows.length < 6) reserveRows.push({ ...blankRow });
    const blankUnit = { titles: "", status: "", opening: "", net: "", closing: "", payable: "", state: "" };
    const unitRows = units.map((u) => ({
      titles: u.names,
      status: u.status,
      opening: u.noAdvance ? "None" : fmt(u.U),
      net: fmt(u.net),
      closing: fmt(u.state === "deficit" ? -u.net : u.closing),
      payable: fmt(u.payable),
      state: u.state,
    }));
    while (unitRows.length < 6) unitRows.push({ ...blankUnit });

    const col = (pick: (l: (typeof ledgers)[number]) => number) => fmt(sum(reserveLedgers.map(pick)));

    return {
      intro: "Royalties for " + periodLabel +
        ", from what was earned to what is paid. A reserve against returns is held on physical formats only; e-books and digital audio carry none.",
      outcome,
      payableNow: fmt(payableNow),
      totalDue: fmt(totalDue),
      minimum: minimum === undefined ? "" : fmt(minimum),
      physical: fmt(physicalTotal),
      digital: fmt(digitalTotal),
      heldLabel: "Reserve held against returns (" + percent.format(bp / 100) + "% of physical royalties)",
      held: neg(heldTotal),
      released: fmt(releasedTotal),
      excess: neg(excessTotal),
      net: fmt(netTotal),
      advanceAdjustment: fmt(payablesTotal - netTotal),
      broughtForward: broughtForward === 0 ? "" : fmt(broughtForward),
      carriedForward: outcome === "below-minimum" ? neg(carriedForward) : "",
      showReserves: reserveLedgers.length > 0,
      reserveRows,
      reserveTotal: {
        start: col((l) => l.start),
        returns: col((l) => l.applied),
        released: col((l) => l.released),
        held: col((l) => l.held),
        end: col((l) => l.end),
      },
      reserveNote: notes.join(" "),
      units: unitRows,
      minor: {
        physicalTotal,
        digitalTotal,
        heldTotal,
        releasedTotal,
        excessTotal,
        netTotal,
        payablesTotal,
        totalDue,
        payableNow,
        carriedForward,
        reserveRows: reserveLedgers.map((l) => ({
          title: l.title,
          startMinor: l.start,
          appliedMinor: l.applied,
          releasedMinor: l.released,
          heldMinor: l.held,
          endMinor: l.end,
        })),
        units: units.map((u) => ({
          titles: u.titles,
          state: u.state,
          openingMinor: u.U,
          netMinor: u.net,
          closingMinor: u.state === "deficit" ? -u.net : u.closing,
          payableMinor: u.payable,
        })),
      },
    };
  },
});
