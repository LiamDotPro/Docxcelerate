// deno-lint-ignore-file no-explicit-any
import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import { COMPONENTS } from "docxcelerate/registry";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { RoyaltyReserveAndRecoupment } = await loadRegistryModule<
  typeof import("../registry/components/royalty-reserve-and-recoupment/royalty-reserve-and-recoupment.node.tsx")
>("components/royalty-reserve-and-recoupment/royalty-reserve-and-recoupment.node.tsx");
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/royalty-reserve-and-recoupment.deriver.ts")
>("components/derivers/royalty-reserve-and-recoupment.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ royaltyReserveAndRecoupment: mod.run }));

type Obj = Record<string, any>;
type Out = Obj;

const base = { currency: "USD", periodIndex: 9, reservePercent: 25, reserveLiquidationPeriods: 4 };
const run = (royalty: Obj): Out => mod.run([{ ...base, ...royalty }], {} as RuntimeState) as Out;
const throws = (royalty: Obj, text: string) => {
  let message = "";
  try {
    mod.run([royalty], {} as RuntimeState);
  } catch (e) {
    message = e instanceof Error ? e.message : String(e);
  }
  assertStringIncludes(message, text);
};
const earn = (title: string, physicalMinor = 0, digitalMinor = 0) => ({ title, physicalMinor, digitalMinor });
const emptyUnit = { titles: "", status: "", opening: "", net: "", closing: "", payable: "", state: "" };

// ---------------------------------------------------------------------------
// Reserves
// ---------------------------------------------------------------------------

test("e-books carry no reserve", () => {
  const o = run({ earned: [earn("A", 0, 40000)] });
  assertEquals(o.showReserves, false);
  assertEquals([o.held, o.minor.heldTotal, o.outcome], ["US$0.00", 0, "payable"]);
  assertEquals(o.minor.reserveRows, []);
  assertEquals([o.minor.units[0].state, o.minor.units[0].payableMinor], ["no-advance", 40000]);
});

test("the reserve rounds half up, once per title", () => {
  const held = (physicalMinor: number, reservePercent = 25) =>
    run({ reservePercent, earned: [earn("A", physicalMinor)] }).minor.heldTotal;
  assertEquals([held(333), held(334), held(1000, 17.5)], [83, 84, 175]);
  assertEquals(run({ earned: [earn("A", 333), earn("B", 334)] }).minor.heldTotal, 167);
  const two = run({ earned: [earn("A", 2), earn("B", 2)] });
  assertEquals(two.minor.heldTotal, 2);
  assertEquals(two.minor.reserveRows.map((r: Obj) => r.heldMinor), [1, 1]);
  assertEquals(two.minor.netTotal, 2);
});

test("returns are charged oldest first, and nothing is released while the oldest is in use", () => {
  const o = run({
    reservePercent: 0,
    earned: [],
    openReserves: [
      { title: "A", setInPeriod: 7, amountMinor: 20000 },
      { title: "A", setInPeriod: 5, amountMinor: 30000 },
    ],
    returnsCharged: [{ title: "A", amountMinor: 35000 }],
  });
  assertEquals(o.minor.reserveRows, [
    { title: "A", startMinor: 50000, appliedMinor: 35000, releasedMinor: 0, heldMinor: 0, endMinor: 15000 },
  ]);
  assertEquals(o.reserveNote, [
    "On A, returns of US$300.00 were charged against the reserve from period 5.",
    "On A, returns of US$50.00 were charged against the reserve from period 7.",
  ].join(" "));
});

test("a reserve is released in full at the limit, and kept just before it", () => {
  const open = { reservePercent: 0, openReserves: [{ title: "A", setInPeriod: 5, amountMinor: 10000 }] };
  const at9 = run(open);
  assertEquals([at9.minor.netTotal, at9.minor.releasedTotal], [10000, 10000]);
  assertStringIncludes(at9.reserveNote, "was released, having been held for 4 periods.");
  const at8 = run({ ...open, periodIndex: 8 });
  assertEquals([at8.minor.releasedTotal, at8.minor.netTotal, at8.minor.reserveRows[0].endMinor], [0, 0, 10000]);
  assertEquals(at8.reserveNote, "");
});

test("returns are charged before release", () => {
  const o = run({
    reservePercent: 0,
    openReserves: [{ title: "A", setInPeriod: 5, amountMinor: 10000 }],
    returnsCharged: [{ title: "A", amountMinor: 4000 }],
  });
  const row = o.minor.reserveRows[0];
  assertEquals([row.appliedMinor, row.releasedMinor, row.endMinor], [4000, 6000, 0]);
});

test("returns beyond the reserves reduce this period's royalties", () => {
  const o = run({
    earned: [earn("A", 10000)],
    openReserves: [{ title: "A", setInPeriod: 8, amountMinor: 1000 }],
    returnsCharged: [{ title: "A", amountMinor: 3000 }],
  });
  assertEquals(o.minor.reserveRows[0].appliedMinor, 1000);
  assertEquals([o.minor.excessTotal, o.minor.heldTotal, o.minor.netTotal], [2000, 2500, 5500]);
  assertEquals(o.excess, "-US$20.00");
  assertStringIncludes(o.reserveNote, "were more than the reserves held for them");
});

test("returns with no reserve and no earnings are a deficit, not a debt", () => {
  const o = run({ earned: [], returnsCharged: [{ title: "A", amountMinor: 5000 }] });
  assertEquals(o.minor.netTotal, -5000);
  assertEquals([o.units[0].state, o.units[0].closing, o.units[0].payable], ["deficit", "US$50.00", "US$0.00"]);
  assertEquals([o.outcome, o.advanceAdjustment], ["nothing", "US$50.00"]);
});

// ---------------------------------------------------------------------------
// Advances
// ---------------------------------------------------------------------------

const adv = (net: number, opening: number, more: Obj = {}) =>
  run({ reservePercent: 0, earned: [earn("A", net)], advances: [{ titles: ["A"], unrecoupedOpeningMinor: opening }], ...more });

test("an unrecouped advance is never a debt", () => {
  const o = adv(30000, 50000);
  assertEquals([o.minor.units[0].state, o.minor.units[0].closingMinor, o.minor.units[0].payableMinor], ["unrecouped", 20000, 0]);
  assertEquals(o.outcome, "nothing");
  assertStringIncludes(o.units[0].status, "still to earn");
  assertStringIncludes(o.units[0].status, "not a debt");
  assertEquals(o.units[0].closing.startsWith("-"), false);
});

test("an unrecouped advance with a negative net grows", () => {
  const o = run({
    earned: [earn("A", 0)],
    returnsCharged: [{ title: "A", amountMinor: 2000 }],
    advances: [{ titles: ["A"], unrecoupedOpeningMinor: 50000 }],
  });
  assertEquals([o.minor.units[0].state, o.minor.units[0].closingMinor], ["unrecouped", 52000]);
});

test("earning out exactly, and earning out this period", () => {
  const exact = adv(50000, 50000);
  assertEquals([exact.minor.units[0].state, exact.minor.units[0].payableMinor], ["earned-out-now", 0]);
  assertStringIncludes(exact.units[0].status, "earned out exactly");
  const now = adv(70000, 50000);
  assertEquals([now.minor.units[0].state, now.minor.units[0].payableMinor], ["earned-out-now", 20000]);
  assertStringIncludes(now.units[0].status, "the remaining US$200.00 is payable");
});

test("an advance that has already earned out", () => {
  const o = adv(12000, 0);
  assertEquals([o.minor.units[0].state, o.minor.units[0].payableMinor], ["earned-out", 12000]);
  assertStringIncludes(adv(0, 0).units[0].status, "nothing was earned");
});

test("separate advances recoup separately; joint accounting pools", () => {
  const earned = [earn("A", 80000), earn("B", 10000)];
  const apart = run({
    reservePercent: 0,
    earned,
    advances: [{ titles: ["A"], unrecoupedOpeningMinor: 50000 }, { titles: ["B"], unrecoupedOpeningMinor: 50000 }],
  });
  assertEquals(apart.minor.units.map((u: Obj) => u.payableMinor), [30000, 0]);
  assertEquals(apart.minor.totalDue, 30000);
  const joint = run({ reservePercent: 0, earned, advances: [{ titles: ["A", "B"], unrecoupedOpeningMinor: 100000 }] });
  assertEquals([joint.minor.units[0].state, joint.minor.units[0].closingMinor, joint.minor.units[0].payableMinor], ["unrecouped", 10000, 0]);
  assertEquals(joint.units[0].titles, "A and B");
});

// ---------------------------------------------------------------------------
// Minimum payment
// ---------------------------------------------------------------------------

const total = (minor: number, more: Obj = {}) => run({ reservePercent: 0, earned: [earn("A", minor)], ...more });

test("a total below the minimum is carried forward", () => {
  const o = total(1800, { minimumPaymentMinor: 2500 });
  assertEquals([o.outcome, o.payableNow, o.carriedForward, o.minor.carriedForward], ["below-minimum", "US$0.00", "-US$18.00", 1800]);
});

test("carried in, exactly the minimum, no threshold, nothing due", () => {
  const carried = total(1800, { minimumPaymentMinor: 2500, carriedForwardMinor: 800 });
  assertEquals([carried.outcome, carried.minor.totalDue, carried.broughtForward], ["payable", 2600, "US$8.00"]);
  assertEquals(total(2500, { minimumPaymentMinor: 2500 }).outcome, "payable");
  const one = total(1);
  assertEquals([one.outcome, one.payableNow], ["payable", "US$0.01"]);
  const none = total(0);
  assertEquals([none.outcome, none.broughtForward, none.carriedForward], ["nothing", "", ""]);
});

// ---------------------------------------------------------------------------
// Optional fields, slots, throws
// ---------------------------------------------------------------------------

test("missing optional fields do not throw", () => {
  const o = mod.run([{ currency: "USD", periodIndex: 9, earned: [{ title: "A", physicalMinor: 0, digitalMinor: 100 }] }], {} as RuntimeState) as Out;
  assertStringIncludes(o.intro, "Royalties for period 9,");
  assertEquals(o.minimum, "");
  const label = run({
    reservePercent: 0,
    earned: [],
    openReserves: [{ title: "A", setInPeriod: 5, amountMinor: 100 }],
    returnsCharged: [{ title: "A", amountMinor: 50 }],
  });
  assertStringIncludes(label.reserveNote, "the reserve from period 5.");
});

test("slots: one item leaves the rest empty; six fill; seven throw", () => {
  const one = run({ earned: [earn("A", 100)] });
  assertEquals(one.units.slice(1), Array(5).fill(emptyUnit));
  const names = ["A", "B", "C", "D", "E", "F"];
  const six = run({ earned: names.map((n) => earn(n, 1000)) });
  assertEquals(six.reserveRows.map((r: Obj) => r.title), names);
  assertEquals(six.units.map((u: Obj) => u.titles), names);
  const seven = [...names, "G"].map((n) => earn(n, 1000));
  throws({ ...base, earned: seven }, "7 titles carry a reserve; this component prints at most 6");
  throws({ ...base, earned: seven.map((e) => ({ ...e, physicalMinor: 0, digitalMinor: 5 })) }, "7 accounts; this component prints at most 6");
});

test("bad data throws, naming the field", () => {
  const ok = { ...base, earned: [earn("A", 100)] };
  throws({ ...ok, currency: undefined }, "royalty.currency must be an ISO 4217 code");
  throws({ ...ok, currency: "XYZ" }, "royalty.currency");
  throws({ ...ok, periodIndex: 0 }, "royalty.periodIndex");
  throws({ ...ok, periodIndex: 2.5 }, "royalty.periodIndex");
  throws({ ...ok, earned: [earn("A", -1)] }, "royalty.earned[0].physicalMinor");
  throws({ ...ok, earned: [earn("A", 1.5)] }, "royalty.earned[0].physicalMinor");
  throws({ ...ok, reservePercent: 101 }, "royalty.reservePercent");
  throws({ ...ok, openReserves: [{ title: "A", setInPeriod: 9, amountMinor: 1 }] }, "royalty.openReserves[0].setInPeriod must be before royalty.periodIndex");
  throws({ ...ok, advances: [{ titles: ["A"], unrecoupedOpeningMinor: 0 }, { titles: ["A"], unrecoupedOpeningMinor: 0 }] }, "royalty.advances[1].titles[0]");
  throws({ ...ok, advances: [{ titles: [], unrecoupedOpeningMinor: 0 }] }, "royalty.advances[0].titles");
  throws({ ...ok, advances: [{ titles: ["A"] }] }, "royalty.advances[0].unrecoupedOpeningMinor");
  throws({ ...ok, reserveLiquidationPeriods: undefined }, "royalty.reserveLiquidationPeriods");
  throws({ ...ok, reservePercent: 0, reserveLiquidationPeriods: undefined, openReserves: [{ title: "A", setInPeriod: 1, amountMinor: 1 }] }, "royalty.reserveLiquidationPeriods");
  throws({ currency: "USD", periodIndex: 9 }, "royalty has no titles");
});

// ---------------------------------------------------------------------------
// The preview data, figure by figure
// ---------------------------------------------------------------------------

const entry = COMPONENTS.find((c) => c.id === "royalty-reserve-and-recoupment")!;
const preview = (entry.previewData as { royalty: Obj }).royalty;

test("the preview data reconciles", () => {
  const o = run(preview);
  assertEquals(o.minor.reserveRows.map((r: Obj) => [r.startMinor, r.appliedMinor, r.releasedMinor, r.heldMinor, r.endMinor]), [
    [2149530, 702560, 481660, 1203091, 2168401],
    [0, 0, 0, 309178, 309178],
    [610000, 610000, 0, 526225, 526225],
  ]);
  assertEquals(o.reserveTotal, {
    start: "US$27,595.30",
    returns: "US$13,125.60",
    released: "US$4,816.60",
    held: "US$20,384.94",
    end: "US$30,038.04",
  });
  assertEquals(
    [o.minor.physicalTotal, o.minor.digitalTotal, o.minor.heldTotal, o.minor.releasedTotal, o.minor.excessTotal, o.minor.netTotal],
    [8153975, 2689905, 2038494, 481660, 235900, 9051146],
  );
  assertEquals(o.held, "-US$20,384.94");
  assertEquals([o.advanceAdjustment, o.broughtForward, o.minor.totalDue, o.payableNow, o.outcome], ["-US$65,927.75", "US$18.40", 2460211, "US$24,602.11", "payable"]);
  assertEquals(o.minor.units.map((u: Obj) => [u.state, u.openingMinor, u.netMinor, u.payableMinor, u.closingMinor]), [
    ["earned-out-now", 5250000, 7395891, 2145891, 0],
    ["unrecouped", 7500000, 1342775, 0, 6157225],
    ["no-advance", 0, 312480, 312480, 0],
  ]);
  assertEquals(o.units[2].opening, "None");
  assertEquals(o.units[3], emptyUnit);
  assertEquals(o.reserveRows[3].title, "");
  assertStringIncludes(o.reserveNote, "On A Field Guide to Disappearing Places (Illustrated Anniversary Edition), returns of US$2,359.00 were more than");
});

// ---------------------------------------------------------------------------
// Publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Royalties">
      <Section id="body" title="Body">
        <RoyaltyReserveAndRecoupment />
      </Section>
    </Document>,
  );

const state = (data: Obj): RuntimeState => ({
  ctx: { ...data },
  derived: {},
  dataProvider: new InMemoryDataProvider(data),
  aiClient: { generateParagraph: () => "" },
});

function collect(nodes: DocumentNode[], into: DocumentNode[] = []): DocumentNode[] {
  for (const n of nodes) {
    into.push(n);
    const children = (n as { children?: DocumentNode[] }).children;
    if (children) collect(children, into);
  }
  return into;
}

test("a published statement resolves its arms and stays within the branch limit", async () => {
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const arms = collect(published.nodes).filter((n) => (n as { when?: unknown }).when !== undefined);
  assertEquals(arms.length, 18);

  const resolved = async (royalty: Obj) => {
    const out = await resolveDocument(published, state({ royalty }), { derivers });
    const text = new Map(collect(out.nodes).map((n) => [n.id, (n as { text?: string }).text]));
    return text;
  };
  const a = await resolved(preview);
  assertEquals(a.get("royalty-payable"), "US$24,602.11 is payable to you for this period.");
  for (const id of ["royalty-below-minimum", "royalty-nothing-payable", "royalty-reserve-4", "royalty-advance-4"]) {
    assertEquals(a.has(id), false);
  }
  assertEquals(a.has("royalty-reserve-3"), true);
  assertEquals(a.has("royalty-advance-3-status"), true);

  const b = await resolved({ ...base, earned: [earn("A", 0, 1800)], minimumPaymentMinor: 2500 });
  assertStringIncludes(b.get("royalty-below-minimum") ?? "", "US$18.00 due for this period is below the minimum payment of US$25.00");
  assertEquals(b.has("royalty-payable"), false);
  assertEquals(b.has("royalty-reserves"), false);
});
