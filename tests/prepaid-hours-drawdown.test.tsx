import { test } from "node:test";
import { COMPONENTS } from "docxcelerate/registry";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/prepaid-hours-drawdown.deriver.ts")
>("components/derivers/prepaid-hours-drawdown.deriver.ts");
const { PrepaidHoursDrawdown } = await loadRegistryModule<
  typeof import("../registry/components/prepaid-hours-drawdown/prepaid-hours-drawdown.node.tsx")
>("components/prepaid-hours-drawdown/prepaid-hours-drawdown.node.tsx");

// Deriver output and test inputs are loosely shaped on purpose.
// deno-lint-ignore no-explicit-any
type Obj = Record<string, any>;
const run = (retainer: Obj): Obj => mod.run([retainer], {} as RuntimeState) as Obj;
const base = { currency: "GBP", increment: 6, periodEnd: "2026-04-30" };
const block = (id: string, more: Obj = {}) => ({
  id,
  hours: 10,
  rate: 20000,
  start: "2026-01-01",
  expiry: "2026-12-31",
  ...more,
});
const entry = (date: string, minutes: number, more: Obj = {}) => ({ date, timekeeper: "Sam", minutes, ...more });
const go = (more: Obj) => run({ ...base, ...more });
const row = (o: Obj, id: string) => o.ledger.find((r: Obj) => r.id === id);
const parts = (o: Obj, i = 0) => o.entries[i].parts.map((p: Obj) => [p.blockId, p.minutes]);

// A block that ends this period and rolls into "b".
const giver = (more: Obj = {}) => block("a", { hours: 40, opening: 12, expiry: "2026-03-31", rolloverCap: { hours: 8 }, ...more });
const receiver = (more: Obj = {}) => block("b", { hours: 40, start: "2026-04-01", expiry: "2026-06-30", rolledInFrom: "a", ...more });

/** closing = opening - used + rolled - forfeited, on every row. */
function invariant(o: Obj) {
  for (const r of o.ledger) assertEquals(r.closing, r.opening - r.used + r.rolled - r.forfeited);
}

test("1. exhaustion split", () => {
  const o = go({ blocks: [block("a", { opening: 1 })], entries: [entry("2026-03-10", 100, { standardRate: 30000 })] });
  assertEquals(o.entries[0].billedMinutes, 102);
  assertEquals(parts(o), [["a", 60], [null, 42]]);
  assertEquals(o.overageAmount, 21000);
  assertStringIncludes(o.overageNote, "£210.00");
  assertEquals(o.splitNote !== "", true);
  assertEquals([row(o, "a").closing, o.balanceArm], [0, "exhausted"]);
});

test("2. round before split", () => {
  const o = go({ blocks: [block("a", { opening: 50 / 60 })], entries: [entry("2026-03-10", 52, { standardRate: 30000 })] });
  assertEquals(parts(o), [["a", 50], [null, 4]]);
  assertEquals(o.overageAmount, 2000);
});

test("3. earliest expiry, not purchase order", () => {
  const o = go({
    blocks: [block("a", { opening: 10 }), block("b", { opening: 5, start: "2026-02-01", expiry: "2026-06-30" })],
    entries: [entry("2026-03-02", 120)],
  });
  assertEquals(parts(o), [["b", 120]]);
  assertEquals(o.closing, [{ id: "b", minutes: 180 }, { id: "a", minutes: 600 }]);
  assertEquals([o.blocks[0].closing, o.blocks[1].closing], ["3.00", "10.00"]);
});

test("4. entry date, not invoice date", () => {
  const blocks = [block("a", { opening: 2, expiry: "2026-03-31" })];
  const o = go({ blocks, entries: [entry("2026-03-31", 60)] });
  assertEquals(parts(o), [["a", 60]]);
  assertEquals([o.forfeitedMinutes, o.forfeitedValue], [60, 20000]);
  assertEquals(o.forfeitedNote.includes("VAT"), false);
  assertStringIncludes(go({ blocks, entries: [entry("2026-03-31", 60)], vatOnPurchase: true }).forfeitedNote, "VAT");
});

test("5. hours cap", () => {
  const o = go({ blocks: [giver(), receiver({ opening: 40 })] });
  const [a, b] = [o.blocks[0], o.blocks[1]];
  assertEquals([a.rolled, a.forfeited, a.closing, a.expires], ["−8.00", "4.00", "0.00", "Ended 31 March 2026"]);
  assertEquals([b.rolled, b.closing], ["+8.00", "48.00"]);
  assertEquals(o.forfeitedValue, 80000);
  invariant(o);
  assertStringIncludes(o.rolledNote, "rolled from a into b");
});

test("6. percent cap", () => {
  const o = go({ blocks: [giver({ rolloverCap: { percent: 25 } }), receiver()] });
  assertEquals([o.blocks[0].rolled, o.blocks[0].forfeited], ["−10.00", "2.00"]);
});

test("7. rolled hours that do not extend", () => {
  const blocks = [giver({ rolledExpiry: "2026-05-15" }), receiver({ opening: 40 })];
  const o = go({ blocks, entries: [entry("2026-04-10", 120)] });
  assertEquals(o.entries[0].parts[0].rolled, true);
  assertEquals(o.blocks[1].expires, "30 June 2026; rolled hours 15 May 2026");
  invariant(o);
  const late = go({ blocks: [giver({ rolledExpiry: "2026-05-15" }), receiver({ opening: 40, rate: 90000 })], entries: [entry("2026-04-10", 120)], periodEnd: "2026-05-31" });
  assertEquals(row(late, "b").forfeited, 360);
  assertEquals(late.forfeitedValue, 200000); // 4h above the cap, then the 6h left in the rolled pool, both at the giver's rate, not b's
  invariant(late);
});

test("7b. a rolled pool keeps the giver's roles", () => {
  const blocks = [giver({ allowedRoles: ["associate"] }), receiver({ opening: 0 })];
  const at = (role: string) => parts(go({ blocks, entries: [entry("2026-04-10", 60, { role, standardRate: 1 })] }));
  assertEquals(at("associate"), [["b", 60]]);
  assertEquals(at("partner"), [[null, 60]]); // cannot draw the rolled 8h in b
  const own = go({ blocks: [giver({ allowedRoles: ["associate"] }), receiver({ opening: 5 })], entries: [entry("2026-04-10", 60, { role: "partner" })] });
  assertEquals(parts(own), [["b", 60]]); // but can draw b's own pool
  assertEquals(own.entries[0].parts[0].rolled, false);
});

test("8. rolled hours never roll twice", () => {
  const c = block("c", { hours: 40, start: "2026-07-01", expiry: "2026-09-30", rolledInFrom: "b" });
  const o = go({
    periodEnd: "2026-09-30",
    blocks: [giver(), receiver({ opening: 0, rolloverCap: { hours: 100 } }), c],
  });
  // a's 8h roll into b, lapse with b, and none of them reach c.
  assertEquals(row(o, "b").forfeited, 480);
  assertEquals(row(o, "c").rolled, 0);
  invariant(o);
});

test("9. role restriction", () => {
  const blocks = [block("a", { allowedRoles: ["associate", "paralegal"] })];
  const o = go({ overageRate: 40000, blocks, entries: [entry("2026-03-10", 60, { role: "partner" })] });
  assertEquals([o.overageMinutes, o.overageAmount], [60, 40000]);
  assertStringIncludes(o.overageNote, "Partner time");
  assertEquals(parts(go({ blocks, entries: [entry("2026-03-10", 60, { role: "associate" })] })), [["a", 60]]);
  assertEquals(go({ overageRate: 1, blocks, entries: [entry("2026-03-10", 60)] }).overageMinutes, 60);
});

test("10. overage rate fallback", () => {
  const blocks = [block("a", { opening: 0 })];
  const one = (more: Obj) => go({ blocks, entries: [entry("2026-03-10", 60, { standardRate: 30000 })], ...more });
  assertEquals(one({}).overageAmount, 30000);
  assertEquals(one({ overageRate: 50000 }).overageAmount, 50000);
  assertThrows(() => go({ blocks, entries: [entry("2026-03-10", 60)] }), "entries[0]");
  const two = go({ blocks, entries: [entry("2026-03-10", 60, { standardRate: 30000 }), entry("2026-03-11", 60, { standardRate: 90000 })] });
  assertEquals(two.overageAmount, 120000);
  assertEquals(two.overageNote.indexOf("£900.00") < two.overageNote.indexOf("£300.00"), true);
});

test("11. zero and empty", () => {
  const blocks = [block("a")];
  const z = go({ blocks, entries: [entry("2026-03-10", 0)] });
  assertEquals([z.entries, z.splitNote, z.overageNote, z.blocks[0].used], [[], "", "", "0.00"]);
  assertEquals(go({ blocks }).blocks[0].used, "0.00");
  for (const none of [{}, { blocks: [] }]) {
    const o = go({ ...none, overageRate: 100, entries: [entry("2026-03-10", 60)] });
    assertEquals([o.show, o.entries[0].parts], [false, [{ blockId: null, rolled: false, minutes: 60, amount: 100 }]]);
  }
});

test("12. invalid input throws", () => {
  const e = [entry("2026-03-10", 6)];
  const bad: Array<[string, Obj]> = [
    ["entries[0].minutes", { entries: [entry("2026-03-10", -6)] }],
    ["blocks[0].hours", { blocks: [block("a", { hours: -1 })] }],
    ["blocks[0].opening", { blocks: [block("a", { opening: -1 })] }],
    ["blocks[0].rate", { blocks: [block("a", { rate: -1 })] }],
    ["whole number of minutes", { blocks: [block("a", { hours: 0.01 })] }],
    ["retainer.increment", { increment: 0 }],
    ["retainer.increment", { increment: undefined }],
    ["retainer.increment", { increment: 2.5 }],
    ["rolloverCap", { blocks: [giver({ rolloverCap: { hours: 1, percent: 5 } }), receiver()] }],
    ["rolloverCap", { blocks: [giver({ rolloverCap: {} }), receiver()] }],
    ["percent", { blocks: [giver({ rolloverCap: { percent: 101 } }), receiver()] }],
    ["rolledInFrom", { blocks: [block("a", { rolledInFrom: "zzz" })] }],
    ["rolledInFrom", { blocks: [block("a", { rolledInFrom: "a" })] }],
    ["rolledInFrom", { blocks: [block("a"), block("b", { rolledInFrom: "a" }), block("c", { rolledInFrom: "a" })] }],
    ["blocks[1].id", { blocks: [block("a"), block("a")] }],
    ["blocks[0].expiry", { blocks: [block("a", { expiry: "2025-12-31" })] }],
    ["entries[0].date", { entries: [entry("2026-05-01", 6)] }],
    ["retainer.currency", { currency: "XXXX" }],
    ["rolledInFrom", { blocks: [giver()] }],
    ["rolledExpiry", { blocks: [giver({ rolledExpiry: "2026-07-01" }), receiver()] }],
    ["rolledExpiry", { blocks: [giver({ rolledExpiry: "2026-03-31" }), receiver()] }],
  ];
  for (const [text, more] of bad) assertThrows(() => go({ blocks: [block("a")], entries: e, ...more }), text);
  assertThrows(() => go({ currency: undefined }), "retainer.currency");
  assertThrows(() => go({ periodEnd: "April" }), "retainer.periodEnd");
});

test("13. a capped block that does not close needs no receiver", () => {
  const o = go({ blocks: [block("a", { rolloverCap: { hours: 4 } })] });
  assertEquals(o.show, true);
});

test("14. one block has no total", () => {
  assertEquals(go({ blocks: [block("a")] }).showTotal, false);
  assertEquals(go({ blocks: [block("a"), block("b")] }).showTotal, true);
});

test("15. many blocks", () => {
  const five = go({ blocks: [1, 2, 3, 4, 5].map((n) => block("b" + n, { expiry: "2026-1" + (n % 3) + "-01", opening: n })) });
  assertEquals(five.blocks.map((b: Obj) => b.name), ["b3", "b1", "b4", "2 other blocks"]);
  assertEquals([five.blocks[3].opening, five.blocks[3].expires], ["7.00", "Various"]);
  const four = go({ blocks: [1, 2, 3, 4].map((n) => block("b" + n)) });
  assertEquals(four.blocks.every((b: Obj) => b.name !== ""), true);
  assertEquals(go({ blocks: [block("a")] }).blocks[1], { name: "", opening: "", used: "", rolled: "", forfeited: "", closing: "", expires: "" });
});

test("16. one entry split across two blocks and overage", () => {
  const o = go({
    blocks: [block("a", { opening: 1, expiry: "2026-06-30" }), block("b", { opening: 1 })],
    entries: [entry("2026-03-10", 180, { standardRate: 60000 })],
  });
  assertEquals(parts(o), [["a", 60], ["b", 60], [null, 60]]);
  assertEquals(o.entries[0].parts.reduce((t: number, p: Obj) => t + p.minutes, 0), o.entries[0].billedMinutes);
  assertStringIncludes(o.splitNote, "drew 1.00 hours from a, 1.00 hours from b and 1.00 hours as overage");
});

test("17. same-date entries keep input order", () => {
  const o = go({
    blocks: [block("a", { opening: 1 })],
    entries: [entry("2026-03-10", 60, { timekeeper: "First" }), entry("2026-03-10", 60, { standardRate: 1 })],
  });
  assertEquals([parts(o, 0), parts(o, 1)], [[["a", 60]], [[null, 60]]]);
});

test("18. low, exhausted and none", () => {
  const left = (hours: number, more: Obj = {}) => go({ blocks: [block("a", { opening: hours })], ...more }).balanceArm;
  assertEquals(left(3, { lowBalanceHours: 5 }), "low");
  assertEquals(left(0, { lowBalanceHours: 5 }), "exhausted");
  assertEquals(left(0), "exhausted");
  assertEquals(left(3), "none");
  assertEquals(go({ blocks: [block("a", { expiry: "2026-03-31", opening: 0 })] }).balanceArm, "exhausted");
  assertEquals(go({ blocks: [block("a", { expiry: "2026-03-31", opening: 3 })] }).balanceArm, "exhausted");
});

test("19. expiring", () => {
  const blocks = [block("a", { opening: 10, expiry: "2026-05-20", rolloverCap: { hours: 4 } }), block("b", { start: "2026-05-21", rolledInFrom: "a" })];
  const note = go({ blocks }).expiringNote;
  for (const s of ["10.00 hours", "20 May 2026", "4.00 can roll", "6.00 would lapse"]) assertStringIncludes(note, s);
  assertEquals(go({ blocks, expiringDays: 0 }).expiringNote, "");
  assertEquals(go({ blocks: [block("a", { expiry: "2026-06-15" })] }).expiringNote, "");
});

// The node, published and resolved by an engine that only has the bundle.
const preview = COMPONENTS.find((c) => c.id === "prepaid-hours-drawdown")!.previewData as Obj;
const documentOf = () =>
  template(
    <Document title="Retainer invoice">
      <PrepaidHoursDrawdown />
    </Document>,
  );

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((n) => [(n as { id?: string }).id ?? "", ...ids((n as { children?: DocumentNode[] }).children ?? [])]);
}

test("20. publishing then resolving matches building directly", async () => {
  const published = await buildDocument(documentOf(), createPublishData() as typeof preview, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const state: RuntimeState = {
    ctx: { ...preview },
    derived: {},
    dataProvider: new InMemoryDataProvider(preview),
    aiClient: { generateParagraph: () => "" },
  };
  const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ prepaidHoursDrawdown: mod.run }));
  const resolved = ids((await resolveDocument(published, state, { derivers })).nodes);
  const direct = ids((await buildDocument(documentOf(), preview, { dynamicMode: "placeholder" })).nodes);

  assertEquals(resolved, direct);
  for (const id of ["block-1", "block-2", "block-3", "block-4", "ledger-total", "split", "rolled", "forfeited", "expiring", "overage"]) {
    assertEquals(resolved.includes("prepaid-hours-" + id), true);
  }
  // The preview shows every behaviour at once: a split into overage, a block with no cap lapsing, a capped rollover.
  const shown = mod.run([preview.retainer], {} as RuntimeState) as Obj;
  assertStringIncludes(shown.splitNote, "as overage");
  assertEquals(row(shown, "q4-carried").forfeited > 0, true);
  assertEquals(row(shown, "q1").rolled < 0, true);
  // The preview leaves Q2 and the paralegal block open, so neither warning fires.
  for (const id of ["low", "exhausted"]) assertEquals(resolved.includes("prepaid-hours-" + id), false);
});

test("21. no blocks builds to nothing", async () => {
  const data = { retainer: { ...base, blocks: [], entries: [] } };
  assertEquals((await buildDocument(documentOf(), data, { dynamicMode: "placeholder" })).nodes.length, 0);
});

function assertThrows(fn: () => unknown, text: string) {
  try {
    fn();
  } catch (error) {
    assertStringIncludes((error as Error).message, text);
    return;
  }
  throw new Error("expected a throw containing " + text);
}
