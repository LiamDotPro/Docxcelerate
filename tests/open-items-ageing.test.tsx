import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";
import { COMPONENTS } from "../src/registry/catalog.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/open-items-ageing.deriver.ts")
>("components/derivers/open-items-ageing.deriver.ts");
const { OpenItemsAgeing } = await loadRegistryModule<
  typeof import("../registry/components/open-items-ageing/open-items-ageing.node.tsx")
>("components/open-items-ageing/open-items-ageing.node.tsx");

// deno-lint-ignore no-explicit-any
type Obj = Record<string, any>;
const base = { currency: "GBP", statementDate: "2026-11-30" };
const inv = (number: string, dueDate: string, extra: Obj = {}): Obj => ({
  number,
  date: "2026-01-01",
  dueDate,
  amount: 100,
  ...extra,
});
// The deriver is synchronous, but a test should not care.
// The result type, read off the deriver itself rather than declared by hand.
type Result = ReturnType<NonNullable<typeof mod.__types>>;
const run = async (account: Obj) =>
  await mod.run([{ ...base, ...account }], {} as RuntimeState) as Result;
const minors = async (account: Obj) => (await run(account)).buckets.map((b) => b.amountMinor);

// ---- E1 -------------------------------------------------------------------
test("E1: bucket edges include their top edge", async () => {
  for (
    const [due, days, bucket] of [
      ["2026-11-30", 0, 0],
      ["2026-10-31", 30, 1],
      ["2026-10-30", 31, 2],
      ["2026-10-01", 60, 2],
      ["2026-09-30", 61, 3],
      ["2026-09-01", 90, 3],
      ["2026-08-31", 91, 4],
    ] as const
  ) {
    const r = await run({ invoices: [inv("A", due)] });
    const want = [0, 0, 0, 0, 0, 0];
    want[bucket] = 10000;
    assertEquals(r.items[0].ageDays, days);
    assertEquals(r.buckets.map((b) => b.amountMinor), want);
  }
  assertEquals((await run({ invoices: [inv("A", "2026-12-05")] })).items[0].age, "Not yet due");
});

// ---- E2 -------------------------------------------------------------------
test("E2: age follows the due date, not the invoice date", async () => {
  const r = await run({ invoices: [inv("A", "2026-12-01", { date: "2026-10-01" })] });
  assertEquals([r.items[0].age, r.overdue], ["Not yet due", false]);
  assertEquals(r.nothingOverdue, "Nothing on this account is overdue.");
});

// ---- E3, E4 ---------------------------------------------------------------
test("E3: default terms, and the basis sentence says so", async () => {
  const noDue = { number: "A", date: "2026-10-15", amount: 100 };
  const r = await run({ defaultTermsDays: 30, invoices: [noDue] });
  assertEquals([r.items[0].due, r.items[0].age, r.buckets[1].amountMinor], ["14 Nov 2026", "16 days", 10000]);
  assertEquals(r.basis.includes("counted as due 30 days after"), true);
  assertEquals((await run({ defaultTermsDays: 1, invoices: [noDue] })).basis.includes("due 1 day after"), true);
  assertEquals((await run({ defaultTermsDays: 30, invoices: [inv("A", "2026-11-01")] })).basis, "Aged by days past each invoice's due date, at 30 November 2026.");
});

test("E4: missing terms throw, naming the invoice; a settled one does not", async () => {
  await assertRejects(() => run({ invoices: [{ number: "X9", date: "2026-10-01", amount: 5 }] }), Error, "invoice X9");
  const settled = { number: "X9", date: "2026-10-01", amount: 5 };
  const credits = [{ type: "payment", date: "2026-10-02", number: "P", amount: 5, allocations: [{ invoice: "X9", amount: 5 }] }];
  assertEquals((await run({ invoices: [settled], credits })).empty, true);
});

// ---- E5 - E8 --------------------------------------------------------------
test("E5: unallocated credit is not netted", async () => {
  const r = await run({
    invoices: [inv("A", "2026-08-01", { amount: 1000 })],
    credits: [{ type: "payment", date: "2026-11-01", number: "P", amount: 1500 }],
  });
  assertEquals(r.buckets.map((b) => [b.amountMinor, b.percent]), [[0, 0], [0, 0], [0, 0], [0, 0], [100000, 100], [0, 0]]);
  assertEquals([r.unallocated, r.balanceLabel, r.balance], ["-£1,500.00", "In credit", "£500.00"]);
});

test("E6: partial allocation", async () => {
  const r = await run({
    invoices: [inv("A", "2026-08-01", { amount: 1000 })],
    credits: [{ type: "payment", date: "2026-11-01", number: "P", amount: 1000, allocations: [{ invoice: "A", amount: 400 }] }],
  });
  assertEquals([r.items[0].amount, r.items[0].open], ["£1,000.00", "£600.00"]);
  assertEquals([r.unallocatedMinor, r.balanceMinor, r.balanceLabel, r.balance], [60000, 0, "Balance", "£0.00"]);
});

test("E7: settled and empty", async () => {
  const settled = {
    invoices: [inv("A", "2026-08-01")],
    credits: [{ type: "payment", date: "2026-11-01", number: "P", amount: 100, allocations: [{ invoice: "A", amount: 100 }] }],
  };
  for (const account of [settled, {}]) {
    const r = await run(account);
    assertEquals([r.items.length, r.empty, r.hasItems], [0, true, false]);
    assertEquals(r.emptyLine, "There are no open items on this account at 30 November 2026.");
  }
});

test("E8: credit only", async () => {
  const r = await run({ credits: [{ type: "credit-note", date: "2026-11-02", number: "CN-1", amount: 50 }] });
  assertEquals([r.empty, r.hasItems, r.overdue], [false, false, false]);
  assertEquals(r.nothingOverdue, "Nothing on this account is overdue.");
  assertEquals([r.balanceLabel, r.balance], ["In credit", "£50.00"]);
  assertEquals(r.unallocatedDetail, "Not yet set against an invoice: credit note CN-1 of 2 November 2026 (£50.00).");
});

// ---- E9, E10 --------------------------------------------------------------
test("E9: over-allocation throws", async () => {
  const pay = (n: string, amount: number, allocations: Obj[]) => ({ type: "payment", date: "2026-11-01", number: n, amount, allocations });
  await assertRejects(
    () => run({ invoices: [inv("A", "2026-08-01")], credits: [pay("P1", 100, [{ invoice: "A", amount: 60 }]), pay("P2", 100, [{ invoice: "A", amount: 60 }])] }),
    Error,
    "is over-allocated",
  );
  await assertRejects(
    () => run({ invoices: [inv("A", "2026-08-01", { amount: 500 })], credits: [pay("P1", 100, [{ invoice: "A", amount: 150 }])] }),
    Error,
    "allocates more than its amount",
  );
});

test("E10: bad references", async () => {
  await assertRejects(
    () => run({ invoices: [inv("A", "2026-08-01")], credits: [{ type: "payment", date: "2026-11-01", number: "P", amount: 10, allocations: [{ invoice: "Z", amount: 10 }] }] }),
    Error,
    "names invoice Z",
  );
  await assertRejects(() => run({ invoices: [inv("A", "2026-08-01"), inv("A", "2026-08-02")] }), Error, "invoice A twice");
});

// ---- E11 ------------------------------------------------------------------
test("E11: disputed invoices have their own bucket", async () => {
  const d = inv("D", "2026-06-01", { amount: 5000, disputed: true });
  const r = await run({ invoices: [d] });
  assertEquals([r.buckets[4].amountMinor, r.buckets[5].amountMinor, r.items[0].age], [0, 500000, "In dispute"]);
  assertEquals([r.overdue, r.nothingOverdue], [false, "Nothing on this account is overdue, leaving aside £5,000.00 in dispute."]);
  const credits = [{ type: "payment", date: "2026-11-01", number: "P", amount: 5000, allocations: [{ invoice: "D", amount: 5000 }] }];
  assertEquals((await run({ invoices: [d], credits })).items.length, 0);
});

// ---- E12 ------------------------------------------------------------------
test("E12: shares are largest-remainder and sum to 100", async () => {
  const r = await run({ invoices: [inv("A", "2026-12-01"), inv("B", "2026-11-15"), inv("C", "2026-10-15")] });
  assertEquals(r.buckets.map((b) => b.percent), [34, 33, 33, 0, 0, 0]);
  const tiny = await run({ invoices: [inv("A", "2026-08-01", { amount: 999.99 }), inv("B", "2026-11-15", { amount: 0.01 })] });
  assertEquals(tiny.buckets.map((b) => [b.share]), [["0%"], ["0%"], ["0%"], ["0%"], ["100%"], ["0%"]]);

  let seed = 7;
  const rand = () => (seed = (seed * 48271) % 2147483647) / 2147483647;
  for (let round = 0; round < 50; round++) {
    const dues = ["2026-12-01", "2026-11-15", "2026-10-15", "2026-09-15", "2026-07-15"];
    const invoices = dues.map((d, i) => inv("N" + i, d, { amount: rand() < 0.3 ? 0.01 : Math.round(rand() * 100000) / 100 + 0.01 }));
    const out = await run({ invoices });
    assertEquals(out.buckets.reduce((s, b) => s + b.percent, 0), 100);
    for (const b of out.buckets) if (b.amountMinor === 0) assertEquals(b.percent, 0);
  }
});

// ---- E13, E14, E15 ----------------------------------------------------------
test("E13: rounding", async () => {
  const credits = [{ type: "payment", date: "2026-11-01", number: "P", amount: 0.3, allocations: [{ invoice: "A", amount: 0.1 }, { invoice: "A", amount: 0.2 }] }];
  assertEquals((await run({ invoices: [inv("A", "2026-08-01", { amount: 0.3 })], credits })).items.length, 0);
  const jpy = await run({ currency: "JPY", invoices: [inv("A", "2026-08-01", { amount: 1234.5 })] });
  assertEquals([jpy.items[0].open, jpy.items[0].openMinor], ["JP¥1,235", 1235]);
});

test("E14: zero, negative and non-finite values", async () => {
  assertEquals((await run({ invoices: [inv("A", "2026-08-01", { amount: 0 })] })).items.length, 0);
  const pay = (over: Obj) => ({ type: "payment", date: "2026-11-01", number: "P", amount: 10, ...over });
  const bad: Obj[] = [
    { invoices: [inv("A", "2026-08-01", { amount: -5 })] },
    { credits: [pay({ amount: 0 })] },
    { invoices: [inv("A", "2026-08-01")], credits: [pay({ allocations: [{ invoice: "A", amount: 0 }] })] },
    { invoices: [inv("A", "2026-08-01", { amount: Infinity })] },
    { invoices: [inv("A", "2026-08-01", { amount: NaN })] },
    { invoices: [inv("A", "2026-08-01", { amount: "5" })] },
    { credits: [pay({ type: "refund" })] },
    { defaultTermsDays: -1 },
    { defaultTermsDays: 2.5 },
  ];
  for (const account of bad) await assertRejects(() => run(account), Error);
});

test("E15: bad dates", async () => {
  await assertRejects(() => run({ statementDate: undefined }), Error, "account.statementDate");
  await assertRejects(() => run({ statementDate: "30/11/2026" }), Error, "account.statementDate");
  const invoices = [inv("A", "2026-10-31")];
  assertEquals((await run({ statementDate: "2026-11-30T23:59:59Z", invoices })).items[0].ageDays, 30);
});

// ---- E16, E17, E18 ----------------------------------------------------------
test("E16: ordering", async () => {
  const order = async (invoices: Obj[]) => (await run({ invoices })).items.map((i) => i.number);
  assertEquals(await order([inv("A", "2026-11-01"), inv("B", "2026-09-01")]), ["B", "A"]);
  assertEquals(await order([inv("A", "2026-11-01", { date: "2026-10-02" }), inv("B", "2026-11-01", { date: "2026-10-01" })]), ["B", "A"]);
  assertEquals(await order([inv("A", "2026-11-01"), inv("B", "2026-11-01")]), ["A", "B"]);
});

test("E17: one item and many", async () => {
  assertEquals((await run({ invoices: [inv("A", "2026-11-01")] })).items.length, 1);
  const invoices = Array.from({ length: 40 }, (_, i) => {
    const due = new Date(Date.UTC(2026, 10, 30) - [-5, 3, 40, 75, 120][i % 5] * 86_400_000).toISOString().slice(0, 10);
    return inv("N" + i, due, { amount: 10 + i * 1.37, disputed: i % 8 === 7 });
  });
  const r = await run({ invoices });
  const rows = r.items.reduce((s, i) => s + i.openMinor, 0);
  assertEquals(r.items.length, 40);
  assertEquals([rows, r.buckets.reduce((s, b) => s + b.amountMinor, 0)], [r.totalOpenMinor, r.totalOpenMinor]);
});

test("E18: overdue leaves out not-yet-due and disputed", async () => {
  const r = await run({
    invoices: [inv("A", "2026-12-01"), inv("B", "2026-11-15"), inv("C", "2026-07-01", { disputed: true })],
  });
  assertEquals([r.overdueMinor, r.totalOpenMinor], [10000, 30000]);
  assertEquals(await minors({ invoices: [inv("A", "2026-12-01")] }), [10000, 0, 0, 0, 0, 0]);
});

// ---- The node ---------------------------------------------------------------
const entry = COMPONENTS.find((c) => c.id === "open-items-ageing")!;
const previewData = entry.previewData as { account: Obj };

function documentOf(arrangement?: "across" | "down") {
  return template(
    <Document title="Statement">
      <OpenItemsAgeing arrangement={arrangement} />
    </Document>,
  );
}

function find(nodes: DocumentNode[], id: string): DocumentNode[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[] };
    return [...(own.id === id || own.id?.startsWith(id + "-") && id === "open-items-ageing-item" ? [node] : []), ...find(own.children ?? [], id)];
  });
}
function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { text?: string; children?: DocumentNode[] };
    return [...(own.text ? [own.text] : []), ...texts(own.children ?? [])];
  });
}

const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ openItemsAgeing: mod.run }));

async function resolved(data: Obj, arrangement?: "across" | "down") {
  const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof data, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
  return (await resolveDocument(published, state, { derivers: registry })).nodes;
}

test("previewData gives the figures the brief expects", async () => {
  const r = await run(previewData.account);
  assertEquals(r.items.map((i) => [i.number.slice(-4), i.age, i.open]), [
    ["4211", "In dispute", "£248,750.00"],
    ["4388", "103 days", "£786,420.50"],
    ["4502", "64 days", "£306,730.75"],
    ["4617", "31 days", "£97,640.00"],
    ["4733", "16 days", "£1,048,215.40"],
    ["4851", "Not yet due", "£655,300.00"],
  ]);
  assertEquals(r.items[3].due, "30 Oct 2026");
  assertEquals(r.buckets.map((b) => b.amountMinor), [65530000, 104821540, 9764000, 30673075, 78642050, 24875000]);
  assertEquals(r.buckets.map((b) => b.percent), [21, 33, 3, 10, 25, 8]);
  assertEquals([r.totalOpen, r.overdueTotal, r.unallocated, r.balance, r.balanceLabel], [
    "£3,143,056.65", "£2,239,006.65", "-£137,250.00", "£3,005,806.65", "Balance due",
  ]);
  assertEquals(r.basis.includes("counted as due 30 days after"), true);
});

test("E19: published, bundled and resolved, the rows are a loop over derived", async () => {
  for (const arrangement of ["across", "down"] as const) {
    const nodes = await resolved(previewData, arrangement);
    const direct = (await buildDocument(documentOf(arrangement), previewData, { dynamicMode: "placeholder" })).nodes;
    assertEquals(texts(nodes), texts(direct));
    const all = texts(nodes);
    const numbers = all.filter((t) => t.startsWith("INV-"));
    assertEquals(numbers.map((n) => n.slice(-4)), ["4211", "4388", "4502", "4617", "4733", "4851"]);
    assertEquals(numbers.includes("INV-2026-0004290"), false);
    for (const t of ["£1,048,215.40", "Open invoices", "£3,143,056.65", "Of which overdue", "Unallocated credit", "-£137,250.00", "Balance due"]) {
      assertEquals([t, all.includes(t)], [t, true]);
    }
    assertEquals(find(nodes, "open-items-ageing-strip").length, 1);
  }
  const empty = await resolved({ account: { ...base } });
  assertEquals(texts(empty), ["There are no open items on this account at 30 November 2026."]);
});

test("E20: arrangements differ, and a collapsed strip is one sentence in both", async () => {
  const build = async (a: "across" | "down", data: Obj) => (await buildDocument(documentOf(a), data, { dynamicMode: "placeholder" })).nodes;
  assertEquals(JSON.stringify(await build("across", previewData)) === JSON.stringify(await build("down", previewData)), false);
  const calm = { account: { ...base, invoices: [inv("A", "2026-12-01", { date: "2026-10-01" })] } };
  for (const a of ["across", "down"] as const) {
    const nodes = await build(a, calm);
    assertEquals(find(nodes, "open-items-ageing-strip").length, 0);
    assertEquals(texts(nodes).includes("Nothing on this account is overdue."), true);
  }
});
