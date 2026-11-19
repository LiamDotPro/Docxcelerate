import { test } from "node:test";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  dataRef,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import { transformDocumentSource } from "docxcelerate/transform";
import { registryRoot } from "docxcelerate/registry/install";
import { COMPONENTS } from "docxcelerate/registry";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/contingency-settlement-statement/contingency-settlement-statement.node.tsx";
const { ContingencySettlementStatement } = await loadRegistryModule<
  typeof import("../registry/components/contingency-settlement-statement/contingency-settlement-statement.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/contingency-settlement-statement.deriver.ts")
>("components/derivers/contingency-settlement-statement.deriver.ts");
// Through a bundle, so the test proves the function survives being shipped as source text.
const derivers = await createDeriverRegistryFromBundle(
  createDeriverBundle({ contingencySettlementStatement: mod.run }),
);
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Out {
  gross: string;
  rows: Array<{ label: string; amount: string }>;
  netLabel: string;
  net: string;
  outcome: string;
  method: string;
  otherBase: string;
  cap: string;
  feeMinor: number;
  netMinor: number;
}

function state(data: Obj): RuntimeState {
  return { ctx: { ...data }, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } };
}

async function run(settlement: Obj): Promise<Out> {
  return await derivers.run(
    { name: "contingencySettlementStatement", output: "s", inputs: [dataRef("settlement")] } as never,
    state({ settlement }),
  ) as Out;
}

const printed = (o: Out) => o.rows.filter((r) => r.label);
const amounts = (o: Out) => printed(o).map((r) => r.amount);
const NY = [
  { upTo: 250000, percent: 30 },
  { upTo: 500000, percent: 25 },
  { upTo: 1000000, percent: 20 },
  { upTo: 1250000, percent: 15 },
  { percent: 10 },
];
const us = (more: Obj = {}): Obj => ({
  regime: "contingency",
  currency: "USD",
  grossRecovery: 100000,
  base: "gross",
  schedule: [{ percent: 30 }],
  ...more,
});
const dba = (more: Obj = {}): Obj => ({
  regime: "DBA",
  currency: "GBP",
  grossRecovery: 200000,
  base: "gross",
  claimType: "personal-injury",
  damagesBreakdown: { general: 60000, pastLoss: 40000, futureLoss: 100000 },
  schedule: [{ percent: 25 }],
  vatPercent: 20,
  counselFees: 6000,
  recoverableCostsPaid: 15000,
  expenses: [{ label: "Expert", amount: 2000 }],
  liens: [{ holder: "CRU", amount: 10000, cru: true }],
  ...more,
});
const cfa = (more: Obj = {}): Obj => ({
  regime: "CFA",
  currency: "GBP",
  grossRecovery: 50000,
  baseCosts: 20000,
  stage: "trial",
  schedule: [{ stage: "trial", percent: 100 }, { stage: "pre-suit", percent: 25 }],
  claimType: "personal-injury",
  damagesBreakdown: { general: 30000, pastLoss: 10000 },
  vatPercent: 20,
  counselFees: 3000,
  recoverableCostsPaid: 23000,
  ...more,
});
const preview = (COMPONENTS.find((c) => c.id === "contingency-settlement-statement")!.previewData as { settlement: Obj })
  .settlement;

test("the preview data: the NY cap binds", async () => {
  const o = await run(preview);
  assertEquals(o.gross, "$4,850,000.00");
  assertEquals(o.rows.length, 16);
  assertEquals(printed(o).length, 9);
  assertEquals(printed(o)[5], { label: "Contingent fee (statutory maximum)", amount: "-$606,228.20" });
  assertEquals([o.netLabel, o.net, o.outcome], ["Net to you", "$3,595,016.25", "remittance"]);
  assertEquals(o.otherBase, "");
  assertStringIncludes(o.cap, "$606,228.20");
  assertStringIncludes(o.cap, "$4,562,281.95");
  assertStringIncludes(o.cap, "above it, so the statutory maximum is charged");
  assertStringIncludes(o.method, "the recovery after expenses ($4,562,281.95)");
  assertStringIncludes(o.method, "33.33% of");
});

test("tiered rates are marginal, not flat", async () => {
  const o = await run(us({ grossRecovery: 1000000, schedule: NY }));
  assertEquals(amounts(o)[0], "-$237,500.00");
  assertStringIncludes(o.method, "band by band");
  assertStringIncludes(o.method, "30% of the first $250,000.00 ($75,000.00)");
  assertStringIncludes(o.method, "25% of the next $250,000.00 ($62,500.00)");
  assertStringIncludes(o.method, "20% of the remaining $500,000.00 ($100,000.00)");
  const open = await run(us({ grossRecovery: 2000000, schedule: NY }));
  assertStringIncludes(open.method, "10% of the amount above $1,250,000.00 ($75,000.00)");
});

test("gross versus net names the other base", async () => {
  const gross = await run(us({ grossRecovery: 100000, base: "gross", schedule: [{ percent: 100 / 3 }], expenses: [{ label: "E", amount: 10000 }] }));
  assertEquals([gross.feeMinor, gross.net], [3333333, "$56,666.67"]);
  assertStringIncludes(gross.otherBase, "recovery after expenses instead");
  assertStringIncludes(gross.otherBase, "$30,000.00");
  const net = await run(us({ grossRecovery: 100000, base: "net", schedule: [{ percent: 100 / 3 }], expenses: [{ label: "E", amount: 10000 }] }));
  assertEquals([net.feeMinor, net.net], [3000000, "$60,000.00"]);
  assertStringIncludes(net.otherBase, "gross recovery instead");
  assertStringIncludes(net.otherBase, "$33,333.33");
  assertEquals((await run(us())).otherBase, "");
});

test("a lien comes off after the fee, unless marked beforeFee", async () => {
  const lien = { holder: "Aetna", amount: 20000 };
  const after = await run(us({ liens: [lien] }));
  assertEquals([after.feeMinor, after.net], [3000000, "$50,000.00"]);
  const before = await run(us({ liens: [{ ...lien, beforeFee: true }] }));
  assertEquals([before.feeMinor, before.net], [2400000, "$56,000.00"]);
  assertEquals(printed(before)[0].label, "Aetna, paid before the fee");
  assertEquals(printed(before)[1].label, "Contingent fee");
  assertStringIncludes(before.method, "and the liens paid before the fee");
});

test("a reduced lien says what it was reduced from", async () => {
  const o = await run(us({ liens: [{ holder: "Aetna", amount: 20000, reducedAmount: 12500 }] }));
  assertEquals(printed(o)[1], { label: "Aetna (reduced from $20,000.00)", amount: "-$12,500.00" });
  const waived = await run(us({ liens: [{ holder: "Aetna", amount: 20000, reducedAmount: 0 }] }));
  assertEquals(printed(waived)[1].amount, "$0.00");
  await assertRejects(() => run(us({ liens: [{ holder: "A", amount: 1, reducedAmount: 2 }] })), Error, "settlement.liens[0].reducedAmount");
});

test("the stage picks the rate and is never guessed", async () => {
  const schedule = [{ stage: "pre-suit", percent: 25 }, { stage: "filed", percent: 100 / 3 }];
  const o = await run(us({ stage: "pre-suit", schedule }));
  assertEquals(amounts(o)[0], "-$25,000.00");
  assertStringIncludes(o.method, "25% of the gross recovery ($100,000.00), the rate the agreement sets before suit was filed.");
  await assertRejects(() => run(us({ schedule })), Error, "stage");
  await assertRejects(() => run(us({ stage: "appeal", schedule })), Error, 'no rate in settlement.schedule for stage "appeal"');
  const fallback = await run(us({ stage: "appeal", schedule: [...schedule, { percent: 40 }] }));
  assertEquals(amounts(fallback)[0], "-$40,000.00");
});

test("California caps at 25% before filing and 33% after", async () => {
  const ca = (stage?: string) => us({ grossRecovery: 300000, base: "net", schedule: [{ percent: 100 / 3 }], statutoryCap: "ca-medical-negligence", stage });
  const pre = await run(ca("pre-suit"));
  assertEquals(printed(pre)[0], { label: "Contingent fee (statutory maximum)", amount: "-$75,000.00" });
  assertStringIncludes(pre.cap, "settled before filing, so the limit is $75,000.00");
  assertStringIncludes(pre.cap, "The agreement's fee of $100,000.00 is above it");
  const filed = await run(ca("filed"));
  assertEquals(amounts(filed)[0], "-$99,000.00");
  assertStringIncludes(filed.cap, "was filed, so the limit is $99,000.00");
  await assertRejects(() => run(ca()), Error, "stage");
});

test("a NY cap above the agreed fee changes nothing", async () => {
  const o = await run(us({ grossRecovery: 500000, schedule: [{ percent: 10 }], statutoryCap: "ny-medical-malpractice" }));
  assertEquals(printed(o)[0], { label: "Contingent fee", amount: "-$50,000.00" });
  assertStringIncludes(o.cap, "within that limit, so it is charged");
});

test("DBA personal injury: cap, VAT, counsel and costs credit", async () => {
  const o = await run(dba());
  assertEquals(printed(o).map((r) => r.label), [
    "Payment under the agreement",
    "VAT at 20% on the payment",
    "Counsel's fees",
    "Reduced to the statutory maximum",
    "Costs paid by the other side, credited",
    "Expert",
    "CRU",
  ]);
  assertEquals(amounts(o), ["-£50,000.00", "-£10,000.00", "-£6,000.00", "£43,500.00", "£15,000.00", "-£2,000.00", "-£10,000.00"]);
  assertEquals(o.net, "£180,500.00");
  assertStringIncludes(o.cap, "£90,000.00");
  assertStringIncludes(o.cap, "£22,500.00");
  assertStringIncludes(o.cap, "future loss of £100,000.00 is excluded");
  assertStringIncludes(o.cap, "is above it, so it is reduced to £22,500.00");
  assertStringIncludes(o.method, "The payment is 25% of the gross recovery (£200,000.00)");
  const capped = await run(dba({ recoverableCostsPaid: 30000 }));
  assertEquals(printed(capped)[4].amount, "£22,500.00");
});

test("DBA employment and other claims", async () => {
  const emp = await run(dba({ grossRecovery: 80000, claimType: "employment", schedule: [{ percent: 40 }], counselFees: undefined, recoverableCostsPaid: undefined, expenses: [], liens: [] }));
  assertEquals(amounts(emp).slice(0, 3), ["-£32,000.00", "-£6,400.00", "£10,400.00"]);
  assertStringIncludes(emp.cap, "35% of the sums recovered (£80,000.00)");
  assertStringIncludes(emp.cap, "The payment of £38,400.00 is above it");
  const other = await run(dba({ grossRecovery: 80000, claimType: "other", schedule: [{ percent: 40 }], counselFees: undefined, recoverableCostsPaid: undefined, expenses: [], liens: [] }));
  assertStringIncludes(other.cap, "50% of the sums recovered (£80,000.00), which is £40,000.00");
  assertStringIncludes(other.cap, "is within it");
});

test("CFA personal injury caps the success fee before VAT", async () => {
  const o = await run(cfa());
  assertEquals(amounts(o), ["-£20,000.00", "-£20,000.00", "£10,000.00", "-£6,000.00", "-£3,000.00", "£23,000.00"]);
  assertEquals(printed(o)[1].label, "Success fee (100% of base costs)");
  assertEquals(o.net, "£34,000.00");
  assertStringIncludes(o.method, "The success fee is 100% of base costs of £20,000.00, the rate the agreement sets at trial.");
  assertStringIncludes(o.cap, "Conditional Fee Agreements Order 2013");
  await assertRejects(() => run(cfa({ schedule: [{ stage: "trial", percent: 120 }] })), Error, "percent");
  await assertRejects(() => run(cfa({ schedule: [{ stage: "trial", upTo: 5, percent: 50 }] })), Error, "upTo");
  await assertRejects(() => run(cfa({ baseCosts: undefined })), Error, "settlement.baseCosts");
});

test("a shortfall carries the agreement's terms, or warns", async () => {
  const s = us({ grossRecovery: 10000, schedule: [{ percent: 100 / 3 }], expenses: [{ label: "E", amount: 9000 }], liens: [{ holder: "L", amount: 2000 }] });
  const o = await run(s);
  assertEquals([o.outcome, o.netLabel, o.net], ["shortfall", "Shortfall", "$4,333.33"]);
  const exact = await run(us({ grossRecovery: 100, schedule: [{ percent: 0 }], expenses: [{ label: "E", amount: 100 }] }));
  assertEquals([exact.outcome, exact.net], ["remittance", "$0.00"]);

  const terms = { ...s, shortfallTerms: "The firm waives the balance." };
  const ids = async (settlement: Obj) => (await built(settlement)).map((n) => n.id);
  assertEquals((await ids(terms)).includes("contingency-settlement-shortfall"), true);
  assertEquals((await ids(terms)).includes("contingency-settlement-shortfall-incomplete"), false);
  assertEquals((await ids(s)).includes("contingency-settlement-shortfall-incomplete"), true);
  assertEquals((await ids(us())).filter((id) => id.includes("shortfall")), []);
  assertStringIncludes(JSON.stringify(await built(terms)), "leaves a shortfall of $4,333.33. The firm waives the balance.");
});

test("a zero recovery never prints a negative zero", async () => {
  const o = await run(us({ grossRecovery: 0, base: "net", expenses: [{ label: "E", amount: 1500 }] }));
  assertEquals(amounts(o), ["-$1,500.00", "$0.00"]);
  assertStringIncludes(o.method, "nothing to work the fee out on");
  assertEquals([o.outcome, o.net], ["shortfall", "$1,500.00"]);
});

test("invalid input throws, naming the field", async () => {
  const bad: Array<[Obj, string]> = [
    [us({ grossRecovery: -1 }), "settlement.grossRecovery"],
    [us({ expenses: [{ label: "a", amount: 1 }, { label: "b", amount: 1 }, { label: "c", amount: -1 }] }), "settlement.expenses[2].amount must be a finite amount of 0 or more"],
    [us({ schedule: [{ percent: 101 }] }), "percent"],
    [us({ schedule: [{ percent: 10 }, { percent: 20 }] }), "open band"],
    [us({ schedule: [{ upTo: 10, percent: 10 }] }), "open band"],
    [us({ schedule: [{ upTo: 10, percent: 10 }, { upTo: 10, percent: 20 }, { percent: 5 }] }), "same upTo"],
    [dba({ statutoryCap: "ny-medical-malpractice" }), "statutoryCap"],
    [dba({ base: undefined }), "settlement.base"],
    [dba({ damagesBreakdown: undefined }), "damagesBreakdown"],
    [dba({ damagesBreakdown: { general: 150000, pastLoss: 60000 } }), "above settlement.grossRecovery"],
    [us({ regime: "other" }), "settlement.regime"],
    [us({ currency: "XYZ" }), "settlement.currency"],
  ];
  for (const [settlement, message] of bad) await assertRejects(() => run(settlement), Error, message);
});

test("missing optional fields print only the fee; 16 rows build, 17 do not", async () => {
  const o = await run(us());
  assertEquals(printed(o).length, 1);
  assertEquals([o.otherBase, o.cap], ["", ""]);
  assertEquals(printed(await run(us({ expenses: [{ label: "E", amount: 5 }] }))).length, 2);
  const lines = (n: number) => Array.from({ length: n }, (_, i) => ({ label: "E" + i, amount: 1 }));
  assertEquals(printed(await run(us({ expenses: lines(15) }))).length, 16);
  await assertRejects(() => run(us({ expenses: lines(16) })), Error, "17 lines; this component prints at most 16");
});

test("rounding is half up, once, in minor units", async () => {
  assertEquals(amounts(await run(us({ grossRecovery: 1000.05 })))[0], "-$300.02");
  const yen = await run(us({ currency: "JPY", grossRecovery: 100001 }));
  assertStringIncludes(amounts(yen)[0], "30,000");
  assertEquals(yen.feeMinor, 30000);
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Settlement">
      <Section id="body" title="Body">
        <ContingencySettlementStatement />
      </Section>
    </Document>,
  );

function collect(nodes: DocumentNode[], into: DocumentNode[] = []): DocumentNode[] {
  for (const n of nodes) {
    into.push(n);
    const children = (n as { children?: DocumentNode[] }).children;
    if (children) collect(children, into);
  }
  return into;
}

async function built(settlement: Obj): Promise<DocumentNode[]> {
  const b = await buildDocument(templateOf(), { settlement }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);
  return collect(b.nodes);
}

test("a publish build carries 20 arms and resolves to what a direct build prints", async () => {
  assertEquals(branches, 20);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  const arms = all.filter((n) => (n as { when?: unknown }).when !== undefined);
  assertEquals(arms.length, 20);
  assertEquals(arms.filter((n) => /contingency-settlement-line-\d+$/.test(n.id)).length, 16);
  assertStringIncludes(JSON.stringify(published.nodes), "{{derived.");

  const shortfall = us({ grossRecovery: 10000, expenses: [{ label: "E", amount: 9000 }], liens: [{ holder: "L", amount: 2000 }], shortfallTerms: "Firm waives." });
  const seen = (nodes: DocumentNode[]) => nodes.map((n) => [n.kind, (n as { text?: string }).text]);
  for (const settlement of [preview, shortfall, { ...shortfall, shortfallTerms: undefined }]) {
    const resolved = await resolveDocument(published, state({ settlement }), { derivers });
    assertEquals(seen(collect(resolved.nodes)), seen(await built(settlement)));
  }
  const text = JSON.stringify((await resolveDocument(published, state({ settlement: preview }), { derivers })).nodes);
  assertStringIncludes(text, "$3,595,016.25");
  assertStringIncludes(text, "Contingent fee (statutory maximum)");
});
