import { test } from "node:test";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import { transformDocumentSource } from "docxcelerate/transform";
import { registryRoot } from "docxcelerate/registry/install";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";
import { COMPONENTS } from "docxcelerate/registry";

const nodeSource = "components/plan-change-proration/plan-change-proration.node.tsx";
const { PlanChangeProration } = await loadRegistryModule<
  typeof import("../registry/components/plan-change-proration/plan-change-proration.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/plan-change-proration.deriver.ts")
>("components/derivers/plan-change-proration.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ planChangeProration: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Out {
  outcome: string;
  showLines: boolean;
  periodLabel: string;
  days: number;
  lines: Array<{ description: string; period: string; amount: string; amountMinor: number }>;
  net: string;
  netAbs: string;
  netMinor: number;
}

const run = (sub: Obj): Out => mod.run([sub], {} as RuntimeState) as Out;
const nov = { currency: "GBP", period: { start: "2026-11-01", end: "2026-11-30" } };
const dec = { currency: "GBP", period: { start: "2026-12-01", end: "2026-12-31" } };
const chg = (at: string, from: number | string, to: number | string, more: Obj = {}) => ({
  at,
  from: { label: typeof from === "string" ? from : "Basic", paidMinor: typeof from === "number" ? from : 0 },
  to: { label: typeof to === "string" ? to : "Pro", periodAmountMinor: typeof to === "number" ? to : 0 },
  ...more,
});
const amounts = (o: Out) => o.lines.filter((l) => l.description).map((l) => l.amountMinor);

test("a basic upgrade", () => {
  const o = run({
    ...nov,
    changes: [{ at: "2026-11-14", from: { label: "Basic", paidMinor: 4900 }, to: { label: "Pro", periodAmountMinor: 9900 } }],
  });
  assertEquals(amounts(o), [-2777, 5610]);
  assertEquals([o.netMinor, o.net, o.outcome, o.days], [2833, "£28.33", "upgrade", 30]);
  assertEquals(o.lines[0].description, "Unused time on Basic from 14 November 2026");
  assertEquals(o.lines[1].description, "Remaining time on Pro from 14 November 2026");
  assertEquals(o.lines[0].period, "14 November 2026 to 30 November 2026");
  assertEquals(o.lines.length, 8);
});

test("days are counted: 31-day month, leap February, no rate drift", () => {
  const d = run({ ...dec, changes: [{ at: "2026-12-16", from: { label: "A", paidMinor: 3000 }, to: { label: "B", periodAmountMinor: 6000 } }] });
  assertEquals([d.days, ...amounts(d)], [31, -1548, 3097]);
  const l = run({
    currency: "GBP",
    period: { start: "2028-02-01", end: "2028-02-29" },
    changes: [chg("2028-02-15", 2900, 5800)],
  });
  assertEquals(l.days, 29);
  assertEquals(amounts(l), [-1500, 3000]);
  assertEquals(amounts(run({ ...dec, changes: [chg("2026-12-16", 1000, 2000)] }))[0], -516);
});

test("equal and opposite halves make a lateral change that prints nothing", async () => {
  const sub = {
    ...nov,
    changes: [{ at: "2026-11-16", from: { label: "A", paidMinor: 4901 }, to: { label: "B", periodAmountMinor: 4901 } }],
  };
  const o = run(sub);
  assertEquals([o.outcome, o.showLines, o.netMinor], ["lateral", false, 0]);
  assertEquals(o.lines.every((l) => l.description === "" && l.amountMinor === 0), true);
  assertEquals(o.lines.length, 8);
  const ids = await ids_of(sub);
  assertEquals(ids, ["plan-change-proration", "plan-change-proration-intro", "plan-change-proration-lateral"]);
});

test("a half unit rounds away from zero, in both directions", () => {
  assertEquals(amounts(run({ ...nov, changes: [chg("2026-11-16", 0, 4901)] })), [2451]);
  assertEquals(amounts(run({ ...nov, changes: [chg("2026-11-16", 4901, 0)] }))[0], -2451);
});

test("the credit is what was paid, and later changes chain", () => {
  assertEquals(amounts(run({ ...nov, changes: [chg("2026-11-11", 2400, 6000)] }))[0], -1600);
  const o = run({
    ...nov,
    changes: [
      chg("2026-11-11", "Basic", "Pro", { from: { label: "Basic", paidMinor: 3000 }, to: { label: "Pro", periodAmountMinor: 6000 } }),
      { at: "2026-11-21", from: { label: "Pro", paidMinor: 999 }, to: { label: "Business", periodAmountMinor: 9000 } },
    ],
  });
  assertEquals(amounts(o), [-2000, 4000, -2000, 3000]);
  assertEquals(o.netMinor, 3000);
  assertEquals(o.lines.filter((l) => l.description.startsWith("Unused")).map((l) => l.description.split(" ")[3]), ["Basic", "Pro"]);
});

test("same-day changes collapse, and a revert prints nothing", async () => {
  const basic = { label: "Basic", paidMinor: 3000 };
  const pro = { label: "Pro", periodAmountMinor: 6000 };
  const o = run({
    ...nov,
    changes: [
      { at: "2026-11-11", from: basic, to: pro },
      { at: "2026-11-11", from: { label: "Pro", paidMinor: 0 }, to: { label: "Business", periodAmountMinor: 9000 } },
    ],
  });
  assertEquals(amounts(o), [-2000, 6000]);
  assertEquals(o.lines[0].description, "Unused time on Basic from 11 November 2026");
  assertEquals(o.lines[1].description, "Remaining time on Business from 11 November 2026");
  assertEquals(o.netMinor, 4000);

  const revert = {
    ...nov,
    changes: [
      { at: "2026-11-11", from: basic, to: pro },
      { at: "2026-11-11", from: { label: "Pro", paidMinor: 0 }, to: { label: "Basic", periodAmountMinor: 3000 } },
    ],
  };
  assertEquals(run(revert).outcome, "none");
  assertEquals(await ids_of(revert), []);
});

test("changes outside the period, on renewal or metered print nothing", () => {
  for (const c of [chg("2026-11-01", 3000, 6000), chg("2026-10-20", 3000, 6000), chg("2026-12-01", 3000, 6000)]) {
    const o = run({ ...nov, changes: [c] });
    assertEquals([o.outcome, o.days, o.periodLabel, o.net, o.netMinor], ["none", 0, "", "", 0]);
  }
  const end = run({ ...nov, changes: [chg("2026-11-30", 3000, 6000)] });
  assertEquals([end.outcome, ...amounts(end)], ["upgrade", -100, 200]);
});

test("metered changes are skipped and do not chain", () => {
  const o = run({
    ...nov,
    changes: [
      chg("2026-11-11", 100, 200, { metered: true, item: "plan" }),
      chg("2026-11-21", 3000, 6000, { item: "plan" }),
    ],
  });
  assertEquals(amounts(o), [-1000, 2000]); // 10 of 30 days left
  assertEquals(o.lines[0].period, "21 November 2026 to 30 November 2026");
});

test("items chain apart, and lines are ordered by date", () => {
  const o = run({
    ...nov,
    changes: [
      chg("2026-11-21", 10000, 15000, { item: "seats" }),
      chg("2026-11-11", 3000, 6000, { item: "plan" }),
    ],
  });
  assertEquals(amounts(o), [-2000, 4000, -3333, 5000]);
});

test("a seat change names the quantity", () => {
  const o = run({
    ...nov,
    changes: [{
      at: "2026-11-11",
      from: { label: "Team", quantity: 10, paidMinor: 10000 },
      to: { label: "Team", quantity: 15, periodAmountMinor: 15000 },
    }],
  });
  assertEquals(amounts(o), [-6667, 10000]);
  assertEquals(o.lines[0].description, "Unused time on 10 × Team from 11 November 2026");
  assertStringIncludes(o.lines[1].description, "Remaining time on 15 × Team");
});

test("a net credit is balanced unless the policy refunds it", async () => {
  const sub = { ...nov, changes: [{ at: "2026-11-11", from: { label: "Pro", paidMinor: 6000 }, to: { label: "Basic", periodAmountMinor: 3000 } }] };
  const b = run(sub);
  assertEquals([...amounts(b), b.netMinor, b.outcome, b.net, b.netAbs], [-4000, 2000, -2000, "credit-balance", "-£20.00", "£20.00"]);
  assertEquals(run({ ...sub, refundPolicy: "refund" }).outcome, "credit-refund");
  assertEquals(run({ ...sub, refundPolicy: "balance" }).outcome, "credit-balance");

  const text = await texts_of(sub);
  assertStringIncludes(text, "£20.00");
  assertStringIncludes(text, "It is not refunded.");
  assertStringIncludes(text, "-£20.00");
  assertEquals((await ids_of({ ...sub, refundPolicy: "refund" })).includes("plan-change-proration-credit-refund"), true);
});

test("a zero-value line is dropped", () => {
  const o = run({ ...nov, changes: [{ at: "2026-11-11", from: { label: "Trial", paidMinor: 0 }, to: { label: "Pro", periodAmountMinor: 6000 } }] });
  assertEquals([amounts(o), o.outcome], [[4000], "upgrade"]);
});

test("no changes prints nothing and does not throw", async () => {
  for (const sub of [{ ...nov, changes: [] }, { ...nov }]) {
    assertEquals(run(sub).outcome, "none");
    assertEquals(await ids_of(sub), []);
  }
});

test("a zero-decimal currency has no minor units", () => {
  const o = run({ currency: "JPY", period: nov.period, changes: [chg("2026-11-11", 3000, 6000)] });
  assertEquals(o.lines[0].amount, new Intl.NumberFormat("en-GB", { style: "currency", currency: "JPY" }).format(-2000));
});

test("eight lines fit and nine throw", async () => {
  const four = ["2026-11-05", "2026-11-10", "2026-11-15", "2026-11-20"].map((at, i) => chg(at, `P${i}`, `P${i + 1}`, { from: { label: `P${i}`, paidMinor: 3000 }, to: { label: `P${i + 1}`, periodAmountMinor: 3000 + 1000 * (i + 1) } }));
  const sub = { ...nov, changes: four };
  assertEquals(amounts(run(sub)).length, 8);
  const ids = await ids_of(sub);
  for (let i = 1; i <= 8; i++) assertEquals(ids.includes(`plan-change-proration-line-${i}`), true);
  const five = [...four, chg("2026-11-25", "P4", "P5", { to: { label: "P5", periodAmountMinor: 9000 } })];
  await assertRejects(() => Promise.resolve().then(() => run({ ...nov, changes: five })), Error, "at most 8");
});

test("invalid input throws, naming the field", async () => {
  const ok = chg("2026-11-11", 3000, 6000);
  const bad: Array<[Obj, string]> = [
    [{ ...nov, currency: "XX", changes: [] }, "subscription.currency"],
    [{ currency: "GBP", period: { start: "2026-11-30", end: "2026-11-01" } }, "subscription.period.end"],
    [{ currency: "GBP", period: { start: "2026-11-01" } }, "subscription.period.end"],
    [{ ...nov, changes: [{ ...ok, at: "14/11/2026" }] }, "changes[0].at"],
    [{ ...nov, changes: [{ ...ok, from: { label: "A", paidMinor: 49.5 } }] }, "paidMinor"],
    [{ ...nov, changes: [{ ...ok, to: { label: "B", periodAmountMinor: -1 } }] }, "periodAmountMinor"],
    [{ ...nov, changes: [{ ...ok, to: { label: "", periodAmountMinor: 1 } }] }, "to.label"],
    [{ ...nov, changes: [{ ...ok, to: { label: "B", quantity: -1, periodAmountMinor: 1 } }] }, "quantity"],
  ];
  for (const [sub, field] of bad) {
    await assertRejects(() => Promise.resolve().then(() => run(sub)), Error, field);
  }
});

test("the catalog's preview data", async () => {
  const entry = COMPONENTS.find((c) => c.id === "plan-change-proration")!;
  const sub = (entry.previewData as { subscription: Obj }).subscription;
  const o = run(sub);
  assertEquals([amounts(o), o.netMinor, o.net, o.outcome], [
    [-927345, 1854765, -2626452, 3227419, -725777, 914487],
    1717097,
    "€17,170.97",
    "upgrade",
  ]);
  const via = await derivers.run(
    { name: "planChangeProration", output: "p", inputs: [{ type: "ref", ref: { scope: "data", path: "subscription" } }] } as never,
    state({ subscription: sub }),
  ) as Out;
  assertEquals(via, o);
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Subscription">
      <Section id="body" title="Body">
        <PlanChangeProration />
      </Section>
    </Document>,
  );

function state(data: Obj): RuntimeState {
  return { ctx: { ...data }, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } };
}

function collect(nodes: DocumentNode[], into: DocumentNode[] = []): DocumentNode[] {
  for (const n of nodes) {
    into.push(n);
    const children = (n as { children?: DocumentNode[] }).children;
    if (children) collect(children, into);
  }
  return into;
}

const textOf = (n: DocumentNode): string => JSON.stringify(n);

async function built(sub: Obj) {
  return await buildDocument(templateOf(), { subscription: sub }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);
}

async function ids_of(sub: Obj): Promise<string[]> {
  const body = (await built(sub)).nodes[0] as { children: DocumentNode[] };
  return collect(body.children).map((n) => n.id).filter((id) => !id.startsWith("table-cell"));
}

async function texts_of(sub: Obj): Promise<string> {
  return textOf((await built(sub)).nodes[0]);
}

test("a publish build carries 14 arms and one deriver call", async () => {
  assertEquals(branches, 14);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  const arms = all.filter((n) => (n as { when?: unknown }).when !== undefined);
  assertEquals(arms.length, 14);

  const section = all.find((n) => n.id === "plan-change-proration") as unknown as {
    derivers: unknown[];
  };
  assertEquals(section.derivers.length, 1);
  assertEquals(
    JSON.parse(JSON.stringify(section.derivers[0])).name,
    "planChangeProration",
  );
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).inputs, [
    { type: "ref", ref: { scope: "data", path: "subscription" } },
  ]);

  const entry = COMPONENTS.find((c) => c.id === "plan-change-proration")!;
  const preview = (entry.previewData as { subscription: Obj }).subscription;
  const cases: Obj[] = [
    preview,
    { ...nov, changes: [{ at: "2026-11-11", from: { label: "Pro", paidMinor: 6000 }, to: { label: "Basic", periodAmountMinor: 3000 } }] },
    { ...nov, changes: [chg("2026-11-16", 4901, 4901)] },
    { ...nov, changes: [] },
  ];
  for (const sub of cases) {
    const direct = await built(sub);
    const resolved = await resolveDocument(published, state({ subscription: sub }), { derivers });
    // A resolved tree keeps the conditions it was published with and numbers its
    // cells differently, so compare what a reader sees: the ids and the text.
    const seen = (nodes: DocumentNode[]) =>
      collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
  }
});
