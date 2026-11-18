import { test } from "node:test";
import { buildDocument, createDeriverBundle, createDeriverRegistryFromBundle, InMemoryDataProvider, resolveDocument } from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/fee-cap-status.deriver.ts")
>("components/derivers/fee-cap-status.deriver.ts");
const { FeeCapStatus } = await loadRegistryModule<
  typeof import("../registry/components/fee-cap-status/fee-cap-status.node.tsx")
>("components/fee-cap-status/fee-cap-status.node.tsx");

type Obj = Record<string, unknown>;
const base: Obj = { kind: "hard", amount: 10000, billedToDate: 0 };
const fc = (o: Obj = {}) => ({ ...base, ...o });

// deno-lint-ignore no-explicit-any
const run = (feeCap: Obj, thisPeriod: unknown): any => mod.run([feeCap, { thisPeriod }], {} as RuntimeState);
const bad = (feeCap: Obj, thisPeriod: unknown, field: string) =>
  assertRejects(() => Promise.resolve().then(() => run(feeCap, thisPeriod)), Error, field);

test("under: remaining shown, nothing written off", () => {
  const r = run(fc({ billedToDate: 4000 }), 2000);
  assertEquals(r.state, "under");
  assertEquals(r.heading, "Fee cap");
  assertEquals(r.status, "£6,000.00 of the agreed cap of £10,000.00 has been billed, including this invoice.");
  assertEquals([r.remainingLabel, r.remaining], ["Remaining under the cap", "£4,000.00"]);
  assertEquals([r.writtenOff, r.charged, r.withheld, r.payable], [null, null, null, null]);
  assertEquals(r.amounts, { charged: 2000, writtenOff: 0, withheld: 0, payable: 2000 });
});

test("a nil period is under, with no charged row", () => {
  const r = run(fc({ billedToDate: 4000 }), 0);
  assertEquals([r.state, r.charged], ["under", null]);
  assertStringIncludes(r.status, "£4,000.00 of");
});

test("thresholds: hard has no default, ceiling defaults to 85 with >=", () => {
  const a = run(fc({ noticeThreshold: 85, billedToDate: 8000 }), 1500);
  assertEquals([a.state, a.notice, a.remaining], ["approaching", null, "£500.00"]);
  assertStringIncludes(a.status, ": 95% of it.");
  assertEquals(run(fc({ billedToDate: 9000 }), 900).state, "under");
  const c = run(fc({ kind: "ceiling", billedToDate: 8000 }), 500);
  assertEquals(c.state, "approaching");
  assertEquals(c.notice, "This is our notice that fees billed have reached 85% of the ceiling.");
  assertEquals(c.heading, "Fee ceiling");
  assertEquals(run(fc({ kind: "ceiling", billedToDate: 8000 }), 499.99).state, "under");
});

test("fractional threshold and flooring", () => {
  const c = fc({ kind: "ceiling", noticeThreshold: 87.5, billedToDate: 8000 });
  assertEquals(run(c, 700).state, "under");
  const a = run(c, 750);
  assertEquals(a.state, "approaching");
  assertStringIncludes(a.notice, "87.5%");
  const f = run(fc({ kind: "ceiling", billedToDate: 8000 }), 1999.99);
  assertEquals(f.state, "approaching");
  assertStringIncludes(f.status, "99%");
});

test("reached exactly, and exceeded with write-off", () => {
  const r = run(fc({ billedToDate: 8000 }), 2000);
  assertEquals(r.state, "reached");
  assertEquals(r.status, "This invoice brings fees billed to the agreed cap of £10,000.00. Work beyond it will not be charged.");
  assertEquals(r.remaining, null);
  const e = run(fc({ billedToDate: 8000 }), 3000.5);
  assertEquals(e.state, "exceeded");
  assertEquals([e.charged, e.writtenOff, e.writtenOffLabel], ["£2,000.00", "£1,000.50", "Above the cap, not charged"]);
  assertEquals(e.status, "This invoice reaches the agreed cap of £10,000.00. Fees of £1,000.50 above it are not charged.");
  assertEquals(run(fc({ kind: "ceiling", billedToDate: 8000 }), 3000.5).writtenOffLabel, "Above the ceiling, not charged");
});

test("a cap already used up charges nil and never goes negative", () => {
  const e = run(fc({ billedToDate: 10000 }), 1200);
  assertEquals([e.state, e.charged, e.writtenOff], ["exceeded", "£0.00", "£1,200.00"]);
  assertEquals(e.status, "The agreed cap of £10,000.00 was reached on an earlier invoice. Fees of £1,200.00 for this period are above it and are not charged.");
  const r = run(fc({ billedToDate: 10500 }), 0);
  assertEquals(r.state, "reached");
  assertEquals(r.status, "The agreed cap of £10,000.00 was reached on an earlier invoice. Nothing further is charged.");
  assertEquals(r.amounts.charged, 0);
  const x = run(fc({ billedToDate: 10500 }), 300);
  assertEquals([x.state, x.charged, x.writtenOff], ["exceeded", "£0.00", "£300.00"]);
  assertEquals(x.amounts, { charged: 0, writtenOff: 300, withheld: 0, payable: 0 });
});

test("float noise does not turn reached into exceeded", () => {
  assertEquals(run({ kind: "hard", amount: 0.3, billedToDate: 0.1 }, 0.2).state, "reached");
});

test("estimate: charged in full, warning on an untold overrun", () => {
  const e = run(fc({ kind: "estimate", billedToDate: 9000 }), 2000);
  assertEquals([e.heading, e.state, e.charged, e.writtenOff, e.writtenOffLabel], ["Fee estimate", "exceeded", null, null, null]);
  assertEquals(e.amounts, { charged: 2000, writtenOff: 0, withheld: 0, payable: 2000 });
  assertEquals(e.status, "£11,000.00 has been billed, including this invoice, against our estimate of £10,000.00.");
  assertEquals(
    e.incomplete,
    "Fee estimate incomplete: fees billed to date, including this invoice, are £11,000.00, above our estimate of £10,000.00, and no revised estimate is recorded. An overrun the client was not told of can be challenged. Record the revised estimate and the date the client was told before this invoice is issued.",
  );
  const r = run(fc({ kind: "estimate", billedToDate: 8000 }), 2000);
  assertEquals([r.state, r.incomplete], ["reached", null]);
  assertStringIncludes(r.status, "only once a revised estimate has been given");
});

test("estimate with a revision is measured against it", () => {
  const revisedEstimate = { amount: 12500, notifiedOn: "2026-03-03" };
  const u = run(fc({ kind: "estimate", revisedEstimate, billedToDate: 9000 }), 2000);
  assertEquals([u.state, u.capLabel, u.incomplete], ["under", "Revised estimate", null]);
  assertEquals(u.revised, "Our estimate of £10,000.00 was revised to £12,500.00, and you were told on 3 March 2026.");
  assertEquals(u.remainingLabel, "Remaining under the revised estimate");
  const e = run(fc({ kind: "estimate", revisedEstimate, billedToDate: 12000 }), 1000);
  assertEquals(e.state, "exceeded");
  assertEquals(e.revised, u.revised);
  assertStringIncludes(e.incomplete, "above our revised estimate of £12,500.00, notified on 3 March 2026, and no later estimate");
});

test("revisedEstimate is ignored for a cap", () => {
  const r = run(fc({ revisedEstimate: { amount: 99999, notifiedOn: "2026-03-03" }, billedToDate: 4000 }), 2000);
  assertEquals([r.revised, r.capValue], [null, "£10,000.00"]);
});

test("withholding", () => {
  const w = (withheldToDate: number, extra: Obj = {}) =>
    run(fc({ billedToDate: 4000, withholding: { percent: 5, limit: 1000, withheldToDate }, ...extra }), 2000);
  const a = w(0);
  assertEquals([a.withheld, a.withheldLabel, a.payable], ["£100.00", "Retained (5%)", "£1,900.00"]);
  assertEquals(a.retention, "We have retained £100.00 (5%) of this invoice under the agreement: £100.00 retained to date, against a limit of £1,000.00.");
  const b = w(950);
  assertEquals([b.withheld, b.withheldLabel, b.payable], ["£50.00", "Retained (to the agreed limit)", "£1,950.00"]);
  assertEquals(b.retention, "We have retained £50.00 of this invoice under the agreement. That brings retention to its agreed limit of £1,000.00, and nothing more will be retained.");
  for (const done of [1000, 1200]) {
    const c = w(done);
    assertEquals([c.withheld, c.payable, c.withheldLabel], [null, null, null]);
    assertEquals(c.retention, "Retention has reached its agreed limit of £1,000.00, so nothing is retained from this invoice.");
    assertEquals(c.amounts.withheld, 0);
  }
  const n = run(fc({ billedToDate: 10000, withholding: { percent: 10, limit: 1000, withheldToDate: 0 } }), 500);
  assertEquals([n.withheld, n.retention], [null, null]);
  const h = run(fc({ withholding: { percent: 2.5, limit: 1000, withheldToDate: 0 } }), 12.34);
  assertEquals(h.withheld, "£0.31");
});

test("JPY has no decimals", () => {
  const r = run({ kind: "hard", amount: 1000000, billedToDate: 400000, currency: "JPY" }, 200000);
  assertEquals(r.status, "¥600,000 of the agreed cap of ¥1,000,000 has been billed, including this invoice.");
});

test("invalid input throws, naming the field", async () => {
  await bad(fc({ kind: "soft" }), 1, "feeCap.kind");
  await bad(fc({ currency: "XX" }), 1, "feeCap.currency");
  for (const amount of [0, -1, NaN, "100"]) await bad(fc({ amount }), 1, "feeCap.amount");
  for (const billedToDate of [undefined, -1]) await bad(fc({ billedToDate }), 1, "feeCap.billedToDate");
  for (const p of [undefined, -5]) await bad(fc(), p, "fees.thisPeriod");
  for (const noticeThreshold of [0, 100, 120]) await bad(fc({ noticeThreshold }), 1, "feeCap.noticeThreshold");
  const est = (revisedEstimate: Obj) => fc({ kind: "estimate", revisedEstimate });
  await bad(est({ amount: 0, notifiedOn: "2026-03-03" }), 1, "feeCap.revisedEstimate.amount");
  await bad(est({ amount: 5, notifiedOn: "next week" }), 1, "feeCap.revisedEstimate.notifiedOn");
  const wh = (withholding: Obj) => fc({ withholding });
  const ok = { percent: 5, limit: 100, withheldToDate: 0 };
  for (const percent of [0, 101]) await bad(wh({ ...ok, percent }), 1, "withholding.percent");
  await bad(wh({ ...ok, limit: 0 }), 1, "withholding.limit");
  for (const withheldToDate of [undefined, -1]) await bad(wh({ ...ok, withheldToDate }), 1, "withholding.withheldToDate");
});

// The node.
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ feeCapStatus: mod.run }));

function documentOf(arrangement?: "table" | "stacked") {
  return template(
    <Document title="Invoice">
      <FeeCapStatus arrangement={arrangement} />
    </Document>,
  );
}

function walk(nodes: DocumentNode[]): { ids: string[]; texts: string[] } {
  const ids: string[] = [];
  const texts: string[] = [];
  const go = (list: DocumentNode[]) =>
    list.forEach((node) => {
      const own = node as { id?: string; title?: string; text?: string; children?: DocumentNode[] };
      if (own.id?.startsWith("fee-cap")) ids.push(own.id);
      if (node.kind === "section" && own.title) texts.push(own.title);
      if (own.text) texts.push(own.text);
      go(own.children ?? []);
    });
  go(nodes);
  return { ids, texts };
}

const build = async (feeCap: Obj, thisPeriod: number, arrangement?: "table" | "stacked") =>
  walk((await buildDocument(documentOf(arrangement), { feeCap, fees: { thisPeriod } }, { dynamicMode: "placeholder" })).nodes);

test("the node prints the state arm and the figures", async () => {
  const u = await build(fc({ billedToDate: 4000 }), 2000);
  assertEquals(u.ids, ["fee-cap", "fee-cap-under", "fee-cap-figures", "fee-cap-amount", "fee-cap-billed", "fee-cap-period", "fee-cap-remaining"]);
  const a = await build(fc({ kind: "ceiling", billedToDate: 8000 }), 500);
  assertEquals(a.ids.filter((i) => /approaching|notice/.test(i)), ["fee-cap-approaching", "fee-cap-notice"]);
  const r = await build(fc({ billedToDate: 8000 }), 2000);
  assertEquals(r.ids.includes("fee-cap-reached"), true);
  const e = await build(fc({ billedToDate: 8000 }), 3000.5);
  assertEquals(e.ids.filter((i) => /exceeded|written|charged/.test(i)), ["fee-cap-exceeded", "fee-cap-written-off", "fee-cap-charged"]);
  const i = await build(fc({ kind: "estimate", billedToDate: 9000 }), 2000);
  assertEquals(i.ids.slice(0, 3), ["fee-cap", "fee-cap-incomplete", "fee-cap-exceeded"]);
  assertEquals(i.texts[0], "Fee estimate");
});

test("both arrangements print the same rows under the same conditions", async () => {
  const cases: [Obj, number][] = [
    [fc({ billedToDate: 4000 }), 2000],
    [fc({ billedToDate: 8000, withholding: { percent: 10, limit: 1000, withheldToDate: 0 } }), 3000],
    [fc({ kind: "estimate", billedToDate: 9000 }), 2000],
  ];
  for (const [feeCap, period] of cases) {
    const t = await build(feeCap, period, "table");
    const s = await build(feeCap, period, "stacked");
    assertEquals(t.ids.filter((i) => i !== "fee-cap-figures"), s.ids);
  }
  const s = await build(fc({ billedToDate: 4000 }), 2000, "stacked");
  assertEquals(s.texts.includes("Billed before this invoice: £4,000.00"), true);
});

const previewData = {
  feeCap: {
    kind: "ceiling",
    amount: 2500000,
    billedToDate: 2312480.55,
    noticeThreshold: 85,
    withholding: { percent: 10, limit: 250000, withheldToDate: 231248.06 },
    currency: "GBP",
  },
  fees: { thisPeriod: 418960.4 },
};

async function viaPublish(data: typeof previewData) {
  const published = await buildDocument(documentOf(), createPublishData() as typeof data, {
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
  return { published, resolved: walk((await resolveDocument(published, state, { derivers: registry })).nodes) };
}

test("published, every optional arm is conditional and the text is derived tokens", async () => {
  const { published } = await viaPublish(previewData);
  const json = JSON.stringify(published);
  for (const id of ["under", "approaching", "reached", "exceeded", "incomplete", "revised", "notice", "retention"]) {
    const at = json.indexOf(`"fee-cap-${id}"`);
    assertEquals(at > -1, true);
    assertStringIncludes(json.slice(Math.max(0, at - 600), at + 600), "when");
  }
  assertStringIncludes(json, "{{derived.");
  assertEquals(json.includes("{{data."), false);
});

test("published then resolved: the preview exceeds, and an under-cap document says only so", async () => {
  const { resolved } = await viaPublish(previewData);
  assertEquals(resolved.ids.filter((i) => /^fee-cap-(under|approaching|reached|exceeded|retention|notice|remaining)$/.test(i)), ["fee-cap-exceeded", "fee-cap-retention"]);
  assertEquals(resolved.texts.includes("This invoice reaches the agreed ceiling of £2,500,000.00. Fees of £231,440.95 above it are not charged."), true);
  assertEquals(resolved.texts.includes("£187,519.45"), true);
  assertEquals(resolved.texts.includes("Retained (to the agreed limit)"), true);
  assertEquals(resolved.texts.includes("£168,767.51"), true);
  assertEquals(resolved.texts.some((t) => t.startsWith("We have retained £18,751.94 of this invoice")), true);

  const under = await viaPublish({ feeCap: { kind: "hard", amount: 10000, billedToDate: 4000 }, fees: { thisPeriod: 2000 } } as never);
  assertEquals(under.resolved.ids.filter((i) => /^fee-cap-(under|approaching|reached|exceeded|incomplete|revised|notice|retention)$/.test(i)), ["fee-cap-under"]);
});
