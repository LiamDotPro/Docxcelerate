// deno-lint-ignore-file no-explicit-any
import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  dataRef,
  literalValue,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, ParagraphNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { COMPONENTS } from "../src/registry/catalog.ts";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/invoice-change-explainer.deriver.ts")
>("components/derivers/invoice-change-explainer.deriver.ts");
const { InvoiceChangeExplainer } = await loadRegistryModule<
  typeof import("../registry/components/invoice-change-explainer/invoice-change-explainer.node.tsx")
>("components/invoice-change-explainer/invoice-change-explainer.node.tsx");
const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ invoiceChangeExplainer: mod.run }),
);

type Obj = Record<string, any>;

const plan = { key: "plan", description: "Plan", kind: "recurring", quantity: 1, unitPrice: 42, amount: 42 };
const api = (q: number, amount: number, unitPrice = 0.00015) => ({
  key: "api", description: "Metered API calls", kind: "usage", quantity: q, unitPrice, amount,
});

function snap(lines: Obj[], extra: Obj = {}): Obj {
  const total = lines.reduce((t, l) => t + l.amount, 0);
  return { total: Math.round(total * 1e6) / 1e6, periodDays: 30, lines, ...extra };
}

function input(current: Obj, previous: Obj | null | undefined, extra: Obj = {}): Obj {
  return { currency: "GBP", periodLabel: "last month", current, previous, ...extra };
}

const prev0 = snap([plan, api(120000, 18)]);
const withApi = (q: number, amount: number, extra: Obj = {}) => snap([plan, api(q, amount)], extra);

const run = (data: Obj, locale?: string): any => mod.run([data, locale] as never, {} as RuntimeState);
// Through a bundle, which is how a deriver fails: shipped as source text.
const viaEngine = (data: Obj) =>
  registry.run(
    { name: "invoiceChangeExplainer", output: "x", inputs: [dataRef("change"), literalValue("en-GB")] } as never,
    { ctx: {}, derived: {}, dataProvider: new InMemoryDataProvider({ change: data }), aiClient: { generateParagraph: () => "" } },
  );

const thrown = async (data: Obj, text: string) => {
  await assertRejects(() => new Promise((_, reject) => { try { run(data); } catch (e) { reject(e); } }), Error, text);
};

test("1 a rise in usage is one quantity cause", async () => {
  const data = input(withApi(180000, 27), prev0);
  const r = run(data);
  assertEquals([r.speak, r.direction, r.omitted, r.causes.length, r.causes[0].kind, r.causes[0].effect], [true, "higher", 0, 1, "quantity", 900]);
  assertEquals(r.sentence, "Your total is £9.00 higher than last month, because Metered API calls rose from 120,000 to 180,000 (+£9.00).");
  assertEquals(await viaEngine(data), r);
});

test("2-4 below the floor, at the floor, and no change", () => {
  const small = run(input(withApi(122000, 18.3), prev0));
  assertEquals([small.speak, small.silentReason, small.sentence, small.causes], [false, "below-threshold", "", []]);
  assertEquals(run(input(withApi(128000, 19.2), prev0)).speak, true);
  assertEquals(run(input(prev0, prev0)).silentReason, "below-threshold");
  // Default pct 0.02 and absolute 1: a 1.05 rise is under 2% and so silent, and from nothing 0.99 is silent while 1.00 speaks.
  assertEquals(run(input(withApi(127000, 19.05), prev0)).speak, false);
  const bare = (a: number) => snap([{ key: "plan", description: "Plan", kind: "recurring", amount: a }]);
  assertEquals([0.99, 1].map((a) => run(input(bare(a), bare(0))).speak), [false, true]);
});

test("5-7 silence: first invoice, credit note, currency change", () => {
  const cur = withApi(180000, 27);
  assertEquals(run(input(cur, undefined)).silentReason, "first-invoice");
  assertEquals(run(input(cur, null)).silentReason, "first-invoice");
  assertEquals(run(input(cur, { ...prev0, documentType: "credit-note" })).silentReason, "credit-note");
  assertEquals(run(input({ ...cur, documentType: "credit-note" }, prev0)).silentReason, "credit-note");
  // an invalid previous is not validated once the currency differs
  assertEquals(run(input(cur, { currency: "EUR" })).silentReason, "currency-change");
});

test("8 a fall", () => {
  assertEquals(
    run(input(withApi(60000, 9), prev0)).sentence,
    "Your total is £9.00 lower than last month, because Metered API calls fell from 120,000 to 60,000 (-£9.00).",
  );
});

test("9 offsetting changes are both named", () => {
  const seats = (q: number, a: number) => ({ key: "seats", description: "Seats", kind: "recurring", quantity: q, unitPrice: 10, amount: a });
  const r = run(input(
    snap([plan, api(120000, 18), seats(20, 200), { key: "intro", description: "Introductory discount", kind: "discount", amount: -95 }]),
    snap([plan, api(120000, 18), seats(10, 100)]),
  ));
  assertEquals(r.sentence, "Your total is £5.00 higher than last month, because Seats rose from 10 to 20 (+£100.00) and Introductory discount was applied (-£95.00).");
});

test("10 the split is exact when the amount was rounded", () => {
  const l = (q: number, p: number, a: number) => ({ key: "k", description: "Item", kind: "recurring", quantity: q, unitPrice: p, amount: a });
  const r = run(input(snap([l(7, 0.347, 2.43)]), snap([l(3, 0.333, 1)]), { threshold: { absolute: 0 } }));
  // 133 is already 93% of the 143 movement, so the price share is omitted, not lost.
  assertEquals([r.causes.map((c: Obj) => [c.kind, c.effect]), r.omitted], [[["quantity", 133]], 1]);
});

test("11-13 period length", () => {
  const p = (q: number, a: number, days?: number) => snap([api(q, a, 0.0015)], { periodDays: days });
  const flat = run(input(p(124000, 186, 31), p(120000, 180, 30)));
  assertEquals(flat.causes.map((c: Obj) => [c.kind, c.effect]), [["period-length", 600]]);
  assertEquals(flat.sentence, "Your total is £6.00 higher than last month, because this period had 31 days against 30 (+£6.00).");
  const rising = run(input(p(155000, 232.5, 31), p(120000, 180, 30)));
  assertEquals(rising.sentence, "Your total is £52.50 higher than last month, because Metered API calls rose from 120,000 to 155,000 (+£46.50) and this period had 31 days against 30 (+£6.00).");
  for (const [c, q] of [[undefined, 30], [31, undefined]] as const) {
    const r = run(input(p(155000, 232.5, c), p(120000, 180, q)));
    assertEquals(r.causes.map((x: Obj) => [x.kind, x.effect]), [["quantity", 5250]]);
  }
});

test("11b an unchanged quantity is all price, whatever the period", () => {
  const p = (price: number, q: number, days: number) => snap([{ ...plan, unitPrice: price, amount: price }, api(q, q * 0.00015)], { periodDays: days });
  const r = run(input(p(60, 120000, 31), p(42, 120000, 30)));
  assertEquals(r.causes.map((c: Obj) => [c.kind, c.effect]), [["price", 1800]]);
  assertEquals(r.sentence, "Your total is £18.00 higher than last month, because the price of Plan went from £42.00 to £60.00 (+£18.00).");
});

test("11c a quantity that rose but whose effect fell says changed", () => {
  const p = (q: number, days: number) => snap([api(q, q * 0.0015, 0.0015)], { periodDays: days });
  const r = run(input(p(121000, 31), p(120000, 30), { threshold: { absolute: 0, pct: 0 } }));
  assertEquals(r.causes.find((c: Obj) => c.kind === "quantity").effect, -450);
  assertStringIncludes(r.sentence, "Metered API calls changed from 120,000 to 121,000 (-£4.50)");
});

test("14 a balance brought forward is not a change", () => {
  const base = run(input(withApi(180000, 27), prev0));
  assertEquals(run(input(withApi(180000, 27, { broughtForward: 250, total: 319 }), prev0)), base);
});

test("15 an unreconciled invoice throws", async () => {
  await thrown(input(withApi(180000, 27, { total: 319 }), prev0), "lines do not add up");
});

test("16 tax splits into base and rate, and the causes sum to the difference", () => {
  const seats = (q: number) => ({ key: "seats", description: "Seats", kind: "recurring", quantity: q, unitPrice: 10, amount: q * 10 });
  const vat = (rate: number, a: number) => ({ key: "vat", description: "VAT", kind: "tax", taxRate: rate, amount: a });
  const r = run(input(snap([seats(12), vat(0.2, 24)]), snap([seats(10), vat(0.175, 17.5)])));
  assertEquals(
    r.sentence,
    "Your total is £26.50 higher than last month, because Seats rose from 10 to 12 (+£20.00), VAT rose with the charges it is worked out on (+£3.50) and VAT changed from 17.5% to 20% (+£3.00).",
  );
  assertEquals(r.causes.reduce((t: number, c: Obj) => t + c.effect, 0), 2650);
});

test("17 many items are capped at three", () => {
  const line = (n: string, a: number) => ({ key: n, description: "Line " + n, kind: "recurring", amount: a });
  const now = [150, 140, 130, 120, 110, 105].map((a, i) => line("ABCDEF"[i], a));
  const r = run(input(snap(now), snap("ABCDEF".split("").map((n) => line(n, 100)))));
  assertEquals(r.omitted, 3);
  assertEquals(
    r.sentence,
    "Your total is £155.00 higher than last month, mostly because Line A rose from £100.00 to £150.00 (+£50.00), Line B rose from £100.00 to £140.00 (+£40.00) and Line C rose from £100.00 to £130.00 (+£30.00), with 3 smaller changes making up the rest.",
  );
});

test("18-21 added, removed, one-off and discount lines", () => {
  const l = (key: string, description: string, kind: string, amount: number) => ({ key, description, kind, amount });
  const say = (cur: Obj[], prev: Obj[]) => run(input(snap(cur), snap(prev))).sentence;
  assertEquals(
    say([plan, l("p", "Premier support", "recurring", 50)], [plan, l("l", "Legacy support", "recurring", 30)]),
    "Your total is £20.00 higher than last month, because Premier support was added (+£50.00) and Legacy support is no longer charged (-£30.00).",
  );
  assertEquals(
    say([plan], [plan, l("s", "Setup fee", "one-off", 25)]),
    "Your total is £25.00 lower than last month, because Setup fee from last month is not repeated (-£25.00).",
  );
  assertStringIncludes(say([plan, l("d", "Data export", "one-off", 15)], [plan]), "Data export is a one-off charge on this invoice (+£15.00)");
  assertStringIncludes(say([plan], [plan, l("w", "Welcome discount", "discount", -10)]), "because Welcome discount has ended (+£10.00).");
});

test("22-23 zero previous, and lines with no quantity or price", () => {
  const zero = (q: number, p: number, a: number) => ({ key: "plan", description: "Plan", kind: "recurring", quantity: q, unitPrice: p, amount: a });
  assertEquals(
    run(input(snap([zero(1, 42, 42)]), snap([zero(1, 0, 0)]))).sentence,
    "Your total is £42.00 higher than last month, because the price of Plan went from £0.00 to £42.00 (+£42.00).",
  );
  const bare = (a: number) => ({ key: "plan", description: "Plan", kind: "recurring", amount: a });
  assertStringIncludes(run(input(snap([bare(50)]), snap([bare(42)]))).sentence, "Plan rose from £42.00 to £50.00 (+£8.00)");
});

test("24 defaults for the label and threshold", () => {
  const r = run({ currency: "GBP", current: withApi(180000, 27), previous: prev0 });
  assertStringIncludes(r.sentence, "than your last invoice,");
  assertEquals(run({ currency: "GBP", current: withApi(122000, 18.3), previous: prev0, threshold: {} }).speak, false);
});

test("25 ties keep generation order", () => {
  const l = (key: string, a: number) => ({ key, description: key.toUpperCase(), kind: "recurring", amount: a });
  const r = run(input(snap([l("x", 20), l("y", 10), l("z", 1)]), snap([l("x", 10), l("y", 20)]), { threshold: { absolute: 0 } }));
  assertEquals(r.causes.map((c: Obj) => c.key), ["x", "y"]);
});

test("26 an amount a penny off quantity x price is all quantity", () => {
  const l = (q: number, a: number) => ({ key: "k", description: "Item", kind: "recurring", quantity: q, unitPrice: 0.333, amount: a });
  const r = run(input(snap([l(300, 99.99)]), snap([l(100, 33.3)]), { threshold: { absolute: 0 } }));
  assertEquals([r.causes.map((c: Obj) => [c.kind, c.effect]), r.omitted], [[["quantity", 6669]], 0]);
});

test("27 a currency with no minor units", () => {
  const l = (a: number) => ({ key: "k", description: "Item", kind: "recurring", amount: a });
  const r = run(input(snap([l(1500)]), snap([l(1000)]), { currency: "JPY" }));
  assertEquals(r.causes[0].effect, 500);
  assertStringIncludes(r.sentence, "¥500");
});

test("28 bad data throws, naming the field", async () => {
  const cur = withApi(180000, 27);
  const dup = snap([plan, { ...plan }]);
  await thrown(input(dup, prev0), "duplicated");
  await thrown(input(snap([{ ...plan, key: "" }]), prev0), "key");
  await thrown(input(snap([{ ...plan, kind: "weird" }]), prev0), "kind");
  await thrown(input(cur, prev0, { currency: "XX" }), "currency");
  await thrown(input(cur, prev0, { threshold: { pct: 1 } }), "threshold.pct");
  await thrown(input(cur, prev0, { threshold: { absolute: -1 } }), "threshold.absolute");
  await thrown(input({ ...cur, periodDays: 0 }, prev0), "periodDays");
  await thrown(input(snap([{ ...plan, amount: NaN }], { total: 1 }), prev0), "amount");
  await thrown(input(snap([{ ...plan, quantity: Infinity }]), prev0), "quantity");
  await thrown({ ...input(cur, prev0), current: undefined }, "current");
});

test("29 every cause is recomputed, and the chosen ones plus the rest equal the difference", () => {
  const { invoiceChange: ic } = COMPONENTS.find((c) => c.id === "invoice-change-explainer")!.previewData as { invoiceChange: Obj };
  const r = run(ic);
  // Each line alone, so the cap of three causes cannot hide one.
  const alone = (key: string) => {
    const side = (x: Obj) => {
      const lines = x.lines.filter((l: Obj) => l.key === key);
      return { total: lines.reduce((t: number, l: Obj) => t + l.amount, 0), periodDays: x.periodDays, lines };
    };
    return run({ currency: "GBP", current: side(ic.current), previous: side(ic.previous), threshold: { absolute: 0, pct: 0 } }).causes
      .map((c: Obj) => [c.kind, c.effect]);
  };
  const all = ["platform", "api-calls", "seats", "loyalty", "migration", "support", "vat"].flatMap(alone);
  assertEquals(all, [
    ["quantity", 2283_67], ["period-length", 351_33], ["quantity", 12420_00], ["price", 5075_00], ["discount-ended", 2220_00],
    ["one-off-ended", -7500_00], ["new", 3250_00], ["tax-base", 3620_00],
  ]);
  const chosen = r.causes.map((c: Obj) => [c.kind, c.effect]);
  const omitted = all.filter(([k, e]: [string, number]) => !chosen.some(([ck, ce]: [string, number]) => ck === k && ce === e));
  assertEquals(r.omitted, omitted.length);
  assertEquals(chosen.concat(omitted).reduce((t: number, [, e]: [string, number]) => t + e, 0), 21720_00);
});

// The node, published and resolved by an engine that only has the bundle.
function documentOf() {
  return template(<Document title="Invoice"><InvoiceChangeExplainer /></Document>);
}

function paragraphs(nodes: DocumentNode[]): ParagraphNode[] {
  return nodes.flatMap((n) => {
    const own = n as { children?: DocumentNode[] };
    return [...(n.kind === "paragraph" ? [n as ParagraphNode] : []), ...paragraphs(own.children ?? [])];
  });
}

test("publishing keeps one dynamic paragraph behind one branch", async () => {
  const published = await buildDocument(documentOf(), createPublishData() as never, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const text = JSON.stringify(published);
  assertStringIncludes(text, '"invoice-change-explainer"');
  assertStringIncludes(text, "invoiceChangeExplainer.speak");
  assertStringIncludes(text, "{{derived.invoiceChangeExplainer.facts}}");
  assertEquals(paragraphs(published.nodes).map((p) => [p.id, p.mode]), [["invoice-change-explainer", "dynamic"]]);
  assertEquals((paragraphs(published.nodes)[0] as Obj).when, {
    type: "truthy",
    ref: { scope: "derived", path: "invoiceChangeExplainer.speak" },
  });
  assertEquals(((paragraphs(published.nodes)[0] as Obj).derivers as Obj[]).map((d) => d.name), ["invoiceChangeExplainer"]);
});

async function resolveWith(data: Obj, captured: Obj[]) {
  const published = await buildDocument(documentOf(), createPublishData() as never, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: (spec: never) => (captured.push(spec as Obj), "ok") },
  };
  return await resolveDocument(published, state, { derivers: registry });
}

test("resolving speaks with the facts, and is silent when it barely differs", async () => {
  const captured: Obj[] = [];
  const spoken = await resolveWith({ invoiceChange: input(withApi(180000, 27), prev0) }, captured);
  assertEquals(paragraphs(spoken.nodes).map((p) => p.id), ["invoice-change-explainer"]);
  assertStringIncludes(
    captured[0].prompts.find((p: Obj) => p.kind === "info").text,
    "Metered API calls rose from 120,000 to 180,000 (+£9.00)",
  );
  const quiet = await resolveWith({ invoiceChange: input(withApi(122000, 18.3), prev0) }, []);
  assertEquals(paragraphs(quiet.nodes).length, 0);
});

test("a local build shows the sentence for the preview data", async () => {
  const preview = COMPONENTS.find((c) => c.id === "invoice-change-explainer")!.previewData;
  const built = await buildDocument(documentOf(), preview as never, { dynamicMode: "placeholder" });
  const [p] = paragraphs(built.nodes);
  assertEquals(p.text, run((preview as { invoiceChange: Obj }).invoiceChange).sentence);
  assertStringIncludes(p.text ?? "", "Your total is £21,720.00 higher than last month, mostly because Named analyst seats");
});
