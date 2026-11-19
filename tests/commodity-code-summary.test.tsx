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
import { COMPONENTS } from "docxcelerate/registry";
import type { CommodityCodeSummaryLines } from "../registry/components/arrangements/commodity-code-summary/lines.ts";
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/commodity-code-summary.deriver.ts")
>("components/derivers/commodity-code-summary.deriver.ts");
const { CommodityCodeSummary } = await loadRegistryModule<
  typeof import("../registry/components/commodity-code-summary/commodity-code-summary.node.tsx")
>("components/commodity-code-summary/commodity-code-summary.node.tsx");

type Obj = Record<string, unknown>;
type Result = CommodityCodeSummaryLines & { empty: boolean };

const line = (over: Obj = {}): Obj => ({
  ref: "R1",
  tariffCode: "0901210000",
  origin: "CO",
  quantity: 1,
  netMass: { value: 1, unit: "kg" },
  valueMinor: 100,
  ...over,
});
const customs = (lines: Obj[], over: Obj = {}): Obj => ({ currency: "GBP", destinationCodeLength: 10, lines, ...over });
const run = (c: Obj) => mod.run([c], {} as RuntimeState) as Result;
const one = (over: Obj = {}, c: Obj = {}) => run(customs([line(over)], c));

test("no lines, or lines missing, is empty", () => {
  for (const c of [customs([]), { currency: "GBP", destinationCodeLength: 10 }]) {
    const r = run(c);
    assertEquals(r.empty, true);
    assertEquals(r.lines, []);
    assertEquals(r.groups, []);
    assertEquals(r.totalValue, "");
  }
});

test("one good line is one group and needs no attention", () => {
  const r = one();
  assertEquals(r.lines.length, 1);
  assertEquals(r.groups.length, 1);
  assertEquals(r.groups[0].lineCount, "1 line");
  assertEquals(r.attention, "");
  assertEquals(r.lines[0].warning, "");
  assertEquals(r.lines[0].item, "R1");
});

test("a line without a ref prints its position, and a description follows a dash", () => {
  assertEquals(one({ ref: undefined, description: "Coffee" }).lines[0].item, "Line 1 — Coffee");
});

test("tariff codes", () => {
  const w = (tariffCode: unknown, L = 10) => one({ tariffCode }, { destinationCodeLength: L }).lines[0];
  const num = w(901210000);
  assertEquals(num.warning.includes("leading zero may have been lost"), true);
  assertEquals(num.warning.includes("has 9 digits; this destination needs 10."), true);
  assertEquals(w("0901 21 00", 8).warning, "");
  assertEquals(w("0901 21 00", 8).tariffCode, "09012100");
  assertEquals(w("090121").warning, "The tariff code has 6 digits; this destination needs 10.");
  for (const empty of ["", undefined]) {
    const r = w(empty);
    assertEquals(r.warning, "No tariff code: this line cannot be declared.");
    assertEquals(r.tariffCode, "None");
  }
  assertEquals(w("0901-21").warning, "“0901-21” is not a tariff code.");
  assertEquals(w("0901210000", 8).warning, "");
});

test("a missing code groups as None and sorts last", () => {
  const r = run(customs([line({ tariffCode: "" }), line({ tariffCode: "9999999999" })]));
  assertEquals(r.groups.map((g) => g.tariffCode), ["9999999999", "None"]);
});

test("origin", () => {
  assertEquals(one({ origin: " co " }).lines[0].origin, "CO");
  assertEquals(one({ origin: " co " }).lines[0].warning, "");
  const none = one({ origin: "" }).lines[0];
  assertEquals(none.warning, "No country of origin.");
  assertEquals(none.origin, "None");
  assertEquals(one({ origin: "COL" }).lines[0].warning, "“COL” is not a two-letter country code.");
});

test("groups split on preference and origin, and sort deterministically", () => {
  const p = run(customs([line({ preference: true }), line({ preference: undefined })]));
  assertEquals(p.groups.map((g) => g.preference), ["No", "Yes"]);
  const o = run(customs([line({ origin: "CO" }), line({ origin: "BR" })]));
  assertEquals(o.groups.map((g) => g.origin), ["BR", "CO"]);
});

test("pounds are converted, and each line rounds for display only", () => {
  assertEquals(one({ netMass: { value: 1, unit: "lb" } }).lines[0].netMass, "0.454");
  assertEquals(one({ netMass: { value: 2.5, unit: "lb" } }).lines[0].netMass, "1.134");
});

test("mass is rounded once per group, and the total is the sum of the groups", () => {
  const tiny = { netMass: { value: 0.0004, unit: "kg" } };
  const together = run(customs([line(tiny), line(tiny)]));
  assertEquals(together.lines.map((l) => l.netMass), ["0.000", "0.000"]);
  assertEquals(together.groups[0].netMass, "0.001");
  const apart = run(customs([line(tiny), line({ ...tiny, origin: "BR" })]));
  assertEquals(apart.groups.map((g) => g.netMass), ["0.000", "0.000"]);
  assertEquals(apart.totalNetMass, "0.000");
  const noise = run(customs([line({ netMass: { value: 0.1, unit: "kg" } }), line({ netMass: { value: 0.2, unit: "kg" } })]));
  assertEquals(noise.groups[0].netMass, "0.300");
  const zero = one({ netMass: { value: 0, unit: "kg" } });
  assertEquals(zero.lines[0].netMass, "0.000");
  assertEquals(zero.lines[0].warning, "");
});

test("bad numbers throw, naming the field", async () => {
  const bad: Array<[Obj, string]> = [
    [{ netMass: { value: -1, unit: "kg" } }, "customs.lines[0].netMass.value"],
    [{ netMass: { value: 1, unit: "oz" } }, "customs.lines[0].netMass.unit"],
    [{ valueMinor: -1 }, "customs.lines[0].valueMinor"],
    [{ valueMinor: 1.5 }, "customs.lines[0].valueMinor"],
    [{ quantity: NaN }, "customs.lines[0].quantity"],
    [{ supplementaryUnit: { code: "", quantity: 1 } }, "customs.lines[0].supplementaryUnit.code"],
    [{ supplementaryUnit: { code: "l", quantity: -1 } }, "customs.lines[0].supplementaryUnit.quantity"],
  ];
  for (const [over, path] of bad) await assertRejects(async () => await one(over), Error, path);
  await assertRejects(async () => await one({}, { currency: "XXZ" }), Error, "customs.currency");
  await assertRejects(async () => await one({}, { destinationCodeLength: 9 }), Error, "customs.destinationCodeLength");
  await assertRejects(async () => await one({}, { grossMassKg: -1 }), Error, "customs.grossMassKg");
});

test("a currency with no minor units prints whole", () => {
  assertEquals(one({ valueMinor: 120000 }, { currency: "JPY" }).lines[0].value, "JP¥120,000");
});

test("supplementary units total only when every line agrees", () => {
  const pa = (quantity: number) => ({ supplementaryUnit: { code: "pa", quantity } });
  const both = run(customs([line(pa(2400)), line(pa(1500))]));
  assertEquals(both.groups[0].supplementary, "3,900 pa");
  assertEquals(both.groups[0].warning, "");
  const mixed = run(customs([line(pa(2400)), line()]));
  assertEquals(mixed.groups[0].supplementary, "");
  assertEquals(
    mixed.groups[0].warning,
    "The lines in this group give different supplementary units, so none is totalled.",
  );
  const unlike = run(customs([line(pa(1)), line({ supplementaryUnit: { code: "l", quantity: 1 } })]));
  assertEquals(unlike.groups[0].warning !== "", true);
  assertEquals(run(customs([line(), line()])).groups[0].warning, "");
});

test("gross mass", () => {
  const net = { netMass: { value: 10, unit: "kg" } };
  const missing = one(net);
  assertEquals([missing.grossMass, missing.grossWarning], ["", ""]);
  const equal = one(net, { grossMassKg: 10 });
  assertEquals([equal.grossMass, equal.grossWarning], ["10.000", ""]);
  assertEquals(
    one(net, { grossMassKg: 9.999 }).grossWarning,
    "The net mass of the lines, 10.000 kg, is more than the declared gross mass of 9.999 kg. One of them is wrong.",
  );
});

test("attention counts lines with a warning", () => {
  assertEquals(
    run(customs([line(), line(), line({ tariffCode: "090121" })])).attention,
    "1 of 3 lines needs attention before these goods can be declared.",
  );
  assertEquals(
    run(customs([line({ origin: "" }), line({ origin: "" })])).attention,
    "2 of 2 lines need attention before these goods can be declared.",
  );
});

test("200 lines across 12 groups add up exactly", () => {
  const lines = Array.from({ length: 200 }, (_, i) => line({ tariffCode: `09012100${10 + (i % 12)}`, valueMinor: 333 + i }));
  const r = run(customs(lines));
  assertEquals(r.lines.length, 200);
  assertEquals(r.groups.length, 12);
  const minor = (s: string) => Math.round(Number(s.replace(/[^\d.]/g, "")) * 100);
  assertEquals(r.groups.reduce((sum, g) => sum + minor(g.value), 0), minor(r.totalValue));
});

// The node.

const previewData = COMPONENTS.find((c: { id: string }) => c.id === "commodity-code-summary")!.previewData as { customs: Obj };

function documentOf(arrangement?: "lines-then-summary" | "summary-first") {
  return template(
    <Document title="Invoice">
      <CommodityCodeSummary arrangement={arrangement} />
    </Document>,
  );
}

function walk(nodes: DocumentNode[]): DocumentNode[] {
  return nodes.flatMap((node) => [node, ...walk((node as { children?: DocumentNode[] }).children ?? [])]);
}
const ids = (nodes: DocumentNode[]) => walk(nodes).map((n) => (n as { id?: string }).id).filter((id) => id?.startsWith("commodity-code-summary")) as string[];

const build = (data: Obj, arrangement?: "lines-then-summary" | "summary-first") =>
  buildDocument(documentOf(arrangement), data as never, { dynamicMode: "placeholder" });

test("an empty invoice has the sentence and no tables", async () => {
  const found = ids((await build({ customs: customs([]) })).nodes);
  assertEquals(found.includes("commodity-code-summary-none"), true);
  assertEquals(found.some((id) => id.endsWith("-lines") || id.endsWith("-groups")), false);
});

test("both arrangements print the same nodes, the summary first in one", async () => {
  const data = { customs: previewData.customs };
  const a = ids((await build(data)).nodes);
  const b = ids((await build(data, "summary-first")).nodes);
  assertEquals([...a].sort(), [...b].sort());
  assertEquals(a.indexOf("commodity-code-summary-lines") < a.indexOf("commodity-code-summary-groups"), true);
  assertEquals(b.indexOf("commodity-code-summary-groups") < b.indexOf("commodity-code-summary-lines"), true);
  assertEquals((await build(data)).nodes, (await build(data, "lines-then-summary")).nodes);
});

for (const arrangement of ["lines-then-summary", "summary-first"] as const) {
test(`published, the ${arrangement} tables loop over derived rows and resolve on the engine`, async () => {
  const bad = structuredClone(previewData.customs) as { lines: Obj[] };
  bad.lines[1].tariffCode = "090121";
  const data = { customs: bad };
  const published = await buildDocument(documentOf(arrangement), createPublishData() as never, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const repeats = walk(published.nodes).filter((n) => n.kind === "repeat") as unknown as Array<{ source: { scope: string } }>;
  assertEquals(repeats.length, 2);
  assertEquals(repeats.every((r) => r.source.scope === "derived"), true);

  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
  const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ commodityCodeSummary: mod.run }));
  const resolved = await resolveDocument(published, state, { derivers: registry });
  const found = walk(resolved.nodes);
  const count = (prefix: string) => found.filter((n) => ((n as { id?: string }).id ?? "").startsWith(prefix)).length;
  const exactly = (re: RegExp) => found.filter((n) => re.test((n as { id?: string }).id ?? "")).length;
  assertEquals(exactly(/^commodity-code-summary-line-\d+$/), 5);
  assertEquals(exactly(/^commodity-code-summary-group-\d+$/), 4);
  const text = JSON.stringify(resolved.nodes);
  assertEquals(text.includes("25,539.986"), true);
  assertEquals(count("commodity-code-summary-line-warning"), 1);
  const line1 = found.find((n) => (n as { id?: string }).id === "commodity-code-summary-line-1")!;
  assertEquals(
    walk([line1]).filter((n) => ((n as { id?: string }).id ?? "").startsWith("commodity-code-summary-line-warning")).length,
    1,
  );
});
}
