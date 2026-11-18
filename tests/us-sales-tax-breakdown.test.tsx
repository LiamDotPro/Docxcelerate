import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  dataRef,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/us-sales-tax-breakdown.deriver.ts")
>("components/derivers/us-sales-tax-breakdown.deriver.ts");
const { UsSalesTaxBreakdown } = await loadRegistryModule<
  typeof import("../registry/components/us-sales-tax-breakdown/us-sales-tax-breakdown.node.tsx")
>("components/us-sales-tax-breakdown/us-sales-tax-breakdown.node.tsx");
const { COMPONENTS } = await import("../src/registry/catalog.ts");

type Obj = Record<string, unknown>;
type Result = ReturnType<NonNullable<typeof mod.__types>>;
// Called the way the engine calls it: inputs in an array.
const run = (salesTax: unknown) => mod.run([salesTax], undefined as never) as Result;
const TX = [
  { level: "state", name: "Texas", rate: "0.0625" },
  { level: "county", name: "Travis County", rate: "0.005" },
  { level: "city", name: "Austin", rate: "0.01" },
  { level: "special", name: "Capital Metro", rate: "0.005" },
];
const line = (amountMinor: number, extra: Obj = {}) => ({ description: "Item", amountMinor, taxable: true, ...extra });
const taxes = (r: Result) => r.rows.map((row) => row.taxMinor);

test("rounding residual goes to the state row", () => {
  const a = run({ jurisdictions: TX, lines: [line(4699)] });
  assertEquals(a.rows.map((r) => r.taxableMinor), [4699, 4699, 4699, 4699]);
  assertEquals(taxes(a), [295, 23, 47, 23]);
  assertEquals([a.total.tax, a.total.taxable, a.total.rate], ["$3.88", "$46.99", "8.25%"]);
  assertEquals(a.rows.map((r) => r.rate), ["6.25%", "0.50%", "1.00%", "0.50%"]);
  const b = run({ jurisdictions: TX, lines: [line(1101)] });
  assertEquals(taxes(b), [68, 6, 11, 6]);
  assertEquals(b.total.tax, "$0.91");
  assertEquals(taxes(run({ jurisdictions: TX, lines: [line(1999)] })), [125, 10, 20, 10]);
});

test("rows always add up to the total, for 1..10000", () => {
  for (let n = 1; n <= 10000; n++) {
    const r = run({ jurisdictions: TX, lines: [line(n)] });
    assertEquals(taxes(r).reduce((a, b) => a + b, 0), r.total.taxMinor);
  }
});

test("a refund mirrors the sale", () => {
  const r = run({ jurisdictions: TX, lines: [line(-4699)] });
  assertEquals(taxes(r), [-295, -23, -47, -23]);
  assertEquals([r.total.taxMinor, r.rows[0].tax], [-388, "-$2.95"]);
});

test("no goods", () => {
  for (const s of [{ jurisdictions: TX, lines: [] }, { jurisdictions: TX }]) {
    const r = run(s);
    assertEquals(r.rows.length, 4);
    assertEquals(r.rows.every((x) => x.taxable === "$0.00" && x.tax === "$0.00"), true);
    assertEquals([r.total.tax, r.total.taxable, r.total.rate], ["$0.00", "$0.00", "8.25%"]);
    assertEquals([r.exemptions, r.shippingWarning], [[], ""]);
  }
});

test("per-item cap", () => {
  const r = run({
    jurisdictions: [
      { level: "state", name: "Tennessee", rate: "0.07", taxesShipping: true },
      { level: "county", name: "Shelby County", rate: "0.0225", perItemCapMinor: 160000 },
    ],
    lines: [line(249900), line(4000)],
  });
  assertEquals(r.rows.map((x) => x.taxableMinor), [253900, 164000]);
  assertEquals(taxes(r), [17773, 3690]);
  assertEquals([r.total.taxMinor, r.total.taxable, r.total.rate, r.exemptions], [21463, "", "", []]);
});

const CLOTHING = ["clothing"];
test("category exemption", () => {
  const r = run({
    jurisdictions: [
      { level: "state", name: "State", rate: "0.04", exemptCategories: CLOTHING },
      { level: "city", name: "City", rate: "0.045", exemptCategories: CLOTHING },
      { level: "special", name: "Transit District", rate: "0.00375" },
    ],
    lines: [line(8999, { category: "clothing", description: "Jacket" }), line(15000)],
  });
  assertEquals(r.rows.map((x) => x.taxableMinor), [15000, 15000, 23999]);
  assertEquals([...taxes(r), r.total.taxMinor], [600, 675, 90, 1365]);
  assertEquals(r.exemptions, [{ reason: "Not taxed by State and City", items: "Jacket ($89.99)" }]);
});

test("category exempt everywhere", () => {
  const r = run({
    jurisdictions: [{ level: "state", name: "S", rate: "0.04", exemptCategories: CLOTHING }, { level: "city", name: "C", rate: "0.01", exemptCategories: CLOTHING }],
    lines: [line(8999, { category: "clothing" })],
  });
  assertEquals(r.rows.map((x) => x.taxableMinor), [0, 0]);
  assertEquals(r.exemptions[0].reason, "Exempt from sales tax (clothing)");
});

const shippingCase = (shipping: Obj) => run({
  jurisdictions: [
    { level: "state", name: "State", rate: "0.06", taxesShipping: false },
    { level: "city", name: "City", rate: "0.02", exemptCategories: CLOTHING },
  ],
  lines: [line(3000, { category: "clothing" }), line(7000), { description: "Warranty", amountMinor: 2500, taxable: false }],
  shipping,
});

test("shipping where delivery is not taxed, then combined, then folded in", () => {
  const a = shippingCase({ amountMinor: 1000, separatelyStated: true });
  assertEquals(a.rows.map((r) => r.taxableMinor), [10000, 7000]);
  assertEquals([...taxes(a), a.total.taxMinor, a.shippingWarning], [600, 140, 740, ""]);
  assertEquals(a.exemptions.some((e) => e.reason === "Not taxable"), true);

  const b = shippingCase({ amountMinor: 1000, separatelyStated: true, combinedWithHandling: true });
  assertEquals(b.rows.map((r) => r.taxableMinor), [10800, 7560]);
  assertEquals([...taxes(b), b.total.taxMinor], [648, 151, 799]);
  assertStringIncludes(b.shippingWarning, "charged together as $10.00. State does not tax");

  const c = shippingCase({ amountMinor: 1000, separatelyStated: false });
  assertEquals([c.rows.map((r) => r.taxableMinor), c.total.taxMinor, c.shippingWarning], [[10800, 7560], 799, ""]);
});

test("shipping follows the goods", () => {
  const st = [{ level: "state", name: "S", rate: "0.06", taxesShipping: true }];
  const ex = run({ jurisdictions: st, lines: [{ description: "x", amountMinor: 5000, taxable: false }], shipping: { amountMinor: 995 } });
  assertEquals([ex.rows[0].taxableMinor, ex.total.taxMinor, ex.shippingWarning], [0, 0, ""]);
  const t = run({ jurisdictions: st, lines: [line(5000)], shipping: { amountMinor: 995 } });
  assertEquals([t.rows[0].taxableMinor, t.total.taxMinor], [5995, 360]);
  const c = run({ jurisdictions: st, lines: [line(-5000)], shipping: { amountMinor: -995 } });
  assertEquals([c.rows[0].taxableMinor, c.total.taxMinor], [-5995, -360]);
});

test("zero state rate, and a six-place rate", () => {
  const r = run({
    jurisdictions: [
      { level: "state", name: "Alaska", rate: "0" },
      { level: "city", name: "Juneau", rate: "0.05" },
      { level: "special", name: "X", rate: "0.0125" },
    ],
    lines: [line(104)],
  });
  assertEquals([taxes(r), r.rows[0].rate], [[0, 6, 1], "0.00%"]);
  const p = run({ jurisdictions: [{ level: "state", name: "S", rate: "0.047125" }], lines: [line(100000)] });
  assertEquals([p.total.taxMinor, p.rows[0].rate], [4713, "4.7125%"]);
});

test("marketplace", () => {
  const r = run({ collectedBy: "marketplace", marketplaceName: "Northwind Marketplace LLC" });
  assertStringIncludes(r.marketplace, "collected and remitted by Northwind Marketplace LLC");
  assertEquals([r.rows, r.exemptions, r.shippingWarning], [[], [], ""]);
  assertStringIncludes(run({ collectedBy: "marketplace", marketplaceName: " " }).marketplace, "by the marketplace facilitator");
});

test("many items", () => {
  const r = run({ jurisdictions: TX, lines: Array.from({ length: 50 }, () => line(1999)) });
  assertEquals(r.rows[0].taxableMinor, 99950);
  assertEquals(taxes(r).reduce((a, b) => a + b, 0), r.total.taxMinor);
});

test("bad input throws, naming the field", async () => {
  const ok = { jurisdictions: TX, lines: [line(100)] };
  const cases: Array<[unknown, string]> = [
    [undefined, "salesTax is required"],
    [{ ...ok, collectedBy: "both" }, "collectedBy"],
    [{ ...ok, jurisdictions: [] }, "jurisdictions"],
    [{ ...ok, jurisdictions: TX.slice(1) }, "state"],
    [{ ...ok, jurisdictions: [...TX, TX[0]] }, "state"],
    [{ ...ok, jurisdictions: [{ ...TX[0], level: "district" }] }, "jurisdictions[0].level"],
    [{ ...ok, jurisdictions: [{ ...TX[0], rate: 0.0625 }] }, "jurisdictions[0].rate"],
    [{ ...ok, jurisdictions: [{ ...TX[0], rate: "6.25" }] }, "jurisdictions[0].rate"],
    [{ ...ok, jurisdictions: [{ ...TX[0], rate: "0.0000001" }] }, "jurisdictions[0].rate"],
    [{ ...ok, jurisdictions: [{ ...TX[0], rate: ".0625" }] }, "jurisdictions[0].rate"],
    [{ ...ok, jurisdictions: [{ ...TX[0], perItemCapMinor: 0 }] }, "perItemCapMinor"],
    [{ ...ok, jurisdictions: [{ ...TX[0], name: " " }] }, "jurisdictions[0].name"],
    [{ ...ok, lines: [line(46.99)] }, "lines[0].amountMinor"],
    [{ ...ok, lines: [{ description: "x", amountMinor: 1 }] }, "lines[0].taxable"],
    [{ ...ok, lines: [{ amountMinor: 1, taxable: false }] }, "lines[0].description"],
    [{ ...ok, currency: "XXXX" }, "currency"],
    [{ ...ok, shipping: { amountMinor: 1.5 } }, "shipping.amountMinor"],
  ];
  for (const [input, message] of cases) await assertRejects(() => Promise.resolve().then(() => run(input)), Error, message);
});

test("exemption grouping", () => {
  const nt = (d: string, a: number, exemptReason?: string) => ({ description: d, amountMinor: a, taxable: false, exemptReason });
  const r = run({ jurisdictions: TX, lines: [nt("A", 100, "Food"), nt("B", 200, "Labour"), nt("C", 300, "Food")] });
  assertEquals(r.exemptions, [{ reason: "Food", items: "A ($1.00); C ($3.00)" }, { reason: "Labour", items: "B ($2.00)" }]);
});

test("catalog preview data", () => {
  const entry = COMPONENTS.find((c) => c.id === "us-sales-tax-breakdown");
  const r = run((entry as unknown as { previewData: Obj }).previewData.salesTax);
  assertEquals(r.rows.map((x) => [x.taxable, x.rate, x.tax]), [
    ["$25,228.24", "2.90%", "$731.62"],
    ["$25,228.24", "0.25%", "$63.07"],
    ["$1,534,473.72", "3.00%", "$46,034.21"],
    ["$25,228.24", "1.00%", "$252.28"],
    ["$25,228.24", "0.10%", "$25.23"],
  ]);
  assertEquals([r.total.tax, r.total.taxable, r.total.rate], ["$47,106.41", "", ""]);
  assertEquals(r.exemptions.map((e) => e.reason), [
    "Not taxed by Colorado, Arapahoe County, Regional Transportation District, and Scientific and Cultural Facilities District",
    "Optional service contract, not tangible personal property",
  ]);
  assertStringIncludes(r.shippingWarning, "$24,850.00");
});

// The node, published and resolved by an engine that only has the bundle.
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ usSalesTaxBreakdown: mod.run }));
const combined = {
  jurisdictions: [{ level: "state", name: "State", rate: "0.06", taxesShipping: false }, { level: "city", name: "City", rate: "0.02", exemptCategories: CLOTHING }],
  lines: [line(3000, { category: "clothing" }), line(7000), { description: "Warranty", amountMinor: 2500, taxable: false }],
  shipping: { amountMinor: 1000, separatelyStated: true, combinedWithHandling: true },
};
const market = { collectedBy: "marketplace", marketplaceName: "Northwind Marketplace LLC" };

const documentOf = () => template(<Document title="Tax"><UsSalesTaxBreakdown /></Document>);
const publishOptions = { branchMode: "publish", deriverMode: "preserve", dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } } as const;

// Every named id with its text, depth first. A cell's generated id is numbered by position and
// differs between a loop's passes and a local build, and an empty cell is written by one and not the other, so only
// its words are compared.
function flat(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((n) => {
    const own = n as { id?: string; text?: string; children?: DocumentNode[] };
    const generated = n.kind === "tableCell" || (own.id ?? "").startsWith("table-cell");
    return [...(generated ? (!own.text ? [] : [`cell:${own.text}`]) : [`${own.id ?? n.kind}:${own.text ?? ""}`]), ...flat(own.children ?? [])];
  });
}

test("published node carries one invocation, two loops and three conditions", async () => {
  const published = await buildDocument(documentOf(), createPublishData() as never, publishOptions);
  const json = JSON.stringify(published);
  assertEquals((json.match(/"name":"usSalesTaxBreakdown"/g) ?? []).length, 1);
  assertStringIncludes(json, JSON.stringify(dataRef("salesTax")));
  for (const path of ["rows", "exemptions", "marketplace", "shippingWarning"]) {
    assertStringIncludes(json, `"scope":"derived","path":"usSalesTaxBreakdown.${path}"`);
  }
  assertEquals((json.match(/"kind":"repeat"/g) ?? []).length, 2);
});

for (const [label, salesTax] of [["combined shipping", combined], ["marketplace", market]] as const) {
  test(`publishing then resolving matches building directly (${label})`, async () => {
    const data = { salesTax };
    const published = await buildDocument(documentOf(), createPublishData() as typeof data, publishOptions);
    const state: RuntimeState = {
      ctx: { ...data },
      derived: {},
      dataProvider: new InMemoryDataProvider(data),
      aiClient: { generateParagraph: () => "" },
    };
    const resolved = await resolveDocument(published, state, { derivers: registry });
    const direct = await buildDocument(documentOf(), data as never, { dynamicMode: "placeholder" });
    const ids = flat(resolved.nodes);

    assertEquals(ids, flat(direct.nodes));
    assertEquals(ids.some((s) => s.startsWith("us-sales-tax-marketplace")), label === "marketplace");
    assertEquals(ids.some((s) => s.startsWith("us-sales-tax-shipping-warning")), label !== "marketplace");
    assertEquals(ids.some((s) => s.startsWith("us-sales-tax-total")), label !== "marketplace");
  });
}
