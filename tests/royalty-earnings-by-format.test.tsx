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
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { RoyaltyEarningsByFormat } = await loadRegistryModule<
  typeof import("../registry/components/royalty-earnings-by-format/royalty-earnings-by-format.node.tsx")
>("components/royalty-earnings-by-format/royalty-earnings-by-format.node.tsx");
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/royalty-earnings-by-format.deriver.ts")
>("components/derivers/royalty-earnings-by-format.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ royaltyEarningsByFormat: mod.run }));

// deno-lint-ignore no-explicit-any
type Obj = Record<string, any>;
interface Line {
  description: string;
  units: string;
  basis: string;
  rate: string;
  amount: string;
  memo: boolean;
  amountMinor: number;
}
interface Result {
  outcome: string;
  negative: boolean;
  hasMemo: boolean;
  lines: Line[];
  total: string;
  totalAbs: string;
  totalMinor: number;
  lifeToDate: string;
}

const run = (royalty: Obj): Result => mod.run([royalty], {} as RuntimeState) as unknown as Result;
const tiers = [{ fromUnit: 1, percent: 10 }, { fromUnit: 2501, percent: 12.5 }];
const fmt = (o: Obj = {}) => ({
  format: "hardback",
  basis: "list",
  listPriceMinor: 2000,
  tiers,
  lifeToDateUnitsBefore: 0,
  ...o,
});
const gbp = (formats: Obj[], more: Obj = {}) => ({ currency: "GBP", formats, ...more });
const amounts = (r: Result) => r.lines.map((l) => l.amountMinor);

test("an escalator crossed mid-period splits at the copy where the rate changes", () => {
  const r = run(gbp([fmt({ lifeToDateUnitsBefore: 2300, sales: [{ units: 500, discountPercent: 40 }] })]));
  assertEquals(r.lines.map((l) => [l.description, l.units, l.basis, l.rate, l.amount]), [
    ["Hardback: copies 1–2,500", "200", "£20.00 list", "10%", "£400.00"],
    ["Hardback: copies 2,501 and over", "300", "£20.00 list", "12.5%", "£750.00"],
  ]);
  assertEquals([r.total, r.totalMinor, r.lifeToDate], ["£1,150.00", 115000, "Hardback 2,800"]);
});

test("a threshold exactly at the boundary", () => {
  const at = (before: number) => run(gbp([fmt({ lifeToDateUnitsBefore: before, sales: [{ units: 100 }] })]));
  assertEquals(at(2400).lines.map((l) => [l.units, l.rate]), [["100", "10%"]]);
  assertEquals(at(2500).lines.map((l) => [l.units, l.rate]), [["100", "12.5%"]]);
});

test("two thresholds in one row", () => {
  const r = run(gbp([
    fmt({
      tiers: [...tiers, { fromUnit: 5001, percent: 15 }],
      lifeToDateUnitsBefore: 2000,
      sales: [{ units: 4000 }],
    }),
  ]));
  assertEquals(r.lines.map((l) => [l.units, l.rate]), [["500", "10%"], ["2,500", "12.5%"], ["1,000", "15%"]]);
});

test("each line is rounded, not the total", () => {
  const r = run(gbp([fmt({ listPriceMinor: 1099, tiers: [{ fromUnit: 1, percent: 7.5 }, { fromUnit: 4, percent: 12.5 }], sales: [{ units: 4 }] })]));
  assertEquals(amounts(r), [247, 137]); // exact 247.275 and 137.375
  assertEquals([r.totalMinor, r.total], [384, "£3.84"]);
});

test("high-discount fraction steps", () => {
  const hd = { thresholdPercent: 55, treatment: "fraction", fraction: { numerator: 4, denominator: 5 }, deeper: [{ thresholdPercent: 62.5, numerator: 3, denominator: 5 }] };
  const f = (sales: Obj[]) => fmt({ tiers: [{ fromUnit: 1, percent: 10 }], sales });
  const r = run(gbp([f([{ units: 100, discountPercent: 40 }, { units: 50, discountPercent: 55 }, { units: 20, discountPercent: 65 }])], { highDiscount: hd }));
  assertEquals(amounts(r), [20000, 8000, 2400]);
  assertEquals(r.lines.map((l) => l.rate), ["10%", "4/5 of 10%", "3/5 of 10%"]);
  assertStringIncludes(r.lines[1].description, "Hardback: copies 1 and over, high discount (55% or more)");
  assertStringIncludes(r.lines[2].description, "high discount (62.5% or more)");
  assertEquals(r.totalMinor, 30400);

  const edge = run(gbp([f([{ units: 100, discountPercent: 54.99 }])], { highDiscount: hd }));
  assertEquals(edge.lines.map((l) => l.rate), ["10%"]);

  const none = run(gbp([f([{ units: 100, discountPercent: 40 }, { units: 50, discountPercent: 55 }, { units: 20, discountPercent: 65 }])]));
  assertEquals(none.lines.length, 1);
  assertEquals(none.lines[0].amountMinor, 34000);
});

test("the net switch earns the full rate on net receipts", () => {
  const hd = { treatment: "net" };
  const f = (sale: Obj) => fmt({ tiers: [{ fromUnit: 1, percent: 10 }], sales: [sale] });
  const a = run(gbp([f({ units: 50, discountPercent: 60 })], { highDiscount: hd }));
  assertEquals(a.lines.map((l) => [l.amountMinor, l.rate, l.basis]), [[4000, "10% of net", "£400.00 net receipts"]]);
  assertEquals(a.lines[0].description, "Hardback: copies 1 and over, high discount (55% or more)");
  const b = run(gbp([f({ units: 50, discountPercent: 60, netReceiptsMinor: 38500 })], { highDiscount: hd }));
  assertEquals([b.lines[0].amountMinor, b.lines[0].basis], [3850, "£385.00 net receipts"]);
});

test("a net-basis format ignores high discount", () => {
  const r = run(gbp([{ format: "ebook", basis: "net", tiers: [{ fromUnit: 1, percent: 25 }], lifeToDateUnitsBefore: 0, sales: [{ units: 100, discountPercent: 70, netReceiptsMinor: 33333 }] }], { highDiscount: {} }));
  assertEquals(r.lines.map((l) => [l.amountMinor, l.rate, l.basis]), [[8333, "25%", "£333.33 net receipts"]]);
});

test("receipts are split across a threshold without rounding first", () => {
  const r = run(gbp([{ format: "audio", basis: "net", tiers: [{ fromUnit: 1, percent: 20 }, { fromUnit: 11, percent: 25 }], lifeToDateUnitsBefore: 0, sales: [{ units: 15, netReceiptsMinor: 1000 }] }]));
  assertEquals(amounts(r), [133, 83]);
  assertEquals(r.lines.map((l) => l.units), ["10", "5"]);
});

test("returns are charged at the top tier reached and never reduce the count", () => {
  const r = run(gbp([fmt({ lifeToDateUnitsBefore: 2300, sales: [{ units: 500, discountPercent: 40 }], returns: { units: 40 } })]));
  const ret = r.lines[2];
  assertEquals([ret.description, ret.units, ret.basis, ret.rate, ret.amountMinor], ["Hardback: returns", "-40", "£20.00 list", "12.5%", -10000]);
  assertEquals(r.lifeToDate, "Hardback 2,800");

  assertEquals(amounts(run(gbp([fmt({ lifeToDateUnitsBefore: 2300, returns: { units: 40 } })]))), [-8000]);
  const first = run(gbp([fmt({ returns: { units: 10 } })]));
  assertEquals([first.lines[0].rate, first.lines[0].amountMinor], ["10%", -2000]);

  const net = run(gbp([{ format: "audio", basis: "net", tiers: [{ fromUnit: 1, percent: 25 }], lifeToDateUnitsBefore: 100, returns: { units: 4, netReceiptsMinor: 1999 } }]));
  assertEquals([net.lines[0].basis, net.lines[0].amountMinor], ["£19.99 net receipts", -500]);
});

test("returns that exceed earnings make a debit", async () => {
  const royalty = gbp([fmt({ returns: { units: 100 } })]);
  const r = run(royalty);
  assertEquals([r.totalMinor, r.negative, r.totalAbs, r.total], [-20000, true, "£200.00", "-£200.00"]);
  const text = await textsOf(royalty);
  assertStringIncludes(text, "shows a debit of £200.00.");
  assertEquals((await idsOf(gbp([fmt({ sales: [{ units: 1 }] })]))).includes("royalty-earnings-debit"), false);
});

test("sub-rights are shared after the withholding tax, shown once", async () => {
  const sr = (more: Obj) => run({ currency: "GBP", subRights: [{ right: "Spanish", licensee: "Planeta", grossMinor: 1000000, authorSharePercent: 50, ...more }] });
  const r = sr({ foreignTaxWithheldMinor: 150000 });
  assertEquals(r.lines.map((l) => [l.description, l.units, l.basis, l.rate, l.amount, l.memo]), [
    ["Spanish, Planeta", "", "£10,000.00 less £1,500.00 tax", "50%", "£4,250.00", false],
    ["Foreign tax withheld by Planeta (Spanish), already deducted", "", "£1,500.00", "", "", true],
  ]);
  assertEquals([r.hasMemo, r.totalMinor, r.total], [true, 425000, "£4,250.00"]);
  assertEquals((await idsOf({ currency: "GBP", subRights: [{ right: "S", licensee: "P", grossMinor: 1000000, foreignTaxWithheldMinor: 1, authorSharePercent: 50 }] })).includes("royalty-earnings-tax-note"), true);

  const plain = sr({ grossMinor: 333333 });
  assertEquals([plain.lines.length, plain.lines[0].amountMinor, plain.lines[0].basis, plain.hasMemo], [1, 166667, "£3,333.33", false]);
  assertEquals(sr({ grossMinor: 999, authorSharePercent: 80 }).lines[0].amountMinor, 799);
});

test("nothing to print says so, and life-to-date still shows when formats exist", async () => {
  for (const royalty of [{ currency: "GBP" }, { currency: "GBP", formats: [], subRights: [] }]) {
    const r = run(royalty);
    assertEquals([r.outcome, r.total, r.lifeToDate, r.negative, r.totalMinor], ["none", "", "", false, 0]);
    assertEquals((await idsOf(royalty)).filter((id) => id.startsWith("royalty-earnings")), ["royalty-earnings", "royalty-earnings-none"]);
  }
  const zero = gbp([fmt({ lifeToDateUnitsBefore: 700, sales: [{ units: 0 }] })]);
  assertEquals([run(zero).outcome, run(zero).lifeToDate], ["none", "Hardback 700"]);
  const ids = await idsOf(zero);
  assertEquals(ids.includes("royalty-earnings-none"), true);
  assertEquals(ids.includes("royalty-earnings-life-to-date"), true);
  assertStringIncludes(await textsOf(zero), "Copies sold to date, before returns: Hardback 700.");
});

test("one item gives one line, and the preview gives more than twelve", () => {
  assertEquals(run(gbp([fmt({ sales: [{ units: 10 }] })])).lines.length, 1);
  const preview = (COMPONENTS.find((c) => c.id === "royalty-earnings-by-format")!.previewData as { royalty: Obj }).royalty;
  const r = run(preview);
  assertEquals(r.lines.length > 12, true);
  assertEquals(r.totalMinor, r.lines.reduce((s, l) => s + l.amountMinor, 0));
  assertEquals(
    r.lifeToDate,
    "Hardback (slipcased collector's edition with illustrated endpapers) 52,050; " +
      "B-format paperback (export and open market edition) 338,680; Ebook 121,630; " +
      "Unabridged audiobook (download and streaming) 31,200",
  );
});

test("missing optional fields fall back", () => {
  const r = run(gbp([fmt({ sales: [{ units: 10 }] })]));
  assertEquals([r.lines[0].description, r.lines[0].amountMinor], ["Hardback: copies 1–2,500", 2000]);
  assertEquals(run(gbp([{ ...fmt({ sales: [{ units: 1 }] }), format: "paperback" }])).lines[0].description.startsWith("Paperback"), true);
  assertEquals(run(gbp([{ ...fmt({ sales: [{ units: 1 }] }), format: "ebook" }])).lines[0].description.startsWith("Ebook"), true);
  assertEquals(run(gbp([{ ...fmt({ sales: [{ units: 1 }] }), format: "audio" }])).lines[0].description.startsWith("Audiobook"), true);
  assertEquals(r.lines.length, 1); // no returns line, no memo row
  assertEquals(run(gbp([fmt({ label: "Home edition", sales: [{ units: 1 }] })])).lines[0].description.startsWith("Home edition"), true);
});

test("invalid input throws, naming the field", async () => {
  const hd = (h: Obj) => ({ highDiscount: h });
  const sold = (o: Obj) => gbp([fmt({ sales: [{ units: 1 }], ...o })]);
  const sub = (o: Obj) => ({ currency: "GBP", subRights: [{ right: "R", licensee: "L", grossMinor: 100, authorSharePercent: 50, ...o }] });
  const { lifeToDateUnitsBefore: _, ...noBefore } = fmt();
  const { basis: __, ...noBasis } = fmt();
  const { listPriceMinor: ___, ...noList } = fmt();
  const bad: Array<[Obj, string]> = [
    [sold({ tiers: [{ fromUnit: 1, percent: 10 }, { fromUnit: 1, percent: 12 }] }), "tiers[1].fromUnit"],
    [sold({ tiers: [{ fromUnit: 1, percent: 10 }, { fromUnit: 5001, percent: 15 }, { fromUnit: 2501, percent: 12.5 }] }), "tiers[2].fromUnit"],
    [sold({ tiers: [{ fromUnit: 2, percent: 10 }] }), "tiers[0].fromUnit"],
    [sold({ tiers: [] }), "tiers"],
    [gbp([noBasis]), "basis"],
    [sold({ basis: "gross" }), "basis"],
    [gbp([noList]), "listPriceMinor"],
    [gbp([noBefore]), "lifeToDateUnitsBefore"],
    [sold({ sales: [{ units: -1 }] }), "sales[0].units"],
    [sold({ sales: [{ units: 1.5 }] }), "sales[0].units"],
    [sold({ returns: { units: -1 } }), "returns.units"],
    [sold({ basis: "net", returns: { units: 1 } }), "returns.netReceiptsMinor"],
    [sold({ basis: "net", sales: [{ units: 1 }], listPriceMinor: undefined }), "netReceiptsMinor"],
    [sold({ tiers: [{ fromUnit: 1, percent: 12.345 }] }), "tiers[0].percent"],
    [sub({ foreignTaxWithheldMinor: 101 }), "foreignTaxWithheldMinor"],
    [{ ...sold({}), ...hd({ deeper: [{ thresholdPercent: 55, numerator: 1, denominator: 2 }] }) }, "deeper[0].thresholdPercent"],
    [{ ...sold({}), ...hd({ fraction: { numerator: 5, denominator: 4 } }) }, "numerator"],
    [sold({ format: "vinyl" }), "format"],
    [{ ...sold({}), currency: "XX" }, "royalty.currency"],
    [{ formats: [] }, "royalty.currency"],
    [sub({ right: " " }), "right"],
  ];
  for (const [royalty, field] of bad) {
    await assertRejects(() => Promise.resolve().then(() => run(royalty)), Error, field);
  }
  // An invalid format fails the document even when it sold nothing.
  await assertRejects(() => Promise.resolve().then(() => run(gbp([fmt({ tiers: [] })]))), Error, "tiers");
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Statement">
      <Section id="body" title="Body">
        <RoyaltyEarningsByFormat />
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

async function built(royalty: Obj) {
  return await buildDocument(templateOf(), { royalty }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);
}
async function idsOf(royalty: Obj): Promise<string[]> {
  const body = (await built(royalty)).nodes[0] as { children: DocumentNode[] };
  return collect(body.children).map((n) => n.id);
}
async function textsOf(royalty: Obj): Promise<string> {
  return JSON.stringify((await built(royalty)).nodes[0]);
}

test("a published document resolves its rows to what a local build prints", async () => {
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const section = collect(published.nodes).find((n) => n.id === "royalty-earnings") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).name, "royaltyEarningsByFormat");
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).inputs, [{ type: "ref", ref: { scope: "data", path: "royalty" } }]);

  const preview = (COMPONENTS.find((c) => c.id === "royalty-earnings-by-format")!.previewData as { royalty: Obj }).royalty;
  const cases: Obj[] = [
    { ...gbp([fmt({ lifeToDateUnitsBefore: 2300, sales: [{ units: 500, discountPercent: 40 }] })]), subRights: [{ right: "Serial", licensee: "Paper", grossMinor: 100000, foreignTaxWithheldMinor: 10000, authorSharePercent: 90 }] },
    gbp([fmt({ returns: { units: 100 } })]),
    { currency: "GBP" },
    preview,
  ];
  // Each row's cell texts. A resolved empty cell keeps an empty paragraph where a local build keeps none, so compare text, not shape.
  const textIn = (n: DocumentNode): string =>
    (n as { text?: string }).text ?? ((n as { children?: DocumentNode[] }).children ?? []).map(textIn).join("");
  const cells = (nodes: DocumentNode[]) =>
    collect(nodes).filter((n) => n.kind === "tableRow").map((row) => ((row as { children: DocumentNode[] }).children).map(textIn));
  for (const royalty of cases) {
    const direct = await built(royalty);
    const resolved = await resolveDocument(published, state({ royalty }), { derivers });
    assertEquals(cells(resolved.nodes), cells(direct.nodes));
  }
  // The first case has a sale split in two plus a sub-right and its memo row: header, 4 lines, total.
  const first = await resolveDocument(published, state({ royalty: cases[0] }), { derivers });
  assertEquals(cells(first.nodes).length, 6);
  assertEquals(cells(first.nodes)[1], ["Hardback: copies 1–2,500", "200", "£20.00 list", "10%", "£400.00"]);
});
