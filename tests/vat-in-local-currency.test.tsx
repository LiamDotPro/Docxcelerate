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
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/vat-in-local-currency.deriver.ts")
>("components/derivers/vat-in-local-currency.deriver.ts");
const { VatInLocalCurrency } = await loadRegistryModule<
  typeof import("../registry/components/vat-in-local-currency/vat-in-local-currency.node.tsx")
>("components/vat-in-local-currency/vat-in-local-currency.node.tsx");
const { COMPONENTS } = await import("../src/registry/catalog.ts");

type Obj = Record<string, any>;
const run = (c: Obj): Obj => mod.run([c], {} as RuntimeState) as Obj;
const bad = async (c: Obj, needle: string) =>
  await assertRejects(() => Promise.resolve().then(() => run(c)), Error, "vatInLocalCurrency: " + needle);

/** The worked UK case: EUR to GBP at 0.8571, one 20% band. */
const uk = (over: Obj = {}, rate: Obj = {}): Obj => ({
  jurisdiction: "UK",
  invoiceCurrency: "EUR",
  vatCurrency: "GBP",
  taxPointDate: "2026-03-02",
  bands: [{ ratePercent: "20", net: "1234.56", vat: "246.91" }],
  rate: { value: "0.8571", base: "EUR", quote: "GBP", source: "hmrc-period", rateDate: "2026-03-01", ...rate },
  ...over,
});
const pl = (rate: Obj = {}, over: Obj = {}): Obj =>
  uk({ jurisdiction: "PL", vatCurrency: "PLN", issueDate: "2026-03-31", taxPointDate: "2026-03-31", ...over }, {
    value: "4.2725", quote: "PLN", source: "nbp", rateDate: "2026-03-30", ...rate,
  });
const warns = (c: Obj) => run(c).warnings.map((w: Obj) => w.text);

test("UK band: each figure is rounded once, no warnings, rate line says direction and source", () => {
  const r = run(uk());
  assertEquals(r.converted, true);
  assertEquals(r.bands, [{ rate: "20%", net: "€1,234.56", vat: "€246.91", localNet: "£1,058.14", localVat: "£211.63" }]);
  assertEquals(r.total, { net: "€1,234.56", vat: "€246.91", localNet: "£1,058.14", localVat: "£211.63" });
  assertEquals(r.warnings, []);
  assertStringIncludes(r.statement, "1 EUR = 0.8571 GBP, HMRC period rate for March 2026");
  assertEquals(r.hasAdvances, false);
});

test("an inverse quote divides", () => {
  const r = run(uk({}, { value: "1.1667", base: "GBP", quote: "EUR" }));
  assertEquals([r.bands[0].localNet, r.bands[0].localVat], ["£1,058.16", "£211.63"]);
});

test("bands convert separately and the local total is their sum", () => {
  const r = run(uk({ bands: [
    { ratePercent: "20", net: "1234.56", vat: "246.91" },
    { ratePercent: "5", net: "60.60", vat: "3.03" },
  ] }));
  assertEquals(r.bands.map((b: Obj) => [b.localNet, b.localVat]), [["£1,058.14", "£211.63"], ["£51.94", "£2.60"]]);
  assertEquals(r.total.localVat, "£214.23");
  assertEquals(r.total.vat, "€249.94");
});

test("credit notes mirror, halves go away from zero, zero bands print", () => {
  const credit = run(uk({ bands: [{ ratePercent: "20", net: "-1234.56", vat: "-246.91" }] }));
  assertEquals([credit.bands[0].localNet, credit.bands[0].localVat], ["-£1,058.14", "-£211.63"]);
  const half = (vat: string) => run(uk({ bands: [{ ratePercent: "20", net: "0", vat }] }, { value: "0.5" })).bands[0].localVat;
  assertEquals([half("0.70"), half("-0.70")], ["£0.35", "-£0.35"]);
  const zero = run(uk({ bands: [{ ratePercent: "0", net: "12500.00", vat: "0.00" }] }));
  assertEquals([zero.bands[0].rate, zero.bands[0].localVat], ["0%", "£0.00"]);
});

test("zero-decimal invoice currency", async () => {
  const r = run(uk({ invoiceCurrency: "JPY", bands: [{ ratePercent: "20", net: "123457", vat: "0" }] }, { base: "JPY", value: "0.0052" }));
  assertEquals(r.bands[0].localNet, "£641.98");
  await bad(uk({ invoiceCurrency: "JPY", bands: [{ ratePercent: "20", net: "123457.5", vat: "0" }] }, { base: "JPY" }), "vatConversion.bands[0].net");
});

test("Bulgaria's fixed euro rate", async () => {
  const bg = (rate: Obj = {}, over: Obj = {}) => ({
    jurisdiction: "BG", invoiceCurrency: "BGN", vatCurrency: "EUR", taxPointDate: "2026-02-10",
    bands: [{ ratePercent: "20", net: "1000.00", vat: "200.00" }],
    rate: { value: "1.95583", base: "EUR", quote: "BGN", source: "fixed-euro", ...rate },
    ...over,
  });
  const r = run(bg());
  assertEquals([r.bands[0].localNet, r.bands[0].localVat, r.warnings], ["€511.29", "€102.26", []]);
  await bad(bg({ value: "0.51129", base: "BGN", quote: "EUR" }), "vatConversion.rate");
  await bad(bg({ value: "1.9558" }), "vatConversion.rate");
  await bad(bg({ source: "ecb", rateDate: "2026-02-10" }), "vatConversion.rate.source");
  await bad(uk({}, { source: "fixed-euro" }), "vatConversion.rate.source");
  await bad(bg({}, { taxPointDate: "2025-12-15" }), "vatConversion.vatCurrency");
});

test("same currency converts nothing and ignores the rate", () => {
  const r = run(uk({ vatCurrency: "EUR", jurisdiction: "EU", rate: undefined, bands: undefined }));
  assertEquals([r.converted, r.sameCurrency, r.bands, r.advances, r.warnings], [false, "EUR", [], [], []]);
});

test("bad rates, pairs and currencies throw", async () => {
  await assertRejects(
    () => Promise.resolve().then(() => run(uk({}, { base: "USD" }))),
    Error,
    "EUR and GBP",
  );
  for (const value of ["0", "-1", "1e-7", "0.12345678901"]) await bad(uk({}, { value }), "vatConversion.rate.value");
  await bad(uk({}, { source: "guess" }), "vatConversion.rate.source");
  await bad(uk({}, { rateDate: undefined }), "vatConversion.rate.rateDate");
  await bad(uk({ jurisdiction: "FR" }), "vatConversion.jurisdiction");
  await bad(uk({ invoiceCurrency: "XXX1" }), "vatConversion.invoiceCurrency");
  await bad(uk({ vatCurrency: "EUR" }, {}), "vatConversion.vatCurrency");
  await bad(uk({ taxPointDate: "soon" }), "vatConversion.taxPointDate");
  await bad(pl({}, { vatCurrency: "EUR" }), "vatConversion.vatCurrency");
  await bad(pl({}, { issueDate: undefined }), "vatConversion.issueDate");
});

test("bands must exist, be distinct and be decimals", async () => {
  await bad(uk({ bands: undefined }), "vatConversion.bands");
  await bad(uk({ bands: [] }), "vatConversion.bands");
  await bad(uk({ bands: [{ ratePercent: "20", net: "1", vat: "0.2" }, { ratePercent: "20.0", net: "1", vat: "0.2" }] }), "vatConversion.bands[1].ratePercent");
  await bad(uk({ bands: [{ ratePercent: "-5", net: "1", vat: "0" }] }), "vatConversion.bands[0].ratePercent");
  await bad(uk({ bands: [{ ratePercent: "20", net: 0.1 + 0.2, vat: "0" }] }), "vatConversion.bands[0].net");
});

test("five bands print in the given order and the totals are the band sums", () => {
  const rates = ["20", "5", "0", "7.7", "12.5"];
  const r = run(uk({ bands: rates.map((ratePercent, i) => ({ ratePercent, net: `${100 + i}.10`, vat: `${i}.05` })) }));
  assertEquals(r.bands.map((b: Obj) => b.rate), ["20%", "5%", "0%", "7.7%", "12.5%"]);
  assertEquals(r.total.net, "€510.50");
  assertEquals(r.total.vat, "€10.25");
});

test("UK source and date warnings", () => {
  assertEquals(warns(uk({}, { source: "market-selling", rateDate: "2026-03-05" })), [
    "The exchange rate is dated 5 March 2026, after the tax point of 2 March 2026.",
  ]);
  assertEquals(warns(uk({}, { rateDate: "2026-02-25" })), [
    "An HMRC period rate applies only to its own month: this one is for February 2026, the tax point is in March 2026.",
  ]);
  const ecb = run(uk({}, { source: "ecb", rateDate: "2026-03-01" }));
  assertEquals(ecb.warnings.length, 1);
  assertStringIncludes(ecb.warnings[0].text, "ECB reference rate of 1 March 2026 is not a rate the UK accepts for VAT; use ");
  assertEquals(ecb.bands[0].localNet, "£1,058.14");
});

test("Poland takes the last working day before the tax point", () => {
  const w = (rate: Obj, over: Obj = {}) => warns(pl(rate, over));
  assertEquals(w({}), []);
  const head = "In Poland the rate is the one from the last working day before the tax point, 30 March 2026; this one is dated ";
  assertEquals(w({ rateDate: "2026-03-31" }), [head + "31 March 2026, the day itself."]);
  assertEquals(w({ rateDate: "2026-04-01" }), [head + "1 April 2026, after the tax point."]);
  assertEquals(w({ rateDate: "2026-03-28" }), [head + "28 March 2026, which is not a working day."]);
  assertEquals(w({ rateDate: "2026-03-27" }), [head + "27 March 2026, before the last working day."]);
  // Easter Monday 2026 is 6 April, so the last working day before 7 April is Friday the 3rd.
  assertEquals(w({ rateDate: "2026-04-06" }, { taxPointDate: "2026-04-07", issueDate: "2026-04-07" }), [
    "In Poland the rate is the one from the last working day before the tax point, 3 April 2026; this one is dated 6 April 2026, which is not a working day.",
  ]);
  // Issued before the tax point: the date of issue is the reference.
  assertEquals(w({ rateDate: "2026-03-19" }, { issueDate: "2026-03-20" }), []);
  assertEquals(w({ rateDate: "2026-03-20" }, { issueDate: "2026-03-20" }).length, 1);
  // Corpus Christi 2026 is 4 June, and Christmas Eve is a holiday from 2025.
  const on = (rateDate: string, day: string) => w({ rateDate }, { taxPointDate: day, issueDate: day }).length;
  assertEquals([on("2026-06-03", "2026-06-05"), on("2026-06-04", "2026-06-05")], [0, 1]);
  assertEquals([on("2026-12-23", "2026-12-28"), on("2026-12-24", "2026-12-28")], [0, 1]);
});

const previewData = COMPONENTS.find((c) => c.id === "vat-in-local-currency")!.previewData as Obj;
const advanced = previewData.vatConversion as Obj;

test("advances keep their own rate and the bands are the remainder", () => {
  const r = run(advanced);
  assertEquals(r.hasAdvances, true);
  assertEquals(r.warnings, []);
  assertEquals(r.bands.map((b: Obj) => [b.rate, b.net, b.localNet, b.localVat]), [
    ["23%", "€884,750.00", "PLN 3,780,094.38", "PLN 869,421.71"].map((x) => x),
    ["8%", "€48,210.40", "PLN 205,978.93", "PLN 16,478.31"],
    ["0%", "€12,500.00", "PLN 53,406.25", "PLN 0.00"],
  ].map((b) => b));
  assertEquals(r.advances, [{
    label: "Advance invoice FV/ZAL/2026/02/0147, chargeable 16 February 2026",
    rateLine: "1 EUR = 4.2187 PLN, NBP average rate of 13 February 2026",
    net: "€400,000.00",
    vat: "€92,000.00",
    localNet: "PLN 1,687,480.00",
    localVat: "PLN 388,120.40",
  }]);
  assertStringIncludes(r.statement, "each of which keeps the rate it was declared at");
});

test("advance problems throw or warn under the advance's own name", async () => {
  const adv = (over: Obj = {}, rate: Obj = {}) => ({
    ...advanced,
    advances: [{ ...advanced.advances[0], ...over, rate: { ...advanced.advances[0].rate, ...rate } }],
  });
  const line = (ratePercent: string, net: string, vat: string) => [{ ratePercent, net, vat }];
  await bad(adv({ bands: line("5", "1.00", "0.05") }), "vatConversion.advances[0].bands");
  await bad(adv({ bands: line("23", "1284750.01", "0") }), "the advances exceed the 23% band");
  await bad(adv({ chargeableDate: "2026-04-01" }), "vatConversion.advances[0].chargeableDate");
  await bad(adv({ invoiceNumber: " " }), "vatConversion.advances[0].invoiceNumber");
  await bad({ ...advanced, bands: [{ ratePercent: "23", net: "-10.00", vat: "-2.30" }] }, "vatConversion.bands[0]");
  const w = warns(adv({}, { rateDate: "2026-02-16" }));
  assertEquals(w.length, 1);
  assertStringIncludes(w[0], "Advance invoice FV/ZAL/2026/02/0147: In Poland");
  // Not Poland: a rate dated after the advance's chargeable event.
  const ukAdv = uk({ taxPointDate: "2026-03-20", advances: [{
    invoiceNumber: "A1", chargeableDate: "2026-03-02", bands: line("20", "100.00", "20.00"),
    rate: { value: "0.85", base: "EUR", quote: "GBP", source: "market-selling", rateDate: "2026-03-05" },
  }] });
  assertEquals(warns(ukAdv), ["Advance invoice A1: The exchange rate is dated 5 March 2026, after the tax point of 2 March 2026."]);
});

test("no advances, or an empty list, is the same output", () => {
  assertEquals(run(uk({ advances: [] })), run(uk()));
  assertEquals(run(uk({ advances: undefined })), run(uk()));
});

// ---------------------------------------------------------------------------
// The node
// ---------------------------------------------------------------------------

const ARRANGEMENTS = ["columns", "block"] as const;
type Arrangement = (typeof ARRANGEMENTS)[number];

function documentOf(arrangement?: Arrangement) {
  return template(
    <Document title="VAT">
      <VatInLocalCurrency arrangement={arrangement} />
    </Document>,
  );
}
function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((n) => {
    const own = n as { id?: string; children?: DocumentNode[]; rows?: DocumentNode[] };
    return [...(own.id ? [own.id] : []), ...ids(own.children ?? []), ...ids(own.rows ?? [])];
  });
}
function allNodes(nodes: DocumentNode[]): DocumentNode[] {
  return nodes.flatMap((n) => [n, ...allNodes((n as { children?: DocumentNode[] }).children ?? [])]);
}
function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((n) => {
    const own = n as { text?: string; children?: DocumentNode[]; rows?: DocumentNode[] };
    return [...(own.text ? [own.text] : []), ...texts(own.children ?? []), ...texts(own.rows ?? [])];
  });
}
const built = async (a: Arrangement | undefined, data: Obj) =>
  (await buildDocument(documentOf(a), { vatConversion: data } as never, { dynamicMode: "placeholder" })).nodes;

test("arrangements build and differ", async () => {
  const columns = ids(await built("columns", advanced));
  const block = ids(await built("block", advanced));
  for (const got of [columns, block]) {
    assertEquals(got.filter((i) => i.startsWith("vat-local-band-")).length, 3);
    assertEquals(got.filter((i) => i.startsWith("vat-local-advance-")).length, 1);
    for (const id of ["vat-local-statement", "vat-local-total", "vat-local-advances-intro"]) assertEquals(got.includes(id), true);
  }
  assertEquals([columns.includes("vat-local-bands"), columns.includes("vat-local-advances")], [true, true]);
  assertEquals([block.includes("vat-local-bands"), block.includes("vat-local-advances")], [false, false]);
  assertEquals(ids(await built(undefined, advanced)), columns);
});

test("warnings print after the arrangement, and same currency prints only its sentence", async () => {
  const warned = texts(await built("columns", uk({}, { source: "market-selling", rateDate: "2026-03-05" })));
  assertEquals(warned.filter((t) => t.startsWith("Check before issuing: The exchange rate is dated 5 March 2026")).length, 1);
  const same = await built("block", uk({ vatCurrency: "EUR", jurisdiction: "EU" }));
  assertEquals(ids(same), ["vat-in-local-currency-same-currency", "vat-local-same-currency"]);
});

test("every arrangement prints the same figures, placed differently", async () => {
  const r = run(advanced);
  const figures = [
    ...r.bands.flatMap((b: Obj) => [b.rate, b.net, b.vat, b.localNet, b.localVat]),
    ...Object.values(r.total),
    ...r.advances.flatMap((a: Obj) => [a.label, a.rateLine, a.net, a.vat, a.localNet, a.localVat]),
  ] as string[];
  for (const a of ARRANGEMENTS) {
    const got = (await built(a, advanced)).length && texts(await built(a, advanced)).join("\n");
    for (const f of figures) assertEquals((got as string).includes(f), true);
  }
  assertEquals(texts(await built("block", advanced)).includes(r.statement), true);
  assertEquals(texts(await built("columns", advanced)).includes(r.statement), true);
});

const bundled = createDeriverBundle({ vatInLocalCurrency: mod.run });
for (const arrangement of ARRANGEMENTS) {
  for (const data of [advanced, uk({ vatCurrency: "EUR", jurisdiction: "EU" })]) {
    const converted = data === advanced;
    test(`${arrangement}: published, bundled and resolved (${converted ? "converted" : "same currency"})`, async () => {
      const published = await buildDocument(documentOf(arrangement), createPublishData() as never, {
        branchMode: "publish",
        deriverMode: "preserve",
        dynamicMode: "resolve",
        aiClient: { generateParagraph: () => "" },
      });
      assertStringIncludes(JSON.stringify(published), "vatInLocalCurrency");
      const invocations = JSON.stringify(allNodes(published.nodes).flatMap((n) => (n as { derivers?: unknown[] }).derivers ?? []));
      assertStringIncludes(invocations, '"inputs":[{"type":"ref","ref":{"scope":"data","path":"vatConversion"}}]');

      const state: RuntimeState = {
        ctx: { vatConversion: data },
        derived: {},
        dataProvider: new InMemoryDataProvider({ vatConversion: data }),
        aiClient: { generateParagraph: () => "" },
      };
      const resolved = await resolveDocument(published, state, {
        derivers: await createDeriverRegistryFromBundle(bundled),
      });
      const got = ids(resolved.nodes);
      assertEquals(got.includes("vat-local-same-currency"), !converted);
      assertEquals(got.includes("vat-local-statement"), converted);
      assertEquals(got.includes("vat-local-advances-intro"), converted);
      // Rows of a loop over a deriver's list resolve only with the resolver fix (#83), so the loops
      // are checked as published and the totals, which are plain fields, as resolved.
      assertStringIncludes(JSON.stringify(published), '"source":{"scope":"derived","path":"vatInLocalCurrency.bands"}');
      assertStringIncludes(JSON.stringify(published), '"path":"vatInLocalCurrency.warnings"');
      if (converted) assertStringIncludes(JSON.stringify(resolved.nodes), "4,039,479.56");
    });
  }
}
