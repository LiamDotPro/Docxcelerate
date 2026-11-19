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
import { COMPONENTS } from "docxcelerate/registry";
import { assertEquals } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/incoterm-charges-and-customs-value/incoterm-charges-and-customs-value.node.tsx";
const { IncotermChargesAndCustomsValue } = await loadRegistryModule<
  typeof import("../registry/components/incoterm-charges-and-customs-value/incoterm-charges-and-customs-value.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/incoterm-charges-and-customs-value.deriver.ts")
>("components/derivers/incoterm-charges-and-customs-value.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ incotermChargesAndCustomsValue: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Out {
  term: string;
  goods: string;
  legs: Record<string, { label: string; treatment: string; amount: string }>;
  total: string;
  customsLine: string;
  customsNote: string;
  warnSeaOnly: string;
  warnNamedPlace: string;
  warnDdpVat: string;
  warnImpliedInsurance: string;
  totalMinor: number;
  fobMinor: number;
  cifMinor: number | null;
  customsMinor: number | null;
}

const assertThrows = (fn: () => unknown, text: string) => {
  try {
    fn();
  } catch (e) {
    assertEquals((e as Error).message.includes(text), true);
    return;
  }
  throw new Error("expected a throw containing " + text);
};
const run = (s: Obj): Out => mod.run([s], {} as RuntimeState) as Out;
const cif = { valuationBasis: "cif-border" };
const fob = { valuationBasis: "fob-export" };
const ship = (more: Obj): Obj => ({
  incoterm: "FOB", namedPlace: "Port of Felixstowe", currency: "EUR", goodsMinor: 100000, destination: cif, ...more,
});
const leg = (kind: string, amountMinor: number, inPrice?: boolean) => (inPrice ? { kind, amountMinor, inPrice } : { kind, amountMinor });
const warnings = (o: Out) => [o.warnSeaOnly, o.warnNamedPlace, o.warnDdpVat, o.warnImpliedInsurance].filter(Boolean);

const entry = COMPONENTS.find((c) => c.id === "incoterm-charges-and-customs-value")!;
const preview = (entry.previewData as { shipment: Obj }).shipment;

const exw = ship({
  incoterm: "EXW", namedPlace: "Lyon", mode: "road",
  legs: [leg("origin-inland", 5000), leg("export-clearance", 1000), leg("origin-terminal", 2000), leg("main-carriage", 20000), leg("insurance", 500)],
});
const cifUs = ship({
  incoterm: "CIF", namedPlace: "New York", currency: "USD", mode: "sea", destination: fob,
  legs: [leg("main-carriage", 15000, true), leg("insurance", 400, true)],
});
const ddp = ship({
  incoterm: "DDP", goodsMinor: 200000, mode: "air", destination: { ...cif, airFreightDutiablePercent: 70 },
  legs: [
    leg("main-carriage", 10000, true), leg("insurance", 300), leg("destination-terminal", 800, true),
    leg("onward-delivery", 1500), leg("import-duty", 6000, true), leg("import-tax", 40000),
  ],
});

test("1 preview data", () => {
  const o = run(preview);
  assertEquals([o.totalMinor, o.fobMinor, o.cifMinor, o.customsMinor], [131264700, 128965300, 130993500, 130993500]);
  assertEquals(o.total, "€1,312,647.00");
  assertEquals(warnings(o), []);
  assertEquals(o.customsNote, "");
  assertEquals(o.legs["export-packing"].treatment, "Included in price");
  assertEquals(o.legs["import-duty"].treatment, "For buyer's account");
  assertEquals(o.legs["import-tax"].treatment, "For buyer's account");
});

test("2 EXW customs value is above the invoice", () => {
  const o = run(exw);
  assertEquals([o.totalMinor, o.fobMinor, o.cifMinor, o.customsMinor], [100000, 108000, 128500, 128500]);
  for (const k of ["origin-inland", "export-clearance", "origin-terminal", "main-carriage", "insurance"]) {
    assertEquals(o.legs[k].treatment, "For buyer's account");
  }
});

test("3 CIF to the US is below the invoice", () => {
  const o = run(cifUs);
  assertEquals([o.totalMinor, o.fobMinor, o.customsMinor], [100000, 84600, 84600]);
  assertEquals(o.customsNote.includes("19 CFR 141.86(a)(8)"), true);
  assertEquals(o.customsLine, "Customs value, FOB at the place of export: US$846.00.");
  assertEquals(warnings(o), []);
});

test("4 a sea-only rule for container freight", () => {
  const o = run({ ...cifUs, mode: "sea-container" });
  assertEquals(o.warnSeaOnly.includes("CIP") && o.warnSeaOnly.includes("containerised"), true);
  assertEquals([o.totalMinor, o.customsMinor], [100000, 84600]);
  assertEquals(run({ ...cifUs, incoterm: "FAS", mode: "road", legs: [] }).warnSeaOnly.includes("FCA"), true);
  assertEquals(run({ ...cifUs, incoterm: "CFR", mode: "air", legs: [leg("main-carriage", 1, true)] }).warnSeaOnly.includes("CPT"), true);
});

test("5 DDP by air with a percentage", () => {
  const o = run(ddp);
  assertEquals([o.totalMinor, o.fobMinor, o.cifMinor, o.customsMinor], [241800, 183200, 190500, 190500]);
  assertEquals(o.warnDdpVat !== "", true);
  assertEquals(o.customsNote.includes("counted at 70%"), true);
});

test("6 air freight rounds half up, once", () => {
  const air = (main: number, pct: number) =>
    run(ship({ incoterm: "FCA", mode: "air", goodsMinor: 50000, legs: [leg("main-carriage", main)], destination: { ...cif, airFreightDutiablePercent: pct } }));
  const o = air(12345, 65);
  assertEquals([o.fobMinor, o.cifMinor], [50000, 58024]);
  assertEquals(air(1, 50).cifMinor, 50001);
  assertEquals(air(3, 50).cifMinor, 50002);
});

test("7 air without a percentage counts in full", () => {
  const o = run(ship({ incoterm: "FCA", mode: "air", legs: [leg("main-carriage", 12345)] }));
  assertEquals(o.cifMinor, 112345);
  assertEquals(o.customsNote.startsWith("Air freight counted in full."), true);
});

test("8 the percentage is ignored for a FOB-export destination", () => {
  const o = run(ship({ incoterm: "FCA", mode: "air", legs: [leg("main-carriage", 12345)], destination: { ...fob, airFreightDutiablePercent: 50 } }));
  assertEquals(o.customsNote.includes("19 CFR"), true);
  assertEquals(o.customsMinor, 100000);
});

test("9 CFR with insurance", () => {
  const o = run(ship({ incoterm: "CFR", mode: "sea", legs: [leg("main-carriage", 5000, true), leg("insurance", 300)] }));
  assertEquals(o.warnImpliedInsurance.includes("CIF"), true);
  assertEquals(o.legs["insurance"].treatment, "For buyer's account");
  assertEquals(o.totalMinor, 100000);
});

test("10 a missing or blank named place", () => {
  for (const place of [undefined, "   "]) {
    const o = run(ship({ namedPlace: place }));
    assertEquals(o.term, "FOB Incoterms® 2020");
    assertEquals(o.warnNamedPlace.includes("FOB"), true);
  }
  assertEquals(run(ship({ namedPlace: "  Port of Felixstowe " })).term, "FOB Port of Felixstowe Incoterms® 2020");
});

test("11 no mode, no mode warning", () => {
  assertEquals(run(ship({ legs: [leg("main-carriage", 100)] })).warnSeaOnly, "");
});

test("12 no legs, FOB to the US", () => {
  for (const legs of [undefined, []]) {
    const o = run(ship({ legs, destination: fob }));
    assertEquals(o.legs["main-carriage"].label, "");
    assertEquals([o.totalMinor, o.customsMinor], [100000, 100000]);
  }
});

test("13 no legs, CIF-border cannot be stated", () => {
  const o = run(ship({ legs: undefined }));
  assertEquals([o.cifMinor, o.customsMinor], [null, null]);
  assertEquals(o.customsLine, "Customs value: not stated, because the cost of main carriage to the place of entry was not given.");
});

test("14 CIF to the US needs freight and insurance itemised", () => {
  const base = { incoterm: "CIF", destination: fob };
  const a = run(ship({ ...base, legs: [leg("insurance", 400, true)] }));
  assertEquals(a.customsMinor, null);
  assertEquals(a.customsLine.includes("main carriage"), true);
  const b = run(ship({ ...base, legs: [leg("main-carriage", 15000, true)] }));
  assertEquals(b.customsMinor, null);
  assertEquals(b.cifMinor, 100000);
  assertEquals(b.customsLine.includes("insurance"), true);
});

test("15 zero", () => {
  const o = run(ship({ goodsMinor: 0, destination: fob, legs: [leg("export-packing", 0), leg("main-carriage", 0)] }));
  assertEquals([o.goods, o.total, o.customsMinor, o.legs["export-packing"].amount], ["€0.00", "€0.00", 0, "€0.00"]);
});

test("16 one leg", () => {
  const o = run(ship({ incoterm: "CPT", legs: [leg("main-carriage", 900)] }));
  assertEquals(Object.values(o.legs).filter((l) => l.label).length, 1);
  assertEquals(o.legs["main-carriage"].treatment, "Charged");
  assertEquals(o.totalMinor, 100900);
});

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Invoice">
      <Section id="body" title="Body">
        <IncotermChargesAndCustomsValue />
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

const built = (s: Obj) =>
  buildDocument(templateOf(), { shipment: s }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);
const rowIds = async (s: Obj) =>
  collect((await built(s)).nodes).map((n) => n.id).filter((id) => id.startsWith("incoterm-charges-") && !id.includes("cell"));

test("17 rows print in journey order whatever the input order", async () => {
  const legs = (preview.legs as Array<{ kind: string }>).slice().reverse();
  const ids = await rowIds({ ...preview, legs });
  const kinds = ["export-packing", "origin-inland", "export-clearance", "origin-terminal", "main-carriage", "insurance",
    "destination-terminal", "onward-delivery", "import-duty", "import-tax"];
  assertEquals(ids.filter((i) => kinds.some((k) => i === "incoterm-charges-" + k)), kinds.map((k) => "incoterm-charges-" + k));
  assertEquals(ids.indexOf("incoterm-charges-goods") < ids.indexOf("incoterm-charges-export-packing"), true);
  assertEquals(ids.indexOf("incoterm-charges-import-tax") < ids.indexOf("incoterm-charges-total"), true);
});

test("18 a currency with no minor unit", () => {
  const o = run(ship({ currency: "JPY", goodsMinor: 1500000 }));
  assertEquals(o.goods, "JP¥1,500,000") // en-GB names the yen, and no decimals follow;
});

test("19 invalid input throws", () => {
  const bad = (more: Obj, text: string) => assertThrows(() => run(ship(more)), text);
  bad({ incoterm: "DAT" }, "incoterm");
  bad({ currency: "XYZ" }, "currency");
  bad({ currency: undefined }, "currency");
  for (const g of [-1, 10.5, "100"]) bad({ goodsMinor: g }, "goodsMinor");
  bad({ legs: "x" }, "legs");
  for (const a of [-1, 1.5]) bad({ legs: [leg("main-carriage", a)] }, "legs[0].amountMinor");
  bad({ legs: [leg("shipping", 1)] }, "kind");
  bad({ legs: [leg("main-carriage", 1), leg("main-carriage", 2)] }, "twice");
  bad({ legs: [leg("main-carriage", 1, true)] }, "inPrice");
  bad({ incoterm: "CIF", legs: [leg("main-carriage", 100001, true)] }, "more than shipment.goodsMinor");
  bad({ mode: "ship" }, "mode");
  bad({ destination: undefined }, "valuationBasis");
  bad({ destination: { valuationBasis: "x" } }, "valuationBasis");
  for (const p of [101, 65.5, -1]) bad({ destination: { ...cif, airFreightDutiablePercent: p } }, "airFreightDutiablePercent");
});

test("20 a publish build carries 15 arms and one deriver call", async () => {
  assertEquals(branches, 15);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 15);
  const section = all.find((n) => n.id === "incoterm-charges") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).name, "incotermChargesAndCustomsValue");

  const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);
  for (const s of [exw, ddp, preview, { ...cifUs, mode: "air" }]) {
    const state: RuntimeState = {
      ctx: { shipment: s }, derived: {}, dataProvider: new InMemoryDataProvider({ shipment: s }), aiClient: { generateParagraph: () => "" },
    };
    const resolved = await resolveDocument(published, state, { derivers });
    assertEquals(seen(resolved.nodes), seen((await built(s)).nodes));
  }
});
