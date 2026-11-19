import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  dataRef,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, ParagraphNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/vat-place-of-supply.deriver.ts")
>("components/derivers/vat-place-of-supply.deriver.ts");
const { VatPlaceOfSupply } = await loadRegistryModule<
  typeof import("../registry/components/vat-place-of-supply/vat-place-of-supply.node.tsx")
>("components/vat-place-of-supply/vat-place-of-supply.node.tsx");
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ vatPlaceOfSupply: mod.run }));

type Obj = Record<string, unknown>;
const R: Obj = { NL: 21, DE: 19, FI: 25.5, FR: 20, GB: 20, IE: 23, ES: 21, GR: 24 };
const NOTE = "VAT has been charged because the customer's VAT number is missing or has not been validated.";
const UNDER = { currentYear: 0, priorYear: 0 };

const seller = (country: string, extra: Obj = {}): Obj => ({ country, euDistanceSales: UNDER, ...extra });
const biz = (country: string, extra: Obj = {}): Obj => ({
  country, isBusiness: true, vatNumber: "X1", vatNumberValidated: true, ...extra,
});
const consumer = (country: string, extra: Obj = {}): Obj => ({ country, isBusiness: false, ...extra });
const sup = (kind: string, extra: Obj = {}): Obj => ({ kind, ...extra });

interface Out { treatment: string; rate: number; line: string; note: string | null }
const run = (s: Obj, c: Obj, p: Obj, rates: unknown = R) => mod.run([s, c, p, rates], {} as RuntimeState) as Out;
const want = (treatment: string, rate: number, line: string, note: string | null = null) => ({
  treatment, rate, line, note,
});

test("the function survives being shipped as source text", async () => {
  const data = { seller: seller("NL"), customer: biz("DE"), supply: sup("goods"), vatRates: R };
  const via = await registry.run(
    {
      name: "vatPlaceOfSupply",
      output: "o",
      inputs: [dataRef("seller"), dataRef("customer"), dataRef("supply"), dataRef("vatRates")],
    } as never,
    { ctx: {}, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } },
  );
  assertEquals(via, run(seller("NL"), biz("DE"), sup("goods")));
});

test("domestic is standard whoever the customer is", () => {
  assertEquals(run(seller("NL"), consumer("NL"), sup("goods")), want("standard", 21, "VAT charged at 21% (Netherlands)."));
  assertEquals(run(seller("NL"), biz("NL"), sup("services")), want("standard", 21, "VAT charged at 21% (Netherlands)."));
  assertEquals(run(seller("GB"), consumer("XI"), sup("goods")), want("standard", 20, "VAT charged at 20% (United Kingdom)."));
  assertEquals(run(seller("XI"), consumer("XI"), sup("goods")), want("standard", 20, "VAT charged at 20% (Northern Ireland)."));
});

test("cross-border B2B inside the EU", () => {
  const rc = want("reverse-charge-services", 0, "No VAT charged: reverse charge, VAT to be accounted for by the customer (Germany).");
  assertEquals(run(seller("NL"), biz("DE"), sup("services")), rc);
  assertEquals(
    run(seller("NL"), biz("DE"), sup("goods")),
    want("intra-community-supply", 0, "No VAT charged: intra-Community supply of goods (Netherlands to Germany)."),
  );
  assertEquals(run(seller("NL"), biz("DE"), sup("services"), {}), rc);
});

test("an unproven business claim is charged VAT with a note", () => {
  const out = want("standard", 21, "VAT charged at 21% (Netherlands).", NOTE);
  for (const extra of [
    { vatNumberValidated: false },
    { vatNumberValidated: undefined },
    { vatNumber: "" },
    { vatNumber: "  " },
    { vatNumber: undefined },
  ]) {
    assertEquals(run(seller("NL"), biz("DE", extra), sup("services")), out);
  }
  assertEquals(run(seller("NL"), biz("DE", { vatNumberValidated: false }), sup("goods")), out);
});

test("a consumer never gets reverse charge", () => {
  const out = want("standard", 21, "VAT charged at 21% (Netherlands).");
  assertEquals(run(seller("NL"), consumer("DE", { vatNumber: "X1", vatNumberValidated: true }), sup("services")), out);
  assertEquals(run(seller("NL"), { country: "DE" }, sup("services")), out);
});

test("distance-selling threshold", () => {
  const fi = (cur: number, prior: number, extra: Obj = {}) =>
    run(seller("NL", { euDistanceSales: { currentYear: cur, priorYear: prior }, ...extra }), consumer("FI"), sup("digital"));
  const origin = want("standard", 21, "VAT charged at 21% (Netherlands).");
  const dest = want("standard", 25.5, "VAT charged at 25.5% (Finland) under the Union OSS scheme.");
  assertEquals(fi(10000, 10000), origin);
  assertEquals(fi(10000.01, 0), dest);
  assertEquals(fi(10000.004, 0), origin);
  assertEquals(fi(0, 12000), dest);
  assertEquals(fi(15000, 0), dest);
  assertEquals(fi(0, 0), origin);
  assertEquals(fi(0, 0, { ossRegistered: true }), dest);
  assertEquals(run(seller("NL", { ossRegistered: true, euDistanceSales: undefined }), consumer("FI"), sup("digital")), dest);
});

test("missing or bad sales figures throw", async () => {
  const go = (euDistanceSales: unknown) => () =>
    run(seller("NL", { euDistanceSales }), consumer("FI"), sup("digital"));
  for (const bad of [undefined, { currentYear: -1, priorYear: 0 }, { currentYear: NaN, priorYear: 0 }, { currentYear: "5000", priorYear: 0 }, { currentYear: 0 }]) {
    await assertRejects(() => Promise.resolve().then(() => go(bad)()), Error, "seller.euDistanceSales");
  }
});

test("goods shipped from another country skip the threshold; services ignore it", () => {
  assertEquals(
    run(seller("DE"), consumer("FR"), sup("goods", { shipsFrom: "NL" })),
    want("standard", 20, "VAT charged at 20% (France) under the Union OSS scheme."),
  );
  assertEquals(
    run(seller("NL", { euDistanceSales: { currentYear: 99999, priorYear: 99999 } }), consumer("FI"), sup("services")),
    want("standard", 21, "VAT charged at 21% (Netherlands)."),
  );
});

test("across VAT areas", () => {
  const out = want("outside-scope", 0, "No VAT charged: outside the scope of EU VAT (United States).");
  assertEquals(run(seller("NL"), consumer("US"), sup("goods")), want("zero-rated-export", 0, "Zero-rated: export of goods (United States)."));
  assertEquals(run(seller("NL"), { country: "US", isBusiness: true }, sup("services")), out);
  assertEquals(run(seller("NL"), consumer("US"), sup("services")), want("standard", 21, "VAT charged at 21% (Netherlands)."));
  assertEquals(run(seller("NL"), consumer("US"), sup("digital")), out);
  assertEquals(run(seller("NL"), consumer("GB"), sup("digital")), want("standard", 20, "VAT charged at 20% (United Kingdom)."));
  assertEquals(run(seller("NL"), consumer("GB"), sup("goods")).treatment, "zero-rated-export");
  assertEquals(
    run(seller("GB"), consumer("FR"), sup("digital")),
    want("standard", 20, "VAT charged at 20% (France) under the non-Union OSS scheme."),
  );
  assertEquals(
    run(seller("GB"), biz("FR"), sup("services")),
    want("outside-scope", 0, "No VAT charged: outside the scope of UK VAT (France)."),
  );
});

test("Northern Ireland", () => {
  assertEquals(
    run(seller("XI"), biz("IE"), sup("goods")),
    want("intra-community-supply", 0, "No VAT charged: intra-Community supply of goods (Northern Ireland to Ireland)."),
  );
  assertEquals(
    run(seller("XI"), biz("IE"), sup("services")),
    want("outside-scope", 0, "No VAT charged: outside the scope of UK VAT (Ireland)."),
  );
  assertEquals(
    run(seller("DE", { euDistanceSales: { currentYear: 50000, priorYear: 0 } }), consumer("XI"), sup("goods")),
    want("standard", 20, "VAT charged at 20% (Northern Ireland)."),
  );
});

test("normalisation, rates and errors", async () => {
  const gr = run(seller("NL"), consumer(" el ", { }), sup("services"), R);
  assertEquals(gr, want("standard", 21, "VAT charged at 21% (Netherlands)."));
  assertEquals(
    run(seller("GR"), consumer("GR"), sup("services")),
    want("standard", 24, "VAT charged at 24% (Greece)."),
  );
  assertEquals(run(seller("NL"), consumer(" el "), sup("goods"), { NL: 21 }).treatment, "standard");
  assertEquals(run(seller("nl"), consumer("nl"), sup("goods")).line, "VAT charged at 21% (Netherlands).");
  assertEquals(
    run(seller("NL", { euDistanceSales: { currentYear: 20000, priorYear: 0 } }), consumer(" el "), sup("goods")),
    want("standard", 24, "VAT charged at 24% (Greece) under the Union OSS scheme."),
  );

  const dest = () => run(seller("NL", { ossRegistered: true }), consumer("FI"), sup("digital"), { NL: 21 });
  await assertRejects(() => Promise.resolve().then(() => dest()), Error, "vatRates.FI");
  for (const bad of [-1, 100, "21", NaN]) {
    await assertRejects(() => Promise.resolve().then(() => run(seller("NL"), consumer("NL"), sup("goods"), { NL: bad })), Error, "vatRates.NL");
  }
  assertEquals(run(seller("NL"), consumer("NL"), sup("goods"), { NL: 0 }), want("standard", 0, "VAT charged at 0% (Netherlands)."));
  await assertRejects(() => Promise.resolve().then(() => run(seller("NL"), consumer("NL"), sup("goods"), null)), Error, "vatRates.NL");
  for (const bad of ["US", undefined, "Netherlands"]) {
    await assertRejects(() => Promise.resolve().then(() => run(seller("NL", { country: bad }), consumer("NL"), sup("goods"))), Error, "seller.country");
  }
  await assertRejects(() => Promise.resolve().then(() => run(seller("NL"), { isBusiness: false }, sup("goods"))), Error, "customer.country");
  await assertRejects(() => Promise.resolve().then(() => run(seller("NL"), consumer("NL"), sup("rental"))), Error, "supply.kind");
  await assertRejects(() => Promise.resolve().then(() => run(seller("NL"), consumer("NL"), sup("goods", { shipsFrom: "CN" }))), Error, "supply.shipsFrom");
});

// The node, published and resolved by an engine that only has the bundle.
const documentOf = () =>
  template(
    <Document title="Invoice">
      <VatPlaceOfSupply />
    </Document>,
  );

function paragraphs(nodes: DocumentNode[]): Array<[string | undefined, string | undefined]> {
  return nodes.map((n) => [n.id, (n as ParagraphNode).text]);
}

const cases: Array<[string, Obj]> = [
  ["case 8, note printed", { seller: seller("NL"), customer: biz("DE", { vatNumberValidated: false }), supply: sup("services"), vatRates: R }],
  ["case 5, reverse charge", { seller: seller("NL"), customer: biz("DE"), supply: sup("services"), vatRates: R }],
];

for (const [label, data] of cases) {
  test(`publishing then resolving matches building directly (${label})`, async () => {
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
    const resolved = await resolveDocument(published, state, { derivers: registry });
    const direct = await buildDocument(documentOf(), data, { dynamicMode: "placeholder" });

    assertEquals(paragraphs(resolved.nodes), paragraphs(direct.nodes));
    assertEquals(resolved.nodes.length, label.startsWith("case 8") ? 2 : 1);
  });
}
