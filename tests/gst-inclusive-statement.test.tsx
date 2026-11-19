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
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";
import { COMPONENTS } from "../src/registry/catalog.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/gst-inclusive-statement.deriver.ts")
>("components/derivers/gst-inclusive-statement.deriver.ts");
const { GstInclusiveStatement } = await loadRegistryModule<
  typeof import("../registry/components/gst-inclusive-statement/gst-inclusive-statement.node.tsx")
>("components/gst-inclusive-statement/gst-inclusive-statement.node.tsx");

type Obj = Record<string, unknown>;
interface Result {
  heading: string;
  sellerTaxId: string | null;
  buyerName: string | null;
  buyerTaxId: string | null;
  buyerIdentifier: string | null;
  items: Array<{ description: string; amount: string }>;
  subtotal: string | null;
  gst: string | null;
  gstLabel: string;
  total: string;
  totalLabel: string;
  statement: string | null;
  legend: string | null;
  marginNote: string | null;
}

const run = (gst: unknown) => mod.run([gst], {} as RuntimeState) as Result;
const ABN = "51 824 753 556";
const line = (amountMinor: number, treatment = "taxable", description = "Item") => ({ description, amountMinor, treatment });
const au = (lines: unknown[], extra: Obj = {}) => ({
  country: "AU",
  sellerRegistered: true,
  sellerTaxId: ABN,
  pricesIncludeGst: true,
  lines,
  ...extra,
});
const nz = (lines: unknown[], extra: Obj = {}) => ({
  country: "NZ",
  sellerRegistered: true,
  sellerTaxId: "49091850",
  pricesIncludeGst: true,
  lines,
  ...extra,
});
const throws = (gst: unknown, includes = "") => {
  let message: string | undefined;
  try {
    run(gst);
  } catch (error) {
    message = (error as Error).message;
  }
  if (message === undefined || !message.includes(includes)) {
    throw new Error(`expected a throw including "${includes}", got ${message}`);
  }
};

test("one taxable item, inclusive: the statement replaces the figure", () => {
  const r = run(au([line(8250)]));
  assertEquals(
    [r.heading, r.gst, r.gstLabel, r.statement, r.total, r.totalLabel, r.legend, r.items[0].description],
    ["Tax invoice", null, "", "Total price includes GST.", "$82.50", "Total", null, "Item"],
  );
  assertEquals(r.sellerTaxId, "ABN 51 824 753 556");
});

test("several taxable items under total-invoice", () => {
  const r = run(au([line(11000), line(5500)]));
  assertEquals([r.statement, r.total], ["Total price includes GST.", "$165.00"]);
});

test("taxable-sale drift suppresses the statement; total-invoice does not", () => {
  const lines = [line(1655), line(1655), line(1655)];
  const r = run(au(lines, { rounding: "taxable-sale" }));
  assertEquals([r.gst, r.gstLabel, r.total, r.totalLabel, r.statement, r.legend], [
    "$4.50",
    "GST included",
    "$49.65",
    "Total including GST",
    null,
    null,
  ]);
  assertEquals(run(au(lines)).statement, "Total price includes GST.");
});

test("taxable-sale that is exact keeps the statement", () => {
  assertEquals(run(au([line(1650), line(1650)], { rounding: "taxable-sale" })).statement, "Total price includes GST.");
});

test("exclusive basis never gets the statement, and the rules round differently", () => {
  const lines = [line(1005), line(1005), line(1005)];
  const a = run(au(lines, { pricesIncludeGst: false }));
  assertEquals([a.subtotal, a.gst, a.gstLabel, a.total, a.totalLabel, a.statement], [
    "$30.15",
    "$3.02",
    "GST",
    "$33.17",
    "Total including GST",
    null,
  ]);
  const b = run(au(lines, { pricesIncludeGst: false, rounding: "taxable-sale" }));
  assertEquals([b.gst, b.total, b.statement], ["$3.03", "$33.18", null]);
});

test("many items: 25 x 105 exclusive", () => {
  const lines = Array.from({ length: 25 }, () => line(105));
  const x = { pricesIncludeGst: false };
  const a = run(au(lines, x));
  assertEquals([a.gst, a.total], ["$2.63", "$28.88"]);
  const b = run(au(lines, { ...x, rounding: "taxable-sale" }));
  assertEquals([b.gst, b.total], ["$2.75", "$29.00"]);
});

test("mixed AU: taxable items are marked and the legend matches the basis", () => {
  const lines = [line(22000, "taxable", "Fit-out"), line(5000, "gst-free", "Produce")];
  const r = run(au(lines));
  assertEquals([r.gst, r.gstLabel, r.total, r.totalLabel, r.statement], [
    "$20.00",
    "GST included",
    "$270.00",
    "Total including GST",
    null,
  ]);
  assertEquals(r.items.map((i) => i.description), ["Fit-out *", "Produce"]);
  assertEquals(r.legend, "* Taxable supply. The price of each marked item includes GST.");
  assertEquals(
    run(au(lines, { pricesIncludeGst: false })).legend,
    "* Taxable supply. GST is charged on each marked item.",
  );
});

test("mixed NZ, exclusive", () => {
  const r = run(nz([line(20000), line(5000, "zero-rated")], { pricesIncludeGst: false }));
  assertEquals([r.subtotal, r.gst, r.total, r.sellerTaxId], ["$250.00", "$30.00", "$280.00", "GST number 49-091-850"]);
});

test("NZ all taxable: statement, and the seller number is required above NZ$200", () => {
  assertEquals(run(nz([line(23000)])).statement, "Total price includes GST.");
  throws(nz([line(23000)], { sellerTaxId: undefined }), "sellerTaxId");
});

test("NZ seller threshold sits at NZ$200 exactly", () => {
  run(nz([line(20000)], { sellerTaxId: undefined }));
  throws(nz([line(20001)], { sellerTaxId: undefined }), "sellerTaxId");
});

test("NZ buyer threshold sits at NZ$1,000 exactly", () => {
  run(nz([line(100000)]));
  throws(nz([line(100001)], { buyer: { name: "Kiwi Ltd" } }), "identifier");
  throws(nz([line(100001)]), "name");
  const email = { kind: "email", value: "accounts@example.co.nz" };
  const r = run(nz([line(100001)], { buyer: { name: "Kiwi Ltd", identifier: email } }));
  assertEquals([r.buyerName, r.buyerIdentifier], ["Buyer: Kiwi Ltd", "Buyer email: accounts@example.co.nz"]);
  run(nz([line(100001)], { buyer: { name: "Kiwi Ltd", identifier: { kind: "nzbn", value: "9429041535134" } } }));
  throws(
    nz([line(100001)], { buyer: { name: "K", identifier: { kind: "nzbn", value: "9429041535135" } } }),
    "NZBN",
  );
  throws(nz([line(100)], { buyer: { identifier: { kind: "fax", value: "1" } } }), "kind");
});

test("AU buyer threshold sits at A$1,000", () => {
  run(au([line(99999)]));
  throws(au([line(100000)]), "buyer");
  run(au([line(100000)], { buyer: { abn: "98 765 432 122" } }));
  const r = run(au([line(100000)], { buyer: { name: "Acme" } }));
  assertEquals(r.buyerName, "Buyer: Acme");
  throws(au([line(100000)], { buyer: { name: "  " } }), "buyer");
});

test("the threshold is tested on the inclusive total, not the lines", () => {
  // 90,000 exclusive becomes 99,000 with GST: still under. 91,000 becomes 100,100.
  run(au([line(90000)], { pricesIncludeGst: false }));
  throws(au([line(91000)], { pricesIncludeGst: false }), "buyer");
});

test("checksums", () => {
  throws(au([line(100)], { sellerTaxId: "51 824 753 557" }), "ABN");
  throws(au([line(100)], { buyer: { abn: "51 824 753 557" } }), "ABN");
  assertEquals(run(nz([line(100)], { sellerTaxId: "10000191" })).sellerTaxId, "GST number 10-000-191");
  assertEquals(run(nz([line(100)], { sellerTaxId: "136-410-132" })).sellerTaxId, "GST number 136-410-132");
  for (const id of ["35901981", "49098576"]) run(nz([line(100)], { sellerTaxId: id }));
  for (const id of ["136410133", "9125568", "12345678"]) throws(nz([line(100)], { sellerTaxId: id }), "IRD");
});

test("not registered: plain invoice, no GST wording, ABN still quoted", () => {
  const lines = [line(11000), line(5000, "gst-free")];
  for (const flag of [{ sellerRegistered: false }, { sellerRegistered: undefined }]) {
    const r = run({ country: "AU", lines, ...flag, buyer: { name: "Acme" } });
    assertEquals(
      [r.heading, r.gst, r.gstLabel, r.statement, r.legend, r.totalLabel, r.total, r.sellerTaxId, r.buyerName],
      ["Invoice", null, "", null, null, "Total", "$160.00", null, "Buyer: Acme"],
    );
    assertEquals(r.items.map((i) => i.description), ["Item", "Item"]);
  }
  assertEquals(run({ country: "AU", lines, sellerTaxId: ABN }).sellerTaxId, "ABN 51 824 753 556");
  assertEquals(run({ country: "NZ", lines: [line(100)], sellerTaxId: "49091850" }).sellerTaxId, null);
  // Thresholds do not apply when no GST is claimed.
  run({ country: "AU", lines: [line(500000)] });
});

test("registered with no taxable line", () => {
  const r = run({ country: "AU", sellerRegistered: true, lines: [line(500000, "gst-free")] });
  assertEquals([r.heading, r.statement, r.gst, r.total, r.sellerTaxId], [
    "Invoice",
    "No GST is charged on this invoice.",
    null,
    "$5,000.00",
    null,
  ]);
});

test("zero", () => {
  const r = run(au([line(0)]));
  assertEquals([r.total, r.statement], ["$0.00", "Total price includes GST."]);
});

test("negative lines", () => {
  const r = run(au([line(11000), line(-1100)]));
  assertEquals([r.total, r.statement, r.items[1].amount], ["$99.00", "Total price includes GST.", "-$11.00"]);
  throws(au([line(1000), line(-2000)]), "negative");
  throws(au([line(-1000, "taxable"), line(5000, "gst-free")]), "negative");
});

test("bad fields throw", () => {
  throws(null, "gst");
  throws(au([]), "lines");
  throws({ ...au([line(1)]), lines: undefined }, "lines");
  throws(au([line(1, "taxable", "  ")]), "description");
  throws(au([line(10.5)]), "amountMinor");
  throws(au([line(1, "zero-rated")]), "treatment");
  throws(nz([line(1, "gst-free")]), "treatment");
  throws(nz([line(1, "margin-scheme")]), "treatment");
  throws(au([line(1)], { rounding: "nearest" }), "rounding");
  throws(au([line(1)], { country: "UK" }), "country");
  throws(au([line(1)], { pricesIncludeGst: undefined }), "pricesIncludeGst");
  throws(au([line(1e13), line(1)]), "too large");
});

test("margin scheme: no GST on the line, and a note", () => {
  const r = run(au([line(11000), line(50000, "margin-scheme")]));
  assertEquals([r.gst, r.items[0].description, r.marginNote], [
    "$10.00",
    "Item *",
    "Items sold under the margin scheme carry no GST you can claim as a credit.",
  ]);
  assertEquals(run(au([line(11000)])).marginNote, null);
});

test("the catalog preview prints the expected figures", () => {
  const entry = COMPONENTS.find((c) => c.id === "gst-inclusive-statement")!;
  const r = run((entry.previewData as Obj).gst);
  assertEquals([r.heading, r.gst, r.total, r.totalLabel, r.buyerTaxId], [
    "Tax invoice",
    "$11,537.27",
    "$140,000.00",
    "Total including GST",
    "Buyer ABN 98 765 432 122",
  ]);
  assertEquals(r.items.filter((i) => i.description.endsWith(" *")).length, 2);
});

// The node, published and resolved by an engine that only has the bundle.
const ARRANGEMENTS = ["table", "stacked"] as const;
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ gstInclusiveStatement: mod.run }));

const documentOf = (arrangement?: (typeof ARRANGEMENTS)[number]) =>
  template(
    <Document title="Invoice">
      <GstInclusiveStatement arrangement={arrangement} />
    </Document>,
  );

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { title?: string; text?: string; children?: DocumentNode[] };
    return [
      ...(node.kind === "section" && own.title ? [own.title] : []),
      ...(own.text ? [own.text] : []),
      ...texts(own.children ?? []),
    ];
  });
}

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[]; rows?: DocumentNode[] };
    return [...(own.id ? [own.id] : []), ...ids(own.children ?? []), ...ids(own.rows ?? [])];
  });
}

const mixed = au([line(22000, "taxable", "Fit-out"), line(5000, "gst-free", "Produce")]);
const unregistered = { country: "AU", lines: [line(11000), line(5000, "gst-free")] };

for (const arrangement of ARRANGEMENTS) {
  for (const [name, gst, legend] of [["mixed", mixed, true], ["unregistered", unregistered, false]] as const) {
    test(`${arrangement}, ${name}: publishing then resolving matches building directly`, async () => {
      const data = { gst };
      const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof data, {
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
      const direct = await buildDocument(documentOf(arrangement), data, { dynamicMode: "placeholder" });
      const found = ids(resolved.nodes);

      assertEquals(texts(resolved.nodes), texts(direct.nodes));
      assertEquals(found.filter((id) => id.startsWith("gst-item")).length, 2);
      assertEquals(found.includes("gst-legend"), legend);
    });
  }
}

test("the arrangements differ only in placement", async () => {
  const entry = COMPONENTS.find((c) => c.id === "gst-inclusive-statement")!;
  const idsOf = async (arrangement: (typeof ARRANGEMENTS)[number]) =>
    ids((await buildDocument(documentOf(arrangement), entry.previewData as never, { dynamicMode: "placeholder" })).nodes)
      .filter((id) => id.startsWith("gst-") && id !== "gst-lines" && id !== "gst-lines-header")
      .map((id) => id.replace(/-\d+$/, ""));

  assertEquals(await idsOf("table"), await idsOf("stacked"));
});

test("a bad checksum fails the build of the node", async () => {
  await assertRejects(
    () => buildDocument(documentOf(), { gst: au([line(100)], { sellerTaxId: "51 824 753 557" }) }, { dynamicMode: "placeholder" }),
    Error,
    "ABN",
  );
});
