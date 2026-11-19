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
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/invoice-document-type.deriver.ts")
>("components/derivers/invoice-document-type.deriver.ts");
const { InvoiceDocumentType } = await loadRegistryModule<
  typeof import("../registry/components/invoice-document-type/invoice-document-type.node.tsx")
>("components/invoice-document-type/invoice-document-type.node.tsx");
// Run through a bundle, so the test proves the function survives being shipped
// as source text, which is how a deriver fails.
const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ invoiceDocumentType: mod.run }),
);

type Obj = Record<string, unknown>;
const baseInvoice: Obj = { number: "1", issueDate: "2026-03-01", kind: "invoice", jurisdiction: "UK", grossTotal: 1000 };
const baseSeller: Obj = { vatRegistered: true, vatNumber: "GB1" };

interface Result {
  heading: string;
  statement: string | null;
  issueDate: string;
  supplyLine: string | null;
  supplyLabel: string | null;
  supplyValue: string | null;
  vatNumber: string | null;
}

async function run(invoice: Obj = {}, seller: Obj | undefined = {}): Promise<Result> {
  const inv = { ...baseInvoice, ...invoice };
  const sel = seller === undefined ? undefined : { ...baseSeller, ...seller };
  const data = { invoice: inv, seller: sel };
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
  return await registry.run(
    {
      name: "invoiceDocumentType",
      output: "invoiceType",
      inputs: [dataRef("invoice"), dataRef("seller")],
    } as never,
    state,
  ) as Result;
}

const heading = async (invoice: Obj, seller?: Obj) => (await run(invoice, seller)).heading;
const allowed = { simplifiedAllowed: true };

test("defaults give a VAT invoice with the number and a long date", async () => {
  assertEquals(await run(), {
    heading: "VAT invoice",
    statement: null,
    issueDate: "1 March 2026",
    supplyLine: null,
    supplyLabel: null,
    supplyValue: null,
    vatNumber: "GB1",
  });
});

test("a seller who is not registered, or unknown, gets a plain invoice", async () => {
  assertEquals(await run({}, { vatRegistered: false }), {
    heading: "Invoice",
    statement: null,
    issueDate: "1 March 2026",
    supplyLine: null,
    supplyLabel: null,
    supplyValue: null,
    vatNumber: null,
  });
  assertEquals(await run({}, { vatRegistered: undefined }).then((r) => r.heading), "Invoice");

  const none = await registry.run(
    { name: "invoiceDocumentType", output: "o", inputs: [dataRef("invoice"), dataRef("seller")] } as never,
    { ctx: {}, derived: {}, dataProvider: new InMemoryDataProvider({ invoice: baseInvoice }), aiClient: { generateParagraph: () => "" } },
  ) as Result;
  assertEquals([none.heading, none.vatNumber], ["Invoice", null]);
});

test("a registered seller needs a VAT number", async () => {
  await assertRejects(() => run({}, { vatNumber: undefined }), Error, "seller.vatNumber");
  await assertRejects(() => run({}, { vatNumber: "  " }), Error, "seller.vatNumber");
});

test("simplified invoices follow the ceiling, rounded to the printed total", async () => {
  for (const [gross, want] of [
    [250, "Simplified VAT invoice"],
    [250.004, "Simplified VAT invoice"],
    [250.01, "VAT invoice"],
    [0, "Simplified VAT invoice"],
    [-40, "Simplified VAT invoice"],
    [undefined, "VAT invoice"],
    [NaN, "VAT invoice"],
    ["100", "VAT invoice"],
  ] as const) {
    assertEquals(await heading({ grossTotal: gross }, allowed), want);
  }
  assertEquals(await heading({ grossTotal: 10 }), "VAT invoice");
});

test("EU and DE ceilings", async () => {
  assertEquals(await heading({ jurisdiction: "EU", grossTotal: 100 }, allowed), "Simplified VAT invoice");
  assertEquals(await heading({ jurisdiction: "EU", grossTotal: 100.01 }, allowed), "VAT invoice");
  assertEquals(await heading({ jurisdiction: "DE", grossTotal: 250 }, allowed), "Simplified VAT invoice");
});

test("US has no VAT rules", async () => {
  const r = await run({ jurisdiction: "US" }, { vatNumber: undefined });
  assertEquals([r.heading, r.vatNumber], ["Invoice", null]);
});

test("pro forma", async () => {
  const r = await run({ kind: "proforma" });
  assertEquals([r.heading, r.statement], ["Pro forma invoice", "This is not a VAT invoice."]);
  const n = await run({ kind: "proforma" }, { vatRegistered: false });
  assertEquals([n.heading, n.statement], ["Pro forma invoice", null]);
  assertEquals(await heading({ kind: "proforma", grossTotal: 10 }, allowed), "Pro forma invoice");
});

test("self-billed statements by jurisdiction", async () => {
  for (const [jurisdiction, statement] of [
    ["UK", "The VAT shown is your output tax due to HMRC."],
    ["DE", "Gutschrift (self-billing)."],
    ["FR", "Autofacturation (self-billing)."],
    ["EU", "Self-billing."],
  ]) {
    const r = await run({ kind: "self-billed", jurisdiction });
    assertEquals([r.heading, r.statement], ["Self-billed invoice", statement]);
  }
  assertEquals((await run({ kind: "self-billed" }, { vatRegistered: false })).statement, null);
  const us = await run({ kind: "self-billed", jurisdiction: "US" });
  assertEquals([us.heading, us.statement], ["Self-billed invoice", null]);
});

test("date of supply", async () => {
  assertEquals((await run({ supplyDate: "2026-02-27" })).supplyLine, "Date of supply: 27 February 2026");
  assertEquals((await run({ supplyDate: "2026-03-01T15:00:00Z" })).supplyLine, null);
  assertEquals((await run({ supplyDate: "" })).supplyLine, null);
  await assertRejects(() => run({ supplyDate: "next week" }), Error, "invoice.supplyDate");
});

test("the supply line comes apart into a label and a value", async () => {
  const pick = (r: Result) => [r.supplyLabel, r.supplyValue, r.supplyLine];
  assertEquals(pick(await run({ supplyDate: "2026-02-27" })), [
    "Date of supply",
    "27 February 2026",
    "Date of supply: 27 February 2026",
  ]);
  assertEquals(pick(await run({ jurisdiction: "DE" })), [
    "Leistungsdatum",
    "entspricht Rechnungsdatum",
    "Leistungsdatum entspricht Rechnungsdatum.",
  ]);
  assertEquals(pick(await run({ supplyPeriod: { from: "2026-02-01", to: "2026-02-28" } })), [
    "Supply period",
    "1 February 2026 to 28 February 2026",
    "Supply period: 1 February 2026 to 28 February 2026",
  ]);
});

test("Germany always prints a supply line", async () => {
  const same = "Leistungsdatum entspricht Rechnungsdatum.";
  assertEquals((await run({ jurisdiction: "DE" })).supplyLine, same);
  assertEquals((await run({ jurisdiction: "DE", supplyDate: "2026-03-01" })).supplyLine, same);
  assertEquals(
    (await run({ jurisdiction: "DE", supplyDate: "2026-02-27" })).supplyLine,
    "Leistungsdatum: 27 February 2026",
  );
});

test("supply period", async () => {
  const supplyPeriod = { from: "2026-02-01", to: "2026-02-28" };
  assertEquals((await run({ supplyPeriod })).supplyLine, "Supply period: 1 February 2026 to 28 February 2026");
  assertEquals(
    (await run({ supplyPeriod, jurisdiction: "DE" })).supplyLine,
    "Leistungszeitraum: 1 February 2026 bis 28 February 2026",
  );
  assertEquals(
    (await run({ supplyPeriod, supplyDate: "2026-02-27" })).supplyLine,
    "Supply period: 1 February 2026 to 28 February 2026",
  );
  await assertRejects(() => run({ supplyPeriod: { from: "2026-02-28", to: "2026-02-01" } }), Error, "supplyPeriod");
  await assertRejects(() => run({ supplyPeriod: { from: "2026-02-01" } }), Error, "supplyPeriod");
  await assertRejects(() => run({ supplyPeriod: { to: "2026-02-01" } }), Error, "supplyPeriod");
});

test("US dates", async () => {
  assertEquals((await run({ jurisdiction: "US" })).issueDate, "March 1, 2026");
});

test("required fields and enums", async () => {
  await assertRejects(() => run({ number: undefined }), Error, "invoice.number");
  await assertRejects(() => run({ number: "" }), Error, "invoice.number");
  await assertRejects(() => run({ issueDate: undefined }), Error, "invoice.issueDate");
  await assertRejects(() => run({ issueDate: "01/03/2026" }), Error, "invoice.issueDate");
  await assertRejects(() => run({ kind: "credit-note" }), Error, "invoice.kind");
  await assertRejects(() => run({ jurisdiction: "AU" }), Error, "invoice.jurisdiction");
});

// The node, published and resolved by an engine that only has the bundle.
const nodeData = {
  invoice: { ...baseInvoice, supplyDate: "2026-02-27", purchaseOrder: "PO-7" },
  seller: baseSeller,
};

const ARRANGEMENTS = ["stacked", "table", "split"] as const;

function documentOf(arrangement?: (typeof ARRANGEMENTS)[number]) {
  return template(
    <Document title="Invoice">
      <InvoiceDocumentType arrangement={arrangement} />
    </Document>,
  );
}

/** Every line printed, in reading order, whatever the arrangement nests it in. */
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

for (const arrangement of ARRANGEMENTS) {
  for (const purchaseOrder of ["PO-7", "", undefined]) {
    test(`${arrangement}: publishing then resolving matches building directly (purchase order ${JSON.stringify(purchaseOrder)})`, async () => {
      const data = { ...nodeData, invoice: { ...nodeData.invoice, purchaseOrder } };
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

      assertEquals(texts(resolved.nodes), texts(direct.nodes));
      assertEquals(
        texts(direct.nodes).some((line) => line.startsWith("Purchase order")),
        Boolean(purchaseOrder),
      );
    });
  }
}

test("every arrangement prints the same things, placed differently", async () => {
  const data = { ...nodeData, seller: { ...baseSeller, vatRegistered: true } };
  const lines = async (arrangement: (typeof ARRANGEMENTS)[number]) =>
    texts((await buildDocument(documentOf(arrangement), data, { dynamicMode: "placeholder" })).nodes);

  assertEquals(await lines("stacked"), [
    "VAT invoice",
    "Invoice number: 1",
    "Date of issue: 1 March 2026",
    "Date of supply: 27 February 2026",
    "Purchase order: PO-7",
    "VAT number: GB1",
  ]);
  assertEquals(await lines("table"), [
    "VAT invoice",
    "Invoice number",
    "1",
    "Date of issue",
    "1 March 2026",
    "Date of supply",
    "27 February 2026",
    "Purchase order",
    "PO-7",
    "VAT number",
    "GB1",
  ]);
  assertEquals(await lines("split"), [
    "VAT invoice",
    "Invoice number: 1",
    "Purchase order: PO-7",
    "VAT number: GB1",
    "Date of issue: 1 March 2026",
    "Date of supply: 27 February 2026",
  ]);
});

test("no arrangement is the stacked one", async () => {
  const build = (arrangement?: "stacked") =>
    buildDocument(documentOf(arrangement), nodeData, { dynamicMode: "placeholder" });
  assertEquals((await build()).nodes, (await build("stacked")).nodes);
});
