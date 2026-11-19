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

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/canadian-sales-tax.deriver.ts")
>("components/derivers/canadian-sales-tax.deriver.ts");
const { CanadianSalesTax } = await loadRegistryModule<
  typeof import("../registry/components/canadian-sales-tax/canadian-sales-tax.node.tsx")
>("components/canadian-sales-tax/canadian-sales-tax.node.tsx");

type Obj = Record<string, unknown>;
const line = (amount: number, extra: Obj = {}) => ({ description: "Thing", amount, ...extra });
const baseTax = (): Obj => ({ supplyProvince: "QC", taxPointDate: "2026-11-30", paymentTerms: "Net 30", lines: [line(1000)] });
const baseSeller = (): Obj => ({
  name: "Seller Inc.",
  gstRegistered: true,
  bn: "843219767RT0001",
  provincialRegistered: true,
  qst: "1219876543TQ0001",
});
const baseCustomer = { name: "Buyer Ltd." };

/** The deriver throws synchronously; assertRejects wants a promise. */
const throws = (fn: () => unknown, _error: ErrorConstructor, message: string) =>
  assertRejects(() => Promise.resolve().then(fn), Error, message);
type Result = import("../registry/components/arrangements/canadian-sales-tax/lines.ts").CanadianSalesTaxLines;
const run = (tax: Obj = {}, seller: Obj = {}, customer: Obj | null = baseCustomer): Result =>
  mod.run([{ ...baseTax(), ...tax }, { ...baseSeller(), ...seller }, customer], {} as never) as Result;
const amounts = (r: Result) => [r.federalLabel, r.federalAmount, r.provincialLabel, r.provincialAmount, r.total];

test("Quebec: QST on the price before GST, never tax on tax", () => {
  assertEquals(amounts(run()), [
    "GST 5% on $1,000.00", "$50.00", "QST 9.975% on $1,000.00", "$99.75", "$1,149.75",
  ]);
});

test("Ontario is one HST line", () => {
  const r = run({ supplyProvince: "ON" });
  assertEquals(amounts(r), ["HST 13% on $1,000.00", "$130.00", null, null, "$1,130.00"]);
  assertEquals(r.qst, null);
});

test("rates are picked by the tax point date", () => {
  assertEquals(run({ supplyProvince: "NS", taxPointDate: "2025-03-31" }).total, "$1,150.00");
  assertEquals(run({ supplyProvince: "NS", taxPointDate: "2025-04-01" }).total, "$1,140.00");
  assertEquals(run({ supplyProvince: "NS", taxPointDate: "2025-04-01T23:59:00Z" }).federalLabel, "HST 14% on $1,000.00");
});

test("BC, SK and MB charge GST plus their own tax", () => {
  for (const [supplyProvince, label, tax] of [
    ["BC", "PST 7% on $1,000.00", "$70.00"],
    ["SK", "PST 6% on $1,000.00", "$60.00"],
    ["MB", "RST 7% on $1,000.00", "$70.00"],
  ]) {
    const r = run({ supplyProvince });
    assertEquals([r.federalAmount, r.provincialLabel, r.provincialAmount], ["$50.00", label, tax]);
  }
});

test("Alberta and the territories charge GST alone, and say nothing about provincial registration", () => {
  for (const supplyProvince of ["AB", "NT", "NU", "YT"]) {
    for (const provincialRegistered of [true, false, undefined]) {
      const r = run({ supplyProvince }, { provincialRegistered });
      assertEquals([r.federalAmount, r.provincialLabel, r.total, r.warning], ["$50.00", null, "$1,050.00", null]);
    }
  }
});

test("only the supply province counts: there is no seller province input", () => {
  // The seller object here could say anything about where it is; the sale is still taxed as Quebec's.
  assertEquals(run({}, { province: "ON" }).provincialLabel, "QST 9.975% on $1,000.00");
});

test("each tax has its own base", () => {
  const r = run({
    lines: [
      line(500),
      line(300, { description: "Books", provincialCategory: "exempt" }),
      line(200, { description: "Rent, long-term", gstCategory: "exempt" }),
    ],
  });
  assertEquals(amounts(r), [
    "GST 5% on $800.00", "$40.00", "QST 9.975% on $500.00", "$49.88", "$1,089.88",
  ]);
  assertEquals(r.statusNote, "No GST on: Rent, long-term. No QST on: Books; Rent, long-term.");
  // A blank description is named by its position.
  assertEquals(
    run({ lines: [line(10), { amount: 5, gstCategory: "exempt" }] }).statusNote,
    "No GST on: line 2. No QST on: line 2.",
  );
});

test("each tax is rounded once, not per line", () => {
  const r = run({ lines: [line(0.1), line(0.1), line(0.1)] });
  assertEquals([r.subtotal, r.federalAmount, r.provincialAmount], ["$0.30", "$0.02", "$0.03"]);
});

test("a half cent rounds away from zero without a float slip", () => {
  const r = run({ lines: [line(20)] });
  assertEquals([r.federalAmount, r.provincialAmount, r.total], ["$1.00", "$2.00", "$23.00"]);
  const credit = run({ lines: [line(-20)] });
  assertEquals([credit.federalAmount, credit.provincialAmount, credit.total], ["-$1.00", "-$2.00", "-$23.00"]);
});

test("a credit note takes its tier from the absolute total", () => {
  assertEquals(run({ lines: [line(-200)] }, { bn: undefined }).warning?.includes("$100 or more"), true);
});

test("no lines, and fully zero-rated lines, still print their tax lines at zero", () => {
  const empty = run({ lines: [] });
  assertEquals([empty.subtotal, empty.federalAmount, empty.provincialAmount, empty.total, empty.statusNote, empty.warning], [
    "$0.00", "$0.00", "$0.00", "$0.00", null, null,
  ]);
  const zero = run({ lines: [line(100, { gstCategory: "zero-rated" })] });
  assertEquals(amounts(zero), ["GST 5% on $0.00", "$0.00", "QST 9.975% on $0.00", "$0.00", "$100.00"]);
  assertEquals(zero.statusNote, null);
  assertEquals(run({ lines: [{ amount: 100 }] }).federalAmount, "$5.00");
});

test("a seller that is not registered charges no GST/HST and says so", () => {
  for (const gstRegistered of [false, undefined]) {
    const r = run({ supplyProvince: "ON" }, { gstRegistered });
    assertEquals([r.federalLabel, r.bn, r.total], [null, null, "$1,000.00"]);
    assertEquals(r.smallSupplierNote, "Not registered for GST/HST (small supplier), so no HST is charged.");
  }
  const qc = run({}, { gstRegistered: false });
  assertEquals([qc.federalLabel, qc.provincialAmount, qc.smallSupplierNote], [
    null, "$99.75", "Not registered for GST/HST (small supplier), so no GST is charged.",
  ]);
});

test("provincial registration", () => {
  const bc = run({ supplyProvince: "BC" }, { provincialRegistered: undefined });
  assertEquals(bc.provincialLabel, null);
  assertEquals(bc.warning, "It is not stated whether the seller is registered for PST, so none has been charged.");
  const no = run({}, { provincialRegistered: false });
  assertEquals([no.provincialLabel, no.qst, no.warning], [null, null, null]);
});

test("the GST/HST number is checked before it is printed", () => {
  assertEquals(run({}, { bn: "843 219 767 rt 0001" }).bn, "843219767RT0001");
  for (const bn of ["123456789RT0001", "843219767RT0000", "843219767RC0001"]) {
    const r = run({}, { bn });
    assertEquals([r.bn, r.federalAmount], [null, "$50.00"]);
    assertEquals(r.warning?.startsWith(bn + " is not a valid GST/HST registration number"), true);
  }
});

test("the QST number is checked for format", () => {
  const r = run({}, { qst: "12345TQ0001" });
  assertEquals(r.qst, null);
  assertEquals(r.warning?.startsWith("12345TQ0001 is not a valid QST registration number"), true);
  assertEquals(run({}, { qst: "1219876543 tq 0001" }).qst, "1219876543TQ0001");
});

test("information tiers on the tax-included total", () => {
  const ab = (amount: number, seller: Obj = {}, tax: Obj = {}, customer: Obj | null = baseCustomer) =>
    run({ supplyProvince: "AB", lines: [line(amount)], ...tax }, seller, customer);
  assertEquals(ab(95.23).total, "$99.99");
  assertEquals(ab(95.23, { bn: undefined }).warning, null);
  assertEquals(ab(95.24).total, "$100.00");
  assertEquals(
    ab(95.24, { bn: undefined }).warning,
    "Missing for an invoice of $100 or more: GST/HST registration number. Until it is added the customer cannot claim input tax credits on this invoice.",
  );
  const noExtras = { paymentTerms: undefined, lines: [{ amount: 476.19 }] };
  assertEquals(ab(476.18, {}, { paymentTerms: undefined, lines: [{ amount: 476.18 }] }, null).warning, null);
  const at500 = ab(0, {}, noExtras, null);
  assertEquals(at500.total, "$500.00");
  assertEquals(
    at500.warning?.startsWith("Missing for an invoice of $500 or more: customer name, payment terms, a description for every line."),
    true,
  );
  assertEquals(run({ lines: [line(500)] }, {}, null).warning?.includes("customer name"), true);
});

test("a missing seller name warns at every tier", () => {
  assertEquals(
    run({ lines: [line(10)] }, { name: " " }).warning,
    "Missing for an invoice of under $100: seller name. Until it is added the customer cannot claim input tax credits on this invoice.",
  );
});

test("many lines add up exactly in cents, with names in input order", () => {
  const lines = Array.from({ length: 50 }, (_, i) => line(0.1, { description: "L" + i, gstCategory: i % 2 ? "exempt" : "taxable" }));
  const r = run({ lines });
  assertEquals(r.subtotal, "$5.00");
  assertEquals(r.statusNote?.startsWith("No GST on: L1; L3; L5;"), true);
});

test("custom rates replace the built-in table", async () => {
  const rates = [
    { tax: "GST", rate: 5, from: "2008-01-01" },
    { tax: "QST", province: "QC", rate: 10, from: "2026-01-01" },
  ];
  assertEquals(run({ rates }).provincialAmount, "$100.00");
  // Nothing is merged in, so Ontario has no HST row.
  await throws(() => run({ rates, supplyProvince: "ON" }), Error, "No HST rate for ON");
});

test("bad rates throw, naming the tax, province and date", async () => {
  const gst = { tax: "GST", rate: 5, from: "2008-01-01" };
  const qst = { tax: "QST", province: "QC", rate: 9.975, from: "2013-01-01" };
  await throws(() => run({ rates: [gst, { ...qst, to: "2020-01-01" }] }), Error, "No QST rate for QC on 2026-11-30");
  await throws(() => run({ rates: [gst, qst, { ...qst, from: "2026-01-01" }] }), Error, "Overlapping QST rows for QC");
  await throws(() => run({ rates: [gst, { ...qst, rate: 9.9751 }] }), Error, "salesTax.rates");
  await throws(() => run({ rates: [gst, { ...qst, rate: -1 }] }), Error, "salesTax.rates");
  await throws(() => run({ supplyProvince: "NB", taxPointDate: "2016-06-30" }), Error, "No HST rate for NB");
});

test("bad input throws, naming the field", async () => {
  await throws(() => run({ supplyProvince: "QB" }), Error, "salesTax.supplyProvince");
  await throws(() => run({ taxPointDate: "30/11/2026" }), Error, "salesTax.taxPointDate");
  await throws(() => run({ taxPointDate: undefined }), Error, "salesTax.taxPointDate");
  await throws(() => run({ taxPointDate: "2012-12-31" }), Error, "2013-04-01");
  await throws(() => run({ lines: undefined }), Error, "salesTax.lines");
  await throws(() => run({ lines: [line(NaN)] }), Error, "amount");
  await throws(() => run({ lines: [line(1, { gstCategory: "taxed" })] }), Error, "gstCategory");
  await throws(() => run({ lines: [line(1, { provincialCategory: "zero-rated" })] }), Error, "provincialCategory");
});

// The node, published and resolved by an engine that only has the bundle.
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ canadianSalesTax: mod.run }));
const ARRANGEMENTS = ["table", "stacked"] as const;
const documentOf = (arrangement?: (typeof ARRANGEMENTS)[number]) =>
  template(
    <Document title="Invoice">
      <CanadianSalesTax arrangement={arrangement} />
    </Document>,
  );

/** Every line printed, in reading order, whatever the arrangement nests it in. */
function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { title?: string; text?: string; children?: DocumentNode[] };
    return [...(own.text ? [own.text] : []), ...texts(own.children ?? [])];
  });
}
const ids = (nodes: DocumentNode[]): string[] =>
  nodes.flatMap((n) => [n.id as string, ...ids((n as { children?: DocumentNode[] }).children ?? [])]);

const qcData = { salesTax: baseTax(), seller: baseSeller(), customer: baseCustomer };
const smallData = { salesTax: { ...baseTax(), supplyProvince: "ON", lines: [line(100), line(50, { description: "Rent", gstCategory: "exempt" })] }, seller: { name: "S", gstRegistered: false }, customer: baseCustomer };

for (const arrangement of ARRANGEMENTS) {
  for (const [name, data] of [["QC baseline", qcData], ["ON small supplier", smallData]] as const) {
    test(`${arrangement}, ${name}: publishing then resolving matches building directly`, async () => {
      const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof qcData, {
        branchMode: "publish",
        deriverMode: "preserve",
        dynamicMode: "resolve",
        aiClient: { generateParagraph: () => "" },
      });
      // The conditional rows are conditions on derived values, not baked in.
      const json = JSON.stringify(published);
      for (const id of ["federal", "provincial", "bn", "qst", "small-supplier", "status", "warning"]) {
        assertEquals(json.includes("canadian-sales-tax-" + id), true);
      }
      assertEquals(json.split('"type":"truthy"').length - 1 >= 7, true);
      const state: RuntimeState = {
        ctx: { ...data },
        derived: {},
        dataProvider: new InMemoryDataProvider(data),
        aiClient: { generateParagraph: () => "" },
      };
      const resolved = await resolveDocument(published, state, { derivers: registry });
      const direct = await buildDocument(documentOf(arrangement), data as typeof qcData, { dynamicMode: "placeholder" });
      assertEquals(texts(resolved.nodes), texts(direct.nodes));
      assertEquals(
        texts(resolved.nodes).some((t) => t.includes("$1,149.75") || t.includes("small supplier")),
        true,
      );
    });
  }
}

test("both arrangements print the same ids, placed differently", async () => {
  const build = async (a?: (typeof ARRANGEMENTS)[number]) =>
    (await buildDocument(documentOf(a), qcData, { dynamicMode: "placeholder" })).nodes;
  const [table, stacked] = [await build("table"), await build("stacked")];
  const shared = ["canadian-sales-tax-subtotal", "canadian-sales-tax-federal", "canadian-sales-tax-provincial", "canadian-sales-tax-total", "canadian-sales-tax-bn", "canadian-sales-tax-qst"];
  for (const id of shared) {
    assertEquals(ids(table).includes(id) && ids(stacked).includes(id), true);
  }
  assertEquals(JSON.stringify(table) === JSON.stringify(stacked), false);
  assertEquals(texts(stacked).slice(0, 2), ["Subtotal: $1,000.00", "GST 5% on $1,000.00: $50.00"]);
  assertEquals((await build()), table);
});
