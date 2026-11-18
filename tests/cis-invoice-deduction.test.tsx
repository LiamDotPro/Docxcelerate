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
  typeof import("../registry/components/derivers/cis-invoice-deduction.deriver.ts")
>("components/derivers/cis-invoice-deduction.deriver.ts");
const { CisInvoiceDeduction } = await loadRegistryModule<
  typeof import("../registry/components/cis-invoice-deduction/cis-invoice-deduction.node.tsx")
>("components/cis-invoice-deduction/cis-invoice-deduction.node.tsx");

type Obj = Record<string, unknown>;
type Result = ReturnType<NonNullable<typeof mod.__types>>;

const base: Obj = { grossExVat: 10000, materialsCost: 2500, plantHire: 500, vatRate: 0.2 };
const net: Obj = { vatRegistered: true, cisStatus: "net" };
const notCis: Obj = { vatRegistered: true, cisRegistered: false };
const both: Obj = { vatRegistered: true, cisRegistered: true };

const run = (invoice: Obj = {}, supplier: Obj = {}, customer: Obj = both) =>
  mod.run([{ ...base, ...invoice }, { ...net, ...supplier }, customer], {} as RuntimeState) as Result;

const rejects = (invoice: Obj = {}, supplier: Obj = {}, customer: Obj = both, text = "cisInvoiceDeduction:") =>
  assertRejects(() => Promise.resolve().then(() => run(invoice, supplier, customer)), Error, text);

test("1: a customer not in CIS is charged standard VAT", () => {
  const r = run({}, {}, notCis);
  assertEquals(
    [r.vatTreatment, r.labour, r.cisDeduction, r.vatLabel, r.vatValue, r.payable, r.notionalVat],
    ["standard", "£7,000.00", "£1,400.00", "VAT at 20%", "£2,000.00", "£10,600.00", null],
  );
});

test("2: silence from a registered customer is the reverse charge", () => {
  const r = run();
  assertEquals(
    [r.vatTreatment, r.vatValue, r.notionalVat, r.payable],
    ["reverse-charge", "Not charged: customer accounts to HMRC", 2000, "£8,600.00"],
  );
});

test("3 and 4: only written status leaves the reverse charge; blank is silence", () => {
  for (const notified of ["end-user", "intermediary"]) {
    const r = run({}, {}, { ...both, endUserNotified: notified });
    assertEquals([r.vatTreatment, r.payable], ["standard", "£10,600.00"]);
  }
  for (const notified of ["", null]) {
    assertEquals(run({}, {}, { ...both, endUserNotified: notified }).vatTreatment, "reverse-charge");
  }
});

test("5: the 5% disregard", () => {
  const treatment = (reverseChargeShare: number) => run({ reverseChargeShare }).vatTreatment;
  assertEquals([0.05, 0.051, 0].map(treatment), ["standard", "reverse-charge", "standard"]);
});

test("6: zero-rated work is never reverse-charged", () => {
  const r = run({ vatRate: 0 });
  assertEquals([r.vatTreatment, r.vatValue, r.notionalVat, r.payable], ["zero-rated", "£0.00", null, "£8,600.00"]);
});

test("7: reduced-rate work is reverse-charged on its own VAT", () => {
  const r = run({ vatRate: 0.05 });
  assertEquals([r.vatTreatment, r.notionalVat], ["reverse-charge", 500]);
});

test("8: unverified is deducted at 30% with a warning", () => {
  const r = run({}, { cisStatus: "unverified" });
  assertEquals([r.cisDeduction, r.cisLabel], ["£2,100.00", "CIS deduction at 30% (unverified)"]);
  assertEquals(r.warning.length > 0, true);
});

test("9: gross payment status deducts nothing", () => {
  const r = run({}, { cisStatus: "gross" }, notCis);
  assertEquals(
    [r.cisDeduction, r.cisLabel, r.payable, r.warning],
    ["£0.00", "CIS deduction (gross payment status)", "£12,000.00", ""],
  );
});

test("10 and 11: CIS floors and VAT rounds to the nearest penny", () => {
  const a = run({ grossExVat: 1234.59, materialsCost: 0, plantHire: undefined }, {}, notCis);
  assertEquals([a.cisDeduction, a.vatValue], ["£246.91", "£246.92"]);
  const b = run({ grossExVat: 1234.59, materialsCost: 0, plantHire: undefined, vatRate: 0.05 }, { cisStatus: "unverified" }, notCis);
  assertEquals([b.cisDeduction, b.vatValue], ["£370.37", "£61.73"]);
});

test("12 and 13: nothing, and costs equal to the gross, leave nothing to deduct from", () => {
  const z = run({ grossExVat: 0, materialsCost: 0, plantHire: 0 });
  assertEquals([z.gross, z.labour, z.cisDeduction, z.payable, z.notionalVat], ["£0.00", "£0.00", "£0.00", "£0.00", 0]);
  const e = run({ materialsCost: 8000, plantHire: 1500, consumables: 500 });
  assertEquals([e.labour, e.cisDeduction], ["£0.00", "£0.00"]);
});

test("14 and 15: costs over the gross, and negatives, are refused", async () => {
  await rejects({ materialsCost: 9000, plantHire: 1500 }, {}, both, "exceed the gross");
  await rejects({ grossExVat: -100 }, {}, both, "invoice.grossExVat");
  await rejects({ materialsCost: -1 }, {}, both, "invoice.materialsCost");
  await rejects({ plantHire: -1 }, {}, both, "invoice.plantHire");
  await rejects({ grossExVat: undefined }, {}, both, "invoice.grossExVat");
});

test("16: missing and zero plant hire and consumables are not printed", () => {
  for (const extra of [{ plantHire: undefined }, { plantHire: 0, consumables: 0 }, { plantHire: null }]) {
    const r = run(extra);
    assertEquals([r.plantHire, r.consumables], ["", ""]);
    assertEquals(r.labour, "£7,500.00");
  }
});

test("17: a supplier not registered for VAT has no VAT line and needs no customer flags", () => {
  const r = run({ vatRate: undefined }, { vatRegistered: false }, {});
  assertEquals([r.vatTreatment, r.vatLabel, r.vatValue, r.payable], ["not-registered", "", "", "£8,600.00"]);
});

test("18, 19 and 20: unclear data fails the document", async () => {
  await rejects({}, { vatRegistered: undefined }, both, "supplier.vatRegistered");
  await rejects({}, {}, { vatRegistered: true }, "customer.cisRegistered");
  await rejects({}, { cisStatus: "registered" }, both, "supplier.cisStatus");
  await rejects({ vatRate: 0.175 }, {}, both, "invoice.vatRate");
  await rejects({ reverseChargeShare: 1.5 }, {}, both, "invoice.reverseChargeShare");
  await rejects({}, {}, { ...both, endUserNotified: "contractor" }, "customer.endUserNotified");
});

// The node, built directly and published then resolved.
const previewData = {
  invoice: { ...base },
  supplier: { ...net, cisStatus: "unverified", verificationNumber: "V1" },
  customer: both,
};
const ARRANGEMENTS = ["table", "stacked"] as const;

const documentOf = (arrangement?: (typeof ARRANGEMENTS)[number]) =>
  template(
    <Document title="Invoice">
      <CisInvoiceDeduction arrangement={arrangement} />
    </Document>,
  );

type Found = { id?: string; text?: string; children?: DocumentNode[]; kind: string };

function walk(nodes: DocumentNode[]): Found[] {
  return nodes.flatMap((node) => {
    const n = node as Found;
    return [n, ...walk(n.children ?? [])];
  });
}

const ids = (nodes: DocumentNode[]) => walk(nodes).flatMap((n) => (n.id ? [n.id] : []));
const build = async (data: Obj, arrangement?: (typeof ARRANGEMENTS)[number]) =>
  (await buildDocument(documentOf(arrangement), data as never, { dynamicMode: "placeholder" })).nodes;

for (const arrangement of ARRANGEMENTS) {
  test(`${arrangement}: every row prints for the preview data`, async () => {
    const printed = ids(await build(previewData, arrangement)).filter((id) => id.startsWith("cis-deduction"));
    const rows = [
      "gross",
      "materials",
      "plant-hire",
      "labour",
      "vat",
      "cis",
      "payable",
    ].map((id) => `cis-deduction-${id}`);

    assertEquals(printed, [
      "cis-deduction",
      "cis-deduction-warning",
      ...(arrangement === "table" ? ["cis-deduction-lines"] : []),
      ...rows,
      "cis-deduction-verification",
    ]);
  });

  test(`${arrangement}: lines that do not apply are not printed`, async () => {
    const data = {
      invoice: { ...base, plantHire: 0 },
      supplier: { vatRegistered: false, cisStatus: "net" },
      customer: {},
    };
    const printed = ids(await build(data, arrangement));
    for (const id of ["warning", "plant-hire", "consumables", "vat", "verification"]) {
      assertEquals(printed.includes(`cis-deduction-${id}`), false);
    }
  });
}

test("publishing then resolving prints the payable and no warning", async () => {
  const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ cisInvoiceDeduction: mod.run }));
  const data = { invoice: base, supplier: net, customer: both };
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
    aiClient: { generateParagraph: () => "" },
  };
  const { nodes } = await resolveDocument(published, state, { derivers: registry });
  const found = walk(nodes);
  const payable = found.find((n) => n.id === "cis-deduction-payable");

  assertEquals(JSON.stringify(payable).includes("£8,600.00"), true);
  assertEquals(found.some((n) => n.id === "cis-deduction-warning"), false);
  assertEquals(dataRef("invoice") !== undefined, true);
});
