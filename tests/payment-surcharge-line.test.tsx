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
  typeof import("../registry/components/derivers/payment-surcharge-line.deriver.ts")
>("components/derivers/payment-surcharge-line.deriver.ts");
const { PaymentSurchargeLine } = await loadRegistryModule<
  typeof import("../registry/components/payment-surcharge-line/payment-surcharge-line.node.tsx")
>("components/payment-surcharge-line/payment-surcharge-line.node.tsx");

type Obj = Record<string, unknown>;
type Lines = Record<string, string | null>;

const gbBands = [
  { ratePercent: 20, net: 1184250.0, vat: 236850.0 },
  { ratePercent: 5, net: 96400.0, vat: 4820.0 },
  { ratePercent: 0, net: 48975.5, vat: 0 },
];
const preview: Obj = {
  jurisdiction: "GB",
  customer: "business",
  method: { kind: "commercial-card", network: "mastercard" },
  percent: 2.5,
  costOfAcceptance: 1.95,
  currency: "GBP",
  bands: gbBands,
};
const ny: Obj = {
  jurisdiction: "US",
  usState: "NY",
  method: { kind: "credit-card", network: "visa" },
  percent: 3.5,
  costOfAcceptance: 2.9,
  currency: "USD",
  bands: [{ ratePercent: 8.875, net: 1000, vat: 88.75 }],
};

const run = async (s: Obj): Promise<Lines> => await mod.run([s], {} as RuntimeState) as Lines;
const NULLS = {
  surchargeLabel: null,
  surchargeAmount: null,
  taxLabel: null,
  taxAmount: null,
  taxNote: null,
  totalLabel: null,
  totalAmount: null,
  disclosure: null,
  warning: null,
};
const refused = (why: string) => ({
  ...NULLS,
  warning: `Surcharge not applied: ${why} Remove the surcharge or correct the data before this document is issued.`,
});
const gbConsumer = refused(
  "surcharging a consumer for any payment method is banned in the UK " +
    "(Consumer Rights (Payment Surcharges) Regulations 2012, as amended from 13 January 2018).",
);
const cards = refused(
  "consumer cards (credit, debit and prepaid) may not be surcharged in the UK or EU, " +
    "whoever is paying (PSD2 Art. 62(4) and the UK equivalent).",
);

test("1: the preview data caps to 1.95% and spreads across three bands", async () => {
  const r = await run(preview);
  assertEquals(r.surchargeLabel, "Surcharge for payment by commercial card (1.95%)");
  assertEquals(r.surchargeAmount, "£25,927.69");
  assertEquals(r.taxLabel, "VAT on surcharge");
  assertEquals(r.taxAmount, "£4,712.56");
  assertEquals(r.totalAmount, "£1,601,935.75");
  assertEquals(
    r.taxNote,
    "The surcharge is split across the invoice's VAT rates in proportion to their net: " +
      "£23,092.87 at 20%, £1,879.80 at 5%, £955.02 at 0%.",
  );
  assertStringIncludes(r.disclosure ?? "", "1.95%");
  assertEquals(r.warning, null);
});

test("2: the remainder unit goes to the band with the larger remainder", async () => {
  const r = await run({
    ...preview,
    percent: 2,
    costOfAcceptance: 2.5,
    bands: [{ ratePercent: 20, net: 100, vat: 20 }, { ratePercent: 0, net: 33.33, vat: 0 }],
  });
  assertEquals(r.surchargeAmount, "£2.67");
  assertEquals(r.taxAmount, "£0.40");
  assertEquals(r.totalAmount, "£156.40");
  assertStringIncludes(r.taxNote ?? "", "£2.00 at 20%, £0.67 at 0%");
});

test("3: EU VAT rounds half up, and one band gets no note", async () => {
  const r = await run({
    jurisdiction: "EU",
    method: { kind: "commercial-card" },
    percent: 1.5,
    costOfAcceptance: 2,
    currency: "EUR",
    bands: [{ ratePercent: 21, net: 500, vat: 105 }],
  });
  assertEquals(r.surchargeAmount, "€7.50");
  assertEquals(r.taxAmount, "€1.58");
  assertEquals(r.totalAmount, "€614.08");
  assertEquals(r.taxNote, null);
});

test("4: US credit card is capped at cost, taxed, and gets the US disclosure", async () => {
  const r = await run(ny);
  assertEquals(r.surchargeLabel, "Surcharge for payment by credit card (2.9%)");
  assertEquals(r.surchargeAmount, "$29.00");
  assertEquals(r.taxLabel, "Sales tax on surcharge");
  assertEquals(r.taxAmount, "$2.57");
  assertEquals(r.totalAmount, "$1,120.32");
  assertEquals(
    r.disclosure,
    "We impose a surcharge of 2.9% on credit card payments, which is not greater than our cost of acceptance. " +
      "No surcharge applies to debit card payments.",
  );
});

test("5-7: the network cap, and the lowest cap for a missing or unknown network", async () => {
  const method = (network?: string) => ({ ...ny, costOfAcceptance: 5, method: { kind: "credit-card", network } });
  assertEquals((await run(method("visa"))).surchargeAmount, "$30.00");
  assertEquals((await run(method())).surchargeAmount, "$30.00");
  assertEquals((await run(method(" Amex "))).surchargeAmount, "$30.00");
  assertEquals((await run(method("mastercard"))).surchargeAmount, "$35.00");
  assertEquals((await run(method(" MasterCard "))).surchargeAmount, "$35.00");
});

test("8: Colorado caps at 2%", async () => {
  assertEquals((await run({ ...ny, usState: "CO", percent: 3, costOfAcceptance: 3.5 })).surchargeAmount, "$20.00");
});

test("9: a state that does not tax the surcharge", async () => {
  const r = await run({ ...ny, usTaxable: false });
  assertEquals([r.taxLabel, r.taxAmount, r.taxNote], [null, null, null]);
  assertEquals(r.totalAmount, "$1,117.75");
});

test("10-11: a UK consumer, or a missing customer, is never surcharged", async () => {
  assertEquals(await run({ ...preview, customer: "consumer" }), gbConsumer);
  assertEquals(await run({ ...preview, customer: undefined }), gbConsumer);
  assertEquals(await run({ ...preview, customer: "consumer", method: { kind: "bank-transfer" } }), gbConsumer);
});

test("12: consumer cards are banned in the UK and EU, whoever pays", async () => {
  assertEquals(await run({ ...preview, method: { kind: "debit-card" } }), cards);
  assertEquals(await run({ ...preview, jurisdiction: "EU", method: { kind: "credit-card" } }), cards);
  assertEquals(await run({ ...preview, jurisdiction: "EU", method: { kind: "prepaid-card" } }), cards);
});

test("13: no SEPA transfer in the EU", async () => {
  assertEquals(
    await run({ ...preview, jurisdiction: "EU", method: { kind: "sepa" } }),
    refused("SEPA credit transfers and direct debits may not be surcharged (PSD2 Art. 62(4))."),
  );
});

test("14-15: in the US only credit and commercial cards", async () => {
  const debit = refused("debit and prepaid cards may never be surcharged in the US.");
  assertEquals(await run({ ...ny, method: { kind: "debit-card" } }), debit);
  assertEquals(await run({ ...ny, method: { kind: "prepaid-card" } }), debit);
  const other = refused("only credit cards, commercial cards included, may be surcharged in the US.");
  assertEquals(await run({ ...ny, method: { kind: "bank-transfer" } }), other);
  assertEquals(await run({ ...ny, method: { kind: "sepa" } }), other);
  assertEquals((await run({ ...ny, method: { kind: "commercial-card" } })).warning, null);
});

test("16: a US sale without a state", async () => {
  const missing = refused("the US state decides whether a surcharge is allowed, and no usState was given.");
  assertEquals(await run({ ...ny, usState: undefined }), missing);
  assertEquals(await run({ ...ny, usState: "  " }), missing);
});

test("17: banned states, trimmed and case-insensitive", async () => {
  assertEquals(await run({ ...ny, usState: "ct" }), refused("surcharges are banned in CT."));
  assertEquals(await run({ ...ny, usState: " MA " }), refused("surcharges are banned in MA."));
});

test("18: no cost of acceptance, no surcharge", async () => {
  assertEquals(
    await run({ ...preview, costOfAcceptance: undefined }),
    refused("every rule that allows a surcharge caps it at the cost of acceptance, and no costOfAcceptance was given."),
  );
});

test("19: SEPA in the UK is a bank transfer", async () => {
  const r = await run({ ...preview, method: { kind: "sepa" }, costOfAcceptance: 1 });
  assertEquals(r.surchargeLabel, "Surcharge for payment by bank transfer (1%)");
  assertStringIncludes(r.disclosure ?? "", "payment by bank transfer.");
});

test("20: nothing requested prints nothing, even for a banned combination", async () => {
  assertEquals(await run({ ...preview, percent: undefined, customer: "consumer" }), NULLS);
  assertEquals(await run({ ...preview, percent: 0, customer: "consumer" }), NULLS);
});

test("21: no base prints nothing", async () => {
  assertEquals(await run({ ...preview, bands: [] }), NULLS);
  assertEquals(await run({ ...preview, bands: [{ ratePercent: 20, net: 0, vat: 0 }] }), NULLS);
});

test("22: a surcharge that rounds to nothing prints nothing", async () => {
  assertEquals(await run({ ...preview, percent: 1, bands: [{ ratePercent: 0, net: 0.1 }] }), NULLS);
});

test("23: all bands at 0% carry no tax and no note", async () => {
  const r = await run({
    ...preview,
    bands: [{ ratePercent: 0, net: 100, vat: 0 }, { ratePercent: 0, net: 100, vat: 0 }],
  });
  assertEquals([r.surchargeAmount, r.taxAmount, r.taxNote, r.totalAmount], ["£3.90", null, null, "£203.90"]);
});

test("24: shares sum to the surcharge, extra units to the earliest bands", async () => {
  const r = await run({
    ...preview,
    percent: 1.05,
    costOfAcceptance: 2,
    bands: Array.from({ length: 6 }, () => ({ ratePercent: 20, net: 10, vat: 2 })),
  });
  assertEquals(r.surchargeAmount, "£0.63");
  assertStringIncludes(
    r.taxNote ?? "",
    "£0.11 at 20%, £0.11 at 20%, £0.11 at 20%, £0.10 at 20%, £0.10 at 20%, £0.10 at 20%.",
  );
});

test("24b: a base past 2^53 / S still splits exactly", async () => {
  const r = await run({
    ...preview,
    percent: 1.95,
    costOfAcceptance: 2,
    bands: [
      { ratePercent: 20, net: 13624710.26, vat: 0 },
      { ratePercent: 5, net: 16919167.38, vat: 0 },
      { ratePercent: 0, net: 5505409.69, vat: 0 },
    ],
  });
  assertEquals(r.surchargeAmount, "£702,961.10");
  assertStringIncludes(r.taxNote ?? "", "£265,681.85 at 20%, £329,923.76 at 5%, £107,355.49 at 0%.");
  // 53,136.37 + 16,496.19 (exact shares at 20% and 5%)
  assertEquals(r.taxAmount, "£69,632.56");
});

test("25: a currency with no minor units", async () => {
  const r = await run({
    ...preview,
    currency: "JPY",
    percent: 2,
    costOfAcceptance: 3,
    bands: [{ ratePercent: 10, net: 123457, vat: 12346 }],
  });
  assertEquals([r.surchargeAmount, r.taxAmount], ["JP¥2,469", "JP¥247"]);
});

test("26: bad data throws, naming the field", async () => {
  const bad = (overrides: Obj, field: string) => assertRejects(() => run({ ...preview, ...overrides }), Error, field);
  await bad({ percent: -1 }, "surcharge.percent");
  await bad({ costOfAcceptance: NaN }, "surcharge.costOfAcceptance");
  await bad({ bands: [{ ratePercent: 20, net: -5 }] }, "bands.net");
  await bad({ bands: [{ ratePercent: -1, net: 5 }] }, "bands.ratePercent");
  await bad({ bands: [{ ratePercent: 20, net: 5, vat: -1 }] }, "bands.vat");
  await bad({ bands: "x" }, "surcharge.bands");
  await bad({ jurisdiction: "UK" }, "surcharge.jurisdiction");
  await bad({ method: { kind: "cash" } }, "surcharge.method.kind");
  await bad({ customer: "trader" }, "surcharge.customer");
  await bad({ currency: undefined }, "surcharge.currency");
  await bad({ currency: "xx" }, "surcharge.currency");
});

// The node, built and published.
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ paymentSurchargeLine: mod.run }));

function documentOf(arrangement?: "table" | "stacked") {
  return template(
    <Document title="Invoice">
      <PaymentSurchargeLine arrangement={arrangement} />
    </Document>,
  );
}

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[]; rows?: DocumentNode[]; cells?: DocumentNode[] };
    return [...(own.id ? [own.id] : []), ...ids([...(own.children ?? []), ...(own.rows ?? []), ...(own.cells ?? [])])];
  });
}

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { text?: string; children?: DocumentNode[]; rows?: DocumentNode[]; cells?: DocumentNode[] };
    return [
      ...(own.text ? [own.text] : []),
      ...texts([...(own.children ?? []), ...(own.rows ?? []), ...(own.cells ?? [])]),
    ];
  });
}

const build = (arrangement: "table" | "stacked" | undefined, surcharge: Obj) =>
  buildDocument(documentOf(arrangement), { surcharge }, { dynamicMode: "placeholder" });

test("each arrangement carries the same lines, and they differ", async () => {
  const shared = ["surcharge-amount", "surcharge-tax", "surcharge-total", "surcharge-tax-note", "surcharge-disclosure"];
  const table = ids((await build("table", preview)).nodes);
  const stacked = ids((await build("stacked", preview)).nodes);
  for (const id of shared) {
    assertEquals(table.includes(id), true);
    assertEquals(stacked.includes(id), true);
  }
  assertEquals(table.includes("surcharge-table"), true);
  assertEquals(stacked.includes("surcharge-table"), false);
  assertEquals(table.includes("surcharge-warning"), false);
  assertEquals(JSON.stringify(table) !== JSON.stringify(stacked), true);
  assertEquals((await build(undefined, preview)).nodes, (await build("table", preview)).nodes);
});

test("a banned case prints only the warning", async () => {
  for (const arrangement of ["table", "stacked"] as const) {
    assertEquals(ids((await build(arrangement, { ...preview, customer: "consumer" })).nodes), ["surcharge-warning"]);
  }
});

async function resolved(arrangement: "table" | "stacked", surcharge: Obj) {
  const data = { surcharge };
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
  return { published, resolved: await resolveDocument(published, state, { derivers: registry }) };
}

for (const arrangement of ["table", "stacked"] as const) {
  test(`${arrangement}: published, bundled and resolved, it matches building directly`, async () => {
    const { published, resolved: us } = await resolved(arrangement, ny);
    // The deriver travels and nothing was settled at build time.
    assertStringIncludes(JSON.stringify(published.nodes), "paymentSurchargeLine");
    assertEquals(ids(published.nodes).includes("surcharge-warning"), true);
    assertEquals(ids(published.nodes).includes("surcharge-amount"), true);
    const direct = await build(arrangement, ny);
    assertEquals(texts(us.nodes), texts(direct.nodes));
    assertStringIncludes(texts(us.nodes).join("\n"), "$29.00");
    assertStringIncludes(texts(us.nodes).join("\n"), "We impose a surcharge of 2.9%");

    const { resolved: banned } = await resolved(arrangement, { ...preview, customer: "consumer" });
    assertEquals(ids(banned.nodes), ["surcharge-warning"]);
    assertEquals(texts(banned.nodes), [(gbConsumer.warning as string)]);
  });
}
