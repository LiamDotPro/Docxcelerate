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
  typeof import("../registry/components/derivers/bank-transfer-details.deriver.ts")
>("components/derivers/bank-transfer-details.deriver.ts");
const { BankTransferDetails } = await loadRegistryModule<
  typeof import("../registry/components/bank-transfer-details/bank-transfer-details.node.tsx")
>("components/bank-transfer-details/bank-transfer-details.node.tsx");
// Through a bundle, so each case also proves the function survives being shipped as source.
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ bankTransferDetails: mod.run }));

type Obj = Record<string, unknown>;
const payee: Obj = {
  accountName: "Northern Facilities Management Services (Holdings) Ltd",
  uk: { sortCode: "60-16-13", accountNumber: "31926819" },
  iban: "GB29 NWBK 6016 1331 9268 19",
  bic: "NWBKGB2L",
  us: { achRouting: "121000358", wireRouting: "026009593", accountNumber: "000123456789" },
};
const baseInvoice: Obj = { currency: "GBP", reference: "INV-1" };

interface Result {
  warning: string | null;
  method: string;
  accountName: string | null;
  sortCode: string | null;
  routingLabel: string | null;
  routing: string | null;
  wireRouting: string | null;
  accountNumber: string | null;
  iban: string | null;
  bic: string | null;
  reference: string | null;
}

function stateOf(data: Obj): RuntimeState {
  return {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
}

async function run(
  p: Obj = {},
  payer: Obj | undefined = { country: "GB" },
  invoice: Obj = {},
): Promise<Result> {
  const data = { payee: { ...payee, ...p }, payer, invoice: { ...baseInvoice, ...invoice } };
  return await registry.run(
    { name: "bankTransferDetails", output: "d", inputs: [dataRef("payee"), dataRef("payer"), dataRef("invoice")] } as never,
    stateOf(data),
  ) as Result;
}

const NONE = "We cannot print our bank details on this invoice. Please call us, on a number you already have for us, to arrange payment.";
const IBAN = "GB29 NWBK 6016 1331 9268 19";

test("UK domestic: sort code and account number only", async () => {
  const r = await run();
  assertEquals([r.sortCode, r.accountNumber, r.iban, r.bic, r.reference], ["60-16-13", "31926819", null, null, "INV-1"]);
  assertEquals(r.method.includes("Faster Payments"), true);
});

test("country and sort code are normalised", async () => {
  const base = await run();
  assertEquals(await run({}, { country: "uk" }), base);
  assertEquals(await run({}, { country: " gb " }), base);
  for (const sortCode of ["601613", "60 16 13"]) {
    assertEquals((await run({ uk: { sortCode, accountNumber: "31926819" } })).sortCode, "60-16-13");
  }
});

test("bad UK details fall to international", async () => {
  for (const uk of [{ sortCode: "60161", accountNumber: "31926819" }, { sortCode: "601613", accountNumber: "3192681" }]) {
    const r = await run({ uk });
    assertEquals([r.iban, r.bic, r.sortCode, r.method], [IBAN, "NWBKGB2L", null, "Pay by international (SWIFT) transfer, in GBP."]);
  }
});

test("SEPA", async () => {
  const sepa = { iban: "DE89370400440532013000", bic: undefined };
  const r = await run(sepa, { country: "IE" }, { currency: "EUR" });
  assertEquals([r.iban, r.bic, r.method], ["DE89 3704 0044 0532 0130 00", null, "Pay by SEPA credit transfer, in euros."]);
  assertEquals((await run({ ...sepa, bic: "cobadeffxxx" }, { country: "IE" }, { currency: "EUR" })).bic, "COBADEFFXXX");
  const invalid = await run({ ...sepa, bic: "NWBK" }, { country: "IE" }, { currency: "EUR" });
  assertEquals([invalid.iban, invalid.bic, invalid.method], ["DE89 3704 0044 0532 0130 00", null, "Pay by SEPA credit transfer, in euros."]);
});

test("the currency has to match the payer's rail", async () => {
  assertEquals((await run({}, { country: "FR" }, { currency: "GBP" })).method, "Pay by international (SWIFT) transfer, in GBP.");
  assertEquals((await run({}, { country: "GB" }, { currency: "eur" })).method, "Pay by international (SWIFT) transfer, in EUR.");
});

test("US routing numbers", async () => {
  const us = { country: "US" };
  const usd = { currency: "USD" };
  const mk = (extra: Obj) => ({ us: { achRouting: "121000358", accountNumber: "000123456789", ...extra } });
  const different = await run({}, us, usd);
  assertEquals(
    [different.routingLabel, different.routing, different.wireRouting, different.accountNumber, different.method],
    ["ACH routing number", "121000358", "026009593", "000123456789", "Pay by ACH or wire transfer, in US dollars."],
  );
  for (const extra of [{}, { wireRouting: "121000358" }]) {
    const r = await run(mk(extra), us, usd);
    assertEquals([r.routingLabel, r.wireRouting], ["Routing number (ACH and wire)", null]);
  }
  const bad = await run(mk({ wireRouting: "026009594" }), us, usd);
  assertEquals([bad.routingLabel, bad.wireRouting, bad.method], ["ACH routing number", null, "Pay by ACH transfer, in US dollars."]);
  const noAch = await run(mk({ achRouting: "121000359" }), us, usd);
  assertEquals([noAch.routing, noAch.iban, noAch.method], [null, IBAN, "Pay by international (SWIFT) transfer, in USD."]);
});

test("a failed IBAN is withheld, and with it every IBAN rail", async () => {
  for (const iban of ["GB29NWBK60161331926818", "GB29NWBK60161331928619"]) {
    const r = await run({ iban }, { country: "DE" }, { currency: "EUR", reference: "X".repeat(500) });
    assertEquals(r, {
      warning: null,
      method: NONE,
      accountName: null,
      sortCode: null,
      routingLabel: null,
      routing: null,
      wireRouting: null,
      accountNumber: null,
      iban: null,
      bic: null,
      reference: null,
    });
  }
});

test("unknown or missing country is international; a bad BIC is none", async () => {
  for (const payer of [{ country: "ZZ" }, {}, undefined]) {
    assertEquals((await run({}, payer, { currency: "EUR" })).method, "Pay by international (SWIFT) transfer, in EUR.");
  }
  assertEquals((await run({ bic: "NWBK" }, { country: "ZZ" })).method, NONE);
});

test("reference fitting", async () => {
  const ref = async (reference: string, customerReference?: string, payer = { country: "GB" }, currency = "GBP") =>
    (await run({}, payer, { reference, customerReference, currency })).reference;
  assertEquals(await ref("INV-2026-000123-AB"), "INV-2026-000123-AB");
  assertEquals(await ref("INV-2026-000123-ABC", "cust-9"), "CUST-9");
  assertEquals(await ref("INV-2026-11-004817/REV-2", "nfms-004817"), "NFMS-004817");
  assertEquals(await ref("  inv   12  "), "INV 12");
  assertEquals(await ref("  inv   12  ", undefined, { country: "FR" }, "EUR"), "inv 12");
  assertEquals(await ref("INV/12"), "INV/12");
  assertEquals(await ref("INV/12", undefined, { country: "FR" }, "EUR"), "INV/12");
  assertEquals(await ref("x".repeat(81), undefined, { country: "FR" }, "EUR"), "x".repeat(81));
  assertEquals((await run({}, { country: "GB" }, { currency: "GBP", reference: "INV-1", customerReference: "  " })).reference, "INV-1");
});

test("a reference that fits no rail throws, never truncated", async () => {
  await assertRejects(() => run({}, { country: "GB" }, { reference: "INV#12" }), Error, "does not fit a uk payment reference");
  await assertRejects(() => run({}, { country: "FR" }, { reference: "INV#12", currency: "EUR" }), Error, "does not fit a sepa");
  await assertRejects(
    () => run({}, { country: "US" }, { reference: "x".repeat(81), currency: "USD" }),
    Error,
    "does not fit a us payment reference (80 characters",
  );
  await assertRejects(() => run({}, { country: "GB" }, { reference: "INV#12", customerReference: "ALSO#BAD" }), Error, "no invoice.customerReference that does was given");
});

test("the account name is printed as given", async () => {
  assertEquals((await run({ accountName: "  ACME Widgets Ltd  " })).accountName, "ACME Widgets Ltd");
  assertEquals((await run({ accountName: "ACME  Widgets  Ltd" })).accountName, "ACME  Widgets  Ltd");
});

test("required fields throw their own message", async () => {
  for (const accountName of [undefined, "  "]) {
    await assertRejects(() => run({ accountName }), Error, "payee.accountName is required");
  }
  for (const reference of [undefined, "  "]) {
    await assertRejects(() => run({}, undefined, { reference }), Error, "invoice.reference is required");
  }
  for (const currency of ["GB", undefined]) {
    await assertRejects(() => run({}, undefined, { currency }), Error, "invoice.currency must be a three-letter ISO 4217 code");
  }
});

test("the changed-details warning", async () => {
  const w = async (detailsChangedOn: string | undefined, payer = { country: "GB" }, p: Obj = {}) =>
    (await run({ detailsChangedOn, ...p }, payer, payer.country === "US" ? { currency: "USD" } : {})).warning;
  assertEquals((await w("2026-11-03"))?.startsWith("Our bank details changed on 3 November 2026."), true);
  assertEquals((await w("2026-11-03", { country: "US" }))?.startsWith("Our bank details changed on November 3, 2026."), true);
  assertEquals((await w("last week"))?.startsWith("Our bank details have changed since our last invoice."), true);
  assertEquals((await w("2026-13-45"))?.startsWith("Our bank details have changed since"), true);
  assertEquals(await w(""), null);
  assertEquals(await w(undefined), null);
  assertEquals((await run({ detailsChangedOn: "2026-11-03", iban: "bad" }, { country: "ZZ" })).warning, null);
});

// ---------------------------------------------------------------------------
// Rendering and publishing
// ---------------------------------------------------------------------------

const ARRANGEMENTS = ["stacked", "table", "split"] as const;
const nodeData = {
  payee: { ...payee, detailsChangedOn: "2026-11-03" },
  payer: { country: "GB" },
  invoice: { currency: "GBP", reference: "INV-2026-11-004817/REV-2", customerReference: "NFMS-004817" },
};

function documentOf(arrangement?: (typeof ARRANGEMENTS)[number]) {
  return template(
    <Document title="Invoice">
      <BankTransferDetails arrangement={arrangement} />
    </Document>,
  );
}

type Walk = { id?: string; text?: string; children?: DocumentNode[]; when?: unknown };
function walk(nodes: DocumentNode[]): Walk[] {
  return nodes.flatMap((n) => [n as Walk, ...walk((n as Walk).children ?? [])]);
}
const idsOf = (nodes: DocumentNode[]) => walk(nodes).map((n) => n.id).filter((id) => id?.startsWith("bank-transfer-"));
const textsOf = (nodes: DocumentNode[]) => walk(nodes).flatMap((n) => (n.text ? [n.text] : []));

for (const arrangement of ARRANGEMENTS) {
  test(`${arrangement}: warning, then method, then details; no IBAN on the UK rail`, async () => {
    const built = await buildDocument(documentOf(arrangement), nodeData, { dynamicMode: "placeholder" });
    const ids = idsOf(built.nodes);
    assertEquals(ids.indexOf("bank-transfer-warning") < ids.indexOf("bank-transfer-method"), true);
    assertEquals(ids.indexOf("bank-transfer-method") < ids.indexOf("bank-transfer-account-name"), true);
    assertEquals(ids.includes("bank-transfer-iban"), false);
    assertEquals(ids.includes("bank-transfer-sort-code"), true);
    assertEquals(textsOf(built.nodes).some((t) => t.includes("NFMS-004817")), true);
  });
}

test("arrangements place the same lines differently", async () => {
  const lines = async (a: (typeof ARRANGEMENTS)[number]) =>
    textsOf((await buildDocument(documentOf(a), nodeData, { dynamicMode: "placeholder" })).nodes);
  const table = await lines("table");
  assertEquals(table.slice(-4), ["Account number", "31926819", "Payment reference", "NFMS-004817"]);
  assertEquals((await lines("split")).slice(-2), ["Account number: 31926819", "Quote this reference: NFMS-004817"]);
});

test("published, the rail is decided by the engine, not at publish time", async () => {
  const published = await buildDocument(documentOf(), createPublishData() as typeof nodeData, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const stored = walk(published.nodes);
  for (const id of ["bank-transfer-iban", "bank-transfer-sort-code", "bank-transfer-routing", "bank-transfer-warning"]) {
    assertEquals(stored.find((n) => n.id === id)?.when !== undefined, true);
  }

  const resolve = async (data: Obj) => idsOf((await resolveDocument(published, stateOf(data), { derivers: registry })).nodes);
  const uk = await resolve(nodeData);
  assertEquals(uk.includes("bank-transfer-sort-code") && !uk.includes("bank-transfer-routing") && !uk.includes("bank-transfer-iban"), true);
  const us = await resolve({ ...nodeData, payer: { country: "US" }, invoice: { currency: "USD", reference: "INV-1" } });
  assertEquals(us.includes("bank-transfer-routing") && us.includes("bank-transfer-wire-routing"), true);
  assertEquals(us.includes("bank-transfer-sort-code") || us.includes("bank-transfer-iban"), false);
});

const publishCases: Record<string, Obj> = {
  uk: nodeData,
  us: { ...nodeData, payer: { country: "US" }, invoice: { currency: "USD", reference: "INV-1" } },
  sepa: {
    ...nodeData,
    payee: { ...nodeData.payee, iban: "DE89370400440532013000" },
    payer: { country: "IE" },
    invoice: { currency: "EUR", reference: "INV-1" },
  },
  // ZZ with the IBAN and BIC left in would be the international rail; take them out.
  none: { ...nodeData, payee: { ...nodeData.payee, iban: undefined, bic: undefined }, payer: { country: "ZZ" } },
};

for (const arrangement of ARRANGEMENTS) {
  for (const [rail, data] of Object.entries(publishCases)) {
    test(`${arrangement}: publishing then resolving matches building directly (${rail})`, async () => {
      const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof nodeData, {
        branchMode: "publish",
        deriverMode: "preserve",
        dynamicMode: "resolve",
        aiClient: { generateParagraph: () => "" },
      });
      const resolved = await resolveDocument(published, stateOf(data), { derivers: registry });
      const direct = await buildDocument(documentOf(arrangement), data as typeof nodeData, { dynamicMode: "placeholder" });
      // Generated cell ids count differently once branches are published, so compare what is printed.
      assertEquals(textsOf(resolved.nodes), textsOf(direct.nodes));
      assertEquals(idsOf(resolved.nodes), idsOf(direct.nodes));
    });
  }
}

for (const arrangement of ["table", "split"] as const) {
  test(`${arrangement}: the none rail prints no empty table`, async () => {
    const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof nodeData, {
      branchMode: "publish",
      deriverMode: "preserve",
      dynamicMode: "resolve",
      aiClient: { generateParagraph: () => "" },
    });
    const resolved = await resolveDocument(published, stateOf(publishCases.none), { derivers: registry });
    assertEquals(idsOf(resolved.nodes).includes("bank-transfer-details"), false);
  });
}
