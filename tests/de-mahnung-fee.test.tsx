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
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";
import { COMPONENTS } from "../src/registry/catalog.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/de-mahnung-fee.deriver.ts")
>("components/derivers/de-mahnung-fee.deriver.ts");
const { DeMahnungFee } = await loadRegistryModule<
  typeof import("../registry/components/de-mahnung-fee/de-mahnung-fee.node.tsx")
>("components/de-mahnung-fee/de-mahnung-fee.node.tsx");
// Through a bundle, so the test proves `run` survives being shipped as source text.
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ deMahnungFee: mod.run }));

type Obj = Record<string, unknown>;
interface Row {
  number: string;
  heading: string;
  verzugLine: string;
  feeLine: string;
  pauschaleLine: string;
  route: string;
  verzugSince: string;
  feeNowMinor: number;
  costsMinor: number;
}
interface Result {
  business: boolean;
  proofLine: string;
  invoices: Row[];
}

// Intl puts a no-break space before the euro sign; the tests read it as a space.
const plain = (text: string) => text.replaceAll(" ", " ");

async function call(mahnung: unknown): Promise<Result> {
  const data = { mahnung };
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data as Obj),
    aiClient: { generateParagraph: () => "" },
  };
  return await registry.run(
    { name: "deMahnungFee", output: "o", inputs: [dataRef("mahnung")] } as never,
    state,
  ) as Result;
}

const invoice = (extra: Obj = {}): Obj => ({ number: "R1", amount: 100, dueDate: "2026-10-14", ...extra });
const letter = (inv: Obj = {}, extra: Obj = {}): Obj => ({
  debtorType: "business",
  date: "2026-11-19",
  invoices: [invoice(inv)],
  ...extra,
});
/** One invoice, run, returned as its row with the dashes of the German text normalised. */
async function one(inv: Obj = {}, extra: Obj = {}): Promise<Row & { proofLine: string }> {
  const result = await call(letter(inv, extra));
  const row = result.invoices[0];
  return {
    ...row,
    verzugLine: plain(row.verzugLine),
    feeLine: plain(row.feeLine),
    pauschaleLine: plain(row.pauschaleLine),
    proofLine: result.proofLine,
  };
}

test("1. a contract date on a weekday starts Verzug the next day", async () => {
  const r = await one({ contractPaymentDate: "2026-10-14" });
  assertEquals([r.route, r.verzugSince], ["contract", "2026-10-15"]);
  assertEquals(r.verzugLine.includes("§ 193"), false);
  assertStringIncludes(r.verzugLine, "(§ 286 Abs. 2 Nr. 1 BGB)");
});

test("2. a contract date on a Saturday runs to Monday", async () => {
  const r = await one({ dueDate: "2026-10-17", contractPaymentDate: "2026-10-17" });
  assertEquals(r.verzugSince, "2026-10-20");
  assertStringIncludes(r.verzugLine, "lief er nach § 193 BGB bis 19.10.2026");
});

test("3. Easter holidays are computed", async () => {
  const r = await one(
    { dueDate: "2026-04-03", contractPaymentDate: "2026-04-03" },
    { date: "2026-05-01" },
  );
  assertEquals(r.verzugSince, "2026-04-08");
});

test("4. an event-based date says Nr. 2", async () => {
  const r = await one({ contractPaymentDate: "2026-10-14", contractPaymentFromEvent: true });
  assertEquals(r.route, "event");
  assertStringIncludes(r.verzugLine, "§ 286 Abs. 2 Nr. 2");
});

test("5. 30 days across a holiday", async () => {
  const r = await one({ dueDate: "2026-09-01", invoiceReceivedDate: "2026-09-03" });
  assertEquals([r.route, r.verzugSince], ["invoice-30", "2026-10-06"]);
  assertStringIncludes(r.verzugLine, "nach § 193 BGB auf 05.10.2026");
});

test("6. a Land holiday decides the fee", async () => {
  const inv = { dueDate: "2026-10-19", invoiceReceivedDate: "2026-10-19" };
  const free = await one(inv, { fee: 6.5, holidays: ["2026-11-18"] });
  assertEquals([free.route, free.feeNowMinor], ["this-mahnung", 0]);
  assertStringIncludes(free.feeLine, "keine Mahngebühr, weil erst dieses Schreiben");
  const charged = await one(inv, { fee: 6.5 });
  assertEquals([charged.route, charged.verzugSince], ["invoice-30", "2026-11-19"]);
  assertStringIncludes(charged.feeLine, "6,50 €");
});

const late = { dueDate: "2026-09-01", invoiceReceivedDate: "2026-09-03" };

test("7. a consumer without the warning has no 30-day route", async () => {
  const r = await one(late, { debtorType: "consumer", fee: 2.5 });
  assertEquals([r.route, r.feeNowMinor, r.proofLine, r.pauschaleLine, r.costsMinor], ["this-mahnung", 0, "", "", 0]);
});

test("8. a consumer with the warning is charged and may prove lower damage", async () => {
  const r = await one({ ...late, verzugWarningOnInvoice: true }, { debtorType: "consumer", fee: 2.5 });
  assertEquals([r.route, r.verzugSince, r.feeNowMinor, r.pauschaleLine, r.costsMinor], ["invoice-30", "2026-10-06", 250, "", 250]);
  assertStringIncludes(r.feeLine, "2,50 €");
  assertEquals(r.proofLine !== "", true);
});

test("9. the consumer ceiling", async () => {
  await assertRejects(() => call(letter({}, { debtorType: "consumer", fee: 2.51 })), Error, "mahnung.fee");
  await call(letter({}, { fee: 2.51 }));
});

test("10. goods receipt stands in for a missing invoice date, for a business only", async () => {
  const inv = { dueDate: "2026-09-01", goodsReceivedDate: "2026-09-15" };
  const b = await one(inv);
  assertEquals([b.route, b.verzugSince], ["goods-30", "2026-10-16"]);
  assertStringIncludes(b.verzugLine, "Empfang der Leistung am 15.09.2026 (§ 286 Abs. 3 Satz 2 BGB)");
  assertEquals((await one(inv, { debtorType: "consumer" })).route, "this-mahnung");
});

test("11. a Mahnung on the due date has no effect", async () => {
  const r = await one({
    dueDate: "2026-09-01",
    mahnungen: [{ date: "2026-09-01" }, { date: "2026-09-10", feeCharged: 0 }],
  });
  assertEquals([r.route, r.verzugSince], ["earlier-mahnung", "2026-09-10"]);
});

const earlierFees = (...fees: number[]) => ({
  dueDate: "2026-09-01",
  mahnungen: fees.map((feeCharged, i) => ({ date: `2026-09-0${i + 2}`, feeCharged })),
});

test("12. business crediting, under 40", async () => {
  const r = await one(earlierFees(), { fee: 6.5 });
  // No earlier Mahnung: this one is the first, so it is free; use a later one for the fee.
  assertEquals(r.route, "this-mahnung");
  const c = await one(earlierFees(0), { fee: 6.5 });
  assertEquals([c.feeNowMinor, c.costsMinor], [0, 4000]);
  assertStringIncludes(c.feeLine, "angerechnet und nicht zusätzlich berechnet");
  assertStringIncludes(c.pauschaleLine, "6,50 €");
});

test("13. business crediting, exactly 40", async () => {
  const r = await one(earlierFees(33.5), { fee: 6.5 });
  assertEquals([r.feeNowMinor, r.costsMinor], [0, 4000]);
  assertEquals(r.pauschaleLine.includes("übersteigende"), false);
});

test("14. business crediting, crossing 40", async () => {
  const r = await one(earlierFees(39), { fee: 6.5 });
  assertEquals([r.feeNowMinor, r.costsMinor], [550, 4550]);
  assertStringIncludes(r.feeLine, "Davon berechnen wir 5,50 € zusätzlich");
  assertStringIncludes(r.pauschaleLine, "45,50 €");
  assertStringIncludes(r.pauschaleLine, "5,50 €");
});

test("15. business crediting, already over 40", async () => {
  const r = await one(earlierFees(45), { fee: 6.5 });
  assertEquals([r.feeNowMinor, r.costsMinor], [650, 5150]);
  assertStringIncludes(r.feeLine, "bereits übersteigen");
});

test("16. a zero fee and a missing fee are the same", async () => {
  for (const debtorType of ["business", "consumer"]) {
    const inv = { ...late, verzugWarningOnInvoice: true };
    const zero = await call(letter(inv, { debtorType, fee: 0 }));
    const missing = await call(letter(inv, { debtorType }));
    assertEquals(zero, missing);
    assertEquals(plain(zero.invoices[0].feeLine), "Für dieses Schreiben berechnen wir keine Mahngebühr.");
    assertEquals(zero.proofLine, "");
  }
});

const previewData = COMPONENTS.find((c) => c.id === "de-mahnung-fee")!.previewData as { mahnung: Obj };

test("17. many invoices: print order, one fee carrier, the numbers of the preview", async () => {
  const { invoices } = await call(previewData.mahnung);
  const [a, b, c] = invoices;
  assertEquals(invoices.map((i) => i.number.slice(0, 15)), ["RE-2026-004817/", "RE-2026-005102-", "RE-2026-006233-"]);
  assertEquals([a.route, a.verzugSince, a.feeNowMinor, a.costsMinor], ["contract", "2026-09-01", 550, 4550]);
  assertStringIncludes(plain(a.pauschaleLine), "45,50 €");
  assertStringIncludes(plain(a.pauschaleLine), "5,50 €");
  assertEquals([b.route, b.verzugSince, b.costsMinor], ["invoice-30", "2026-10-06", 4000]);
  assertStringIncludes(b.feeLine, "bei Rechnung RE-2026-004817/TEILLIEFERUNG-3-RAHMENVERTRAG-NORD berechnet");
  assertStringIncludes(plain(b.pauschaleLine), "Verzugspauschale nach § 288 Abs. 5 BGB: 40,00 €.");
  assertEquals(c.route, "this-mahnung");
  assertStringIncludes(c.pauschaleLine, "Mit Eintritt des Verzugs entsteht");
  assertStringIncludes(plain(a.heading), "1.284.750,00 €, fällig am 31.08.2026");
});

test("18. what throws", async () => {
  const bad = (mahnung: unknown, text?: string) => assertRejects(() => call(mahnung), Error, text);
  await bad(letter({}, { debtorType: "Verbraucher" }), "mahnung.debtorType");
  await bad(letter({}, { invoices: [] }), "mahnung.invoices");
  await bad({ debtorType: "business", date: "2026-11-19" }, "mahnung.invoices");
  await bad(letter({ amount: 0 }), "mahnung.invoices[0].amount");
  await bad(letter({ amount: -5 }), "amount");
  await bad(letter({ amount: NaN }), "amount");
  await bad(letter({ number: "  " }), "number");
  await bad(letter({ dueDate: "19.11.2026" }), "mahnung.invoices[0].dueDate must be an ISO 8601 date");
  await bad(letter({ dueDate: "2026-13-01" }), "mahnung.invoices[0].dueDate must be an ISO 8601 date");
  await bad(letter({ contractPaymentDate: "2026-10-01" }), "contractPaymentDate");
  await bad(letter({ mahnungen: [{ date: "2026-11-19" }] }), "mahnungen[0].date");
  await bad(letter({ mahnungen: [{ date: "2026-10-20", feeCharged: -1 }] }), "feeCharged");
  await bad(letter({}, { fee: -1 }), "mahnung.fee");
  await bad(letter({}, { holidays: ["tomorrow"] }), "mahnung.holidays[0]");
  await bad(letter({}, { date: "2026-10-14" }), "mahnung.date");
  // Due Saturday 14.11 is due through Monday 16.11, so a letter on the 16th is too early.
  await bad(letter({ dueDate: "2026-11-14" }, { date: "2026-11-16" }), "16.11.2026");
});

test("19. money is rounded to cents once", async () => {
  const r = await one(late, { debtorType: "consumer", fee: 2.005 });
  // Not chargeable without the warning, so ask with it.
  const w = await one({ ...late, verzugWarningOnInvoice: true }, { debtorType: "consumer", fee: 2.005 });
  assertEquals(r.feeNowMinor, 0);
  assertEquals([w.feeNowMinor, plain(w.feeLine).includes("2,01 €")], [201, true]);
  const sum = await one(
    { dueDate: "2026-09-01", mahnungen: [0.1, 0.1, 0.1].map((feeCharged, i) => ({ date: `2026-09-0${i + 2}`, feeCharged })) },
    { debtorType: "consumer" },
  );
  assertEquals(sum.costsMinor, 30);
});

// ---------------------------------------------------------------------------
// The node, published and resolved by an engine that only has the bundle.
// ---------------------------------------------------------------------------

const ARRANGEMENTS = ["stacked", "table"] as const;

const documentOf = (arrangement: (typeof ARRANGEMENTS)[number]) =>
  template(
    <Document title="Mahnung">
      <DeMahnungFee arrangement={arrangement} />
    </Document>,
  );

function walk(nodes: DocumentNode[]): Array<{ id?: string; text?: string }> {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; text?: string; children?: DocumentNode[]; rows?: DocumentNode[]; cells?: DocumentNode[] };
    return [own, ...walk(own.children ?? []), ...walk(own.rows ?? []), ...walk(own.cells ?? [])];
  });
}

async function resolved(arrangement: (typeof ARRANGEMENTS)[number], mahnung: Obj) {
  const data = { mahnung };
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
  return (await resolveDocument(published, state, { derivers: registry })).nodes;
}

const consumerRequest = {
  debtorType: "consumer",
  date: "2026-11-19",
  fee: 2.5,
  invoices: [{ number: "R1", amount: 10, dueDate: "2026-09-01", invoiceReceivedDate: "2026-09-03", verzugWarningOnInvoice: true }],
};

for (const arrangement of ARRANGEMENTS) {
  test(`20. ${arrangement}: published and resolved, business then consumer`, async () => {
    const unit = arrangement === "stacked" ? "business-invoice" : "business-row";
    const business = walk(await resolved(arrangement, previewData.mahnung));
    assertEquals(business.filter((n) => n.id?.startsWith(`de-mahnung-fee-${unit}`)).length, 3);
    assertEquals(business.some((n) => n.text?.includes("Verzugspauschale")), true);
    assertEquals(business.some((n) => n.id === "de-mahnung-fee-proof"), false);

    const consumerUnit = unit.replace("business", "consumer");
    const consumer = walk(await resolved(arrangement, consumerRequest));
    assertEquals(consumer.filter((n) => n.id?.startsWith(`de-mahnung-fee-${consumerUnit}`)).length, 1);
    assertEquals(consumer.some((n) => n.text?.includes("Verzugspauschale")), false);
    assertEquals(consumer.some((n) => n.id === "de-mahnung-fee-proof"), true);

    for (const text of [...business, ...consumer].map((n) => n.text ?? "")) assertEquals(text.includes("{{"), false);
  });
}
