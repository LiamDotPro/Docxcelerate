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
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/disbursements-and-recharges/disbursements-and-recharges.node.tsx";
const { DisbursementsAndRecharges } = await loadRegistryModule<
  typeof import("../registry/components/disbursements-and-recharges/disbursements-and-recharges.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/disbursements-and-recharges.deriver.ts")
>("components/derivers/disbursements-and-recharges.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ disbursementsAndRecharges: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Row {
  date: string;
  description: string;
  amount: string;
  amountMinor: number;
}
interface Out {
  any: boolean;
  hasRecharges: boolean;
  hasDisbursements: boolean;
  showVat: boolean;
  rechargeTitle: string;
  rechargeStatement: string;
  recharges: Row[];
  rechargeNet: string;
  vatLabel: string;
  vat: string;
  rechargeTotal: string;
  disbursementTitle: string;
  disbursements: Row[];
  disbursementTotal: string;
  notes: string;
  rechargeNetMinor: number;
  vatMinor: number;
  rechargeTotalMinor: number;
  disbursementTotalMinor: number;
}

const UK = { jurisdiction: "UK", currency: "GBP", mainSupplyVatRate: 20 };
const US = { jurisdiction: "US", currency: "USD", mainSupplyVatRate: null };
const run = (invoice: Obj, expenses?: Obj[], strict = false): Out =>
  mod.run([expenses ? { ...invoice, expenses } : invoice, strict], {} as RuntimeState) as Out;
const rejects = (invoice: Obj, expenses: Obj[], text: string, strict = false) =>
  assertRejects(() => Promise.resolve().then(() => run(invoice, expenses, strict)), Error, text);
const ex = (date: string, description: string, net: number, more: Obj = {}) => ({ date, description, net, ...more });
const lines = (rows: Row[]) => rows.filter((r) => r.description).map((r) => r.amount);
const descs = (rows: Row[]) => rows.filter((r) => r.description).map((r) => r.description);

const basic = [
  ex("2026-10-01", "Court fee", 455, { category: "court-fee", paidAsAgent: true }),
  ex("2026-10-02", "Train fare", 182.4, { category: "travel" }),
  ex("2026-10-03", "Courier", 25, { inputVatReclaimed: 5 }),
];

test("basic split", () => {
  const o = run(UK, basic);
  assertEquals(lines(o.recharges), ["£182.40", "£25.00"]);
  assertEquals([o.rechargeNet, o.vat, o.rechargeTotal], ["£207.40", "£41.48", "£248.88"]);
  assertEquals([o.vatLabel, o.showVat, o.rechargeTitle, o.disbursementTitle], ["VAT at 20%", true, "Recharged expenses", "Disbursements"]);
  assertEquals(lines(o.disbursements), ["£455.00"]);
  assertEquals(o.disbursementTotal, "£455.00");
  assertEquals(o.notes, "");
  assertEquals([o.recharges.length, o.disbursements.length], [8, 6]);
  assertStringIncludes(o.rechargeStatement, "at 20%");
  assertEquals(o.recharges[0].date, "2 Oct 2026");
});

test("a zero-rated fare takes the main supply's rate", () => {
  const o = run(UK, [ex("2026-10-02", "Fare", 182.4, { category: "travel" })]);
  assertEquals(o.vatMinor, 3648);
});

test("travel marked as agent is recharged with a note", () => {
  const o = run(UK, [ex("2026-10-02", "Train fare", 182.4, { category: "travel", paidAsAgent: true })]);
  assertEquals(o.hasDisbursements, false);
  assertEquals(
    o.notes,
    "Train fare: recharged with VAT, not passed on as a disbursement, because travel is a cost of providing our services, not a payment made on your behalf.",
  );
});

test("only the first reason is kept", () => {
  const o = run(UK, [ex("2026-10-02", "Fare", 100, { paidAsAgent: true, inputVatReclaimed: 20, markupPercent: 10, category: "travel" })]);
  assertEquals(o.notes, "Fare: recharged with VAT, not passed on as a disbursement, because we reclaimed the VAT on it.");
});

test("a markup forces a recharge", () => {
  const o = run(UK, [ex("2026-10-02", "Counsel", 1000, { paidAsAgent: true, markupPercent: 10 })]);
  assertEquals(lines(o.recharges), ["£1,100.00"]);
  assertStringIncludes(o.notes, "because it includes a handling charge.");
});

test("every excluded category is a recharge; others are not", () => {
  for (const c of ["travel", "subsistence", "accommodation", "postage", "telephone", "photocopying", "printing", "electronic-search"]) {
    assertEquals(run(UK, [ex("2026-10-02", "X", 10, { category: c, paidAsAgent: true })]).hasRecharges, true);
  }
  assertEquals(run(UK, [ex("2026-10-02", "X", 10, { category: "toString", paidAsAgent: true })]).hasDisbursements, true);
  assertStringIncludes(
    run(UK, [ex("2026-10-02", "X", 10, { category: "electronic-search", paidAsAgent: true })]).notes,
    "electronic search fees is a cost",
  );
});

test("strict throws on a reclassification and matches non-strict on clean data", async () => {
  await rejects(UK, [ex("2026-10-02", "Fare", 182.4, { category: "travel", paidAsAgent: true })], "invoice.expenses[0]", true);
  await rejects(UK, [ex("2026-10-02", "Fare", 182.4, { category: "travel", paidAsAgent: true })], "strict mode", true);
  assertEquals(run(UK, basic, true), run(UK, basic, false));
});

test("foreign receipts: half unit rounds up exactly, markup, irrecoverable tax", () => {
  const eur = { currency: "EUR", fxRate: 0.8571 };
  const a = run(UK, [ex("2026-10-02", "Hotel", 1250, eur)]);
  assertEquals(a.recharges[0].amountMinor, 107138);
  assertEquals(a.recharges[0].description, "Hotel: €1,250.00 at 0.8571 = £1,071.38");
  assertEquals(run(UK, [ex("2026-10-02", "Hotel", 1250, { ...eur, markupPercent: 10 })]).recharges[0].amountMinor, 117852);
  const t = run(UK, [ex("2026-10-02", "Hotel", 1250, { ...eur, irrecoverableTax: 87.5 })]);
  assertEquals(t.recharges[0].amount, "£1,146.37");
  assertStringIncludes(t.recharges[0].description, "€1,337.50");
});

test("fxRate is ignored when the currencies match", () => {
  const o = run(UK, [ex("2026-10-02", "X", 10, { currency: "gbp", fxRate: 2 })]);
  assertEquals([o.recharges[0].amount, o.recharges[0].description], ["£10.00", "X"]);
});

test("a zero-exponent invoice currency", async () => {
  const jpy = { ...UK, currency: "JPY" };
  assertEquals(run(jpy, [ex("2026-10-02", "X", 100, { currency: "EUR", fxRate: 162.35 })]).recharges[0].amount, new Intl.NumberFormat("en-GB", { style: "currency", currency: "JPY" }).format(16235));
  await rejects(jpy, [ex("2026-10-02", "X", 1.5)], "invoice.expenses[0].net");
});

test("VAT is worked once on the net, half away from zero", () => {
  const five = { ...UK, mainSupplyVatRate: 5 };
  assertEquals(run(five, [ex("2026-10-02", "X", 0.1)]).vat, "£0.01");
  assertEquals(run(five, [ex("2026-10-02", "X", -0.1)]).vat, "-£0.01");
});

test("negative lines", () => {
  const o = run(UK, [ex("2026-10-02", "A", 182.4), ex("2026-10-03", "B", -50)]);
  assertEquals([o.rechargeNet, o.vat, o.rechargeTotal], ["£132.40", "£26.48", "£158.88"]);
  assertEquals(lines(run(UK, [ex("2026-10-02", "R", -455, { paidAsAgent: true })]).disbursements), ["-£455.00"]);
});

test("zero lines are dropped", () => {
  const o = run(UK, [ex("2026-10-02", "A", 0), ex("2026-10-02", "B", 0.01, { currency: "EUR", fxRate: 0.4 })]);
  assertEquals([o.any, o.hasRecharges, o.rechargeTotal, o.notes], [false, false, "", ""]);
  assertEquals(run(UK, [ex("2026-10-02", "A", 0, { paidAsAgent: true, category: "travel" })]).notes, "");
});

test("missing or empty expenses print nothing and read nothing else", () => {
  for (const inv of [{}, { expenses: [] }, { expenses: "x" }]) {
    const o = mod.run([inv, false], {} as RuntimeState) as Out;
    assertEquals([o.any, o.rechargeTitle, o.notes, o.recharges.length, o.disbursements.length], [false, "", "", 8, 6]);
  }
});

test("one kind only", () => {
  assertEquals(run(UK, [ex("2026-10-02", "A", 1, { paidAsAgent: true })]).hasRecharges, false);
  assertEquals(run(UK, [ex("2026-10-02", "A", 1)]).hasDisbursements, false);
});

test("a null rate carries no VAT", () => {
  const o = run({ ...UK, mainSupplyVatRate: null }, [ex("2026-10-02", "A", 100)]);
  assertEquals([o.showVat, o.vat, o.rechargeTotal, o.vatMinor], [false, "", "£100.00", 0]);
  assertStringIncludes(o.rechargeStatement, "so no VAT is charged on them here");
});

test("a rate of 0 prints VAT at 0%", () => {
  const o = run({ ...UK, mainSupplyVatRate: 0 }, [ex("2026-10-02", "A", 100)]);
  assertEquals([o.showVat, o.vatLabel, o.vat], [true, "VAT at 0%", "£0.00"]);
});

test("DE and FR read as EU", () => {
  const eu = run({ ...UK, jurisdiction: "EU" }, basic);
  assertEquals(run({ ...UK, jurisdiction: "DE" }, basic), eu);
  assertEquals(run({ ...UK, jurisdiction: "FR" }, basic), eu);
});

test("US splits on the taxable flag", async () => {
  const items = [
    ex("2026-10-01", "Airfare", 642.18, { taxable: true }),
    ex("2026-10-02", "Filing fee", 435, { taxable: false }),
    ex("2026-10-03", "Parking", 48),
    ex("2026-10-04", "Agent", 10, { paidAsAgent: true, category: "travel", inputVatReclaimed: 1, taxable: false }),
  ];
  const o = run(US, items);
  assertEquals([o.rechargeTotal, o.disbursementTotal, o.showVat, o.vat], ["$642.18", "$493.00", false, ""]);
  assertEquals(o.notes, "Parking: listed as not taxable, because no taxability was given for it.");
  assertEquals([o.rechargeTitle, o.disbursementTitle], ["Taxable reimbursed expenses", "Non-taxable reimbursed expenses"]);
  assertEquals(run(US, [ex("2026-10-01", "Airfare", 642.18, { taxable: true, markupPercent: 10 })]).rechargeTotal, "$706.40");
  await rejects(US, items, "invoice.expenses[2]", true);
});

test("ordering is by date, then input order", () => {
  const o = run(UK, [
    ex("2026-10-05", "C", 1, { category: "travel", paidAsAgent: true }),
    ex("2026-10-03", "A", 1),
    ex("2026-10-05", "B", 1),
    ex("2026-10-03", "D", 1, { category: "travel", paidAsAgent: true }),
  ]);
  assertEquals(descs(o.recharges), ["A", "D", "C", "B"]);
  assertEquals(o.notes.split("\n").map((n) => n[0]), ["D", "C"]);
});

test("capacity: 8 recharges and 6 disbursements fit, one more throws", async () => {
  const many = (n: number, more: Obj) => Array.from({ length: n }, (_, i) => ex("2026-10-02", "X" + i, 1, more));
  assertEquals(descs(run(UK, many(8, {})).recharges).length, 8);
  assertEquals(descs(run(UK, many(6, { paidAsAgent: true })).disbursements).length, 6);
  await rejects(UK, many(9, {}), "9 recharged expenses; this component prints at most 8");
  await rejects(UK, many(7, { paidAsAgent: true }), "at most 6");
});

test("supplier name", () => {
  assertEquals(run(UK, [ex("2026-10-02", "A", 1, { supplierName: "Acme" })]).recharges[0].description, "A (Acme)");
  assertEquals(run(UK, [ex("2026-10-02", "A", 1, { supplierName: "  " })]).recharges[0].description, "A");
});

test("invalid input throws, naming the field", async () => {
  const ok = ex("2026-10-02", "A", 1);
  const bad: Array<[Obj, Obj, string]> = [
    [UK, { ...ok, date: "02/10/2026" }, "expenses[0].date"],
    [UK, { ...ok, date: "2026-02-30" }, "expenses[0].date"],
    [UK, { ...ok, description: "  " }, "expenses[0].description"],
    [UK, { ...ok, net: Number.NaN }, "expenses[0].net"],
    [UK, { ...ok, net: 12.345 }, "expenses[0].net has more decimals than GBP allows"],
    [UK, { ...ok, inputVatReclaimed: -1 }, "inputVatReclaimed"],
    [UK, { ...ok, irrecoverableTax: -1 }, "irrecoverableTax"],
    [UK, { ...ok, markupPercent: -1 }, "markupPercent"],
    [UK, { ...ok, currency: "EUR" }, "fxRate"],
    [UK, { ...ok, currency: "EUR", fxRate: 0 }, "fxRate"],
    [UK, { ...ok, currency: "EUR", fxRate: 0.1234567 }, "fxRate"],
    [UK, { ...ok, currency: "XYZ", fxRate: 1 }, "expenses[0].currency"],
    [{ ...UK, currency: "XYZ" }, ok, "invoice.currency"],
    [{ ...UK, currency: undefined }, ok, "invoice.currency"],
    [{ ...UK, jurisdiction: "CH" }, ok, "invoice.jurisdiction"],
    [{ ...UK, mainSupplyVatRate: undefined }, ok, "mainSupplyVatRate"],
    [{ ...UK, mainSupplyVatRate: -1 }, ok, "mainSupplyVatRate"],
    [{ ...UK, mainSupplyVatRate: 12.555 }, ok, "mainSupplyVatRate"],
  ];
  for (const [inv, e, field] of bad) await rejects(inv, [e], field);
});

const previewInvoice = () => (COMPONENTS.find((c) => c.id === "disbursements-and-recharges")!.previewData as { invoice: Obj }).invoice;

test("the catalog's preview data", () => {
  const o = mod.run([previewInvoice(), false], {} as RuntimeState) as Out;
  assertEquals(lines(o.recharges), ["£382.40", "£1,146.37", "£2,940.00", "£1,184.60", "£4,524.43"]);
  assertEquals(o.recharges[1].description.endsWith("€1,337.50 at 0.8571 = £1,146.37"), true);
  assertEquals(o.recharges[4].description.endsWith("€4,800.00 at 0.8569 = £4,113.12"), true);
  assertEquals([o.rechargeNet, o.vat, o.rechargeTotal], ["£10,177.80", "£2,035.56", "£12,213.36"]);
  assertEquals(lines(o.disbursements), ["£10,000.00", "£147.00", "£487,250.00", "£18,750.00"]);
  assertEquals(o.disbursementTotal, "£516,147.00");
  const notes = o.notes.split("\n");
  assertEquals(notes.length, 2);
  assertStringIncludes(notes[0], "travel");
  assertStringIncludes(notes[1], "electronic search fees");
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Matter">
      <Section id="body" title="Body">
        <DisbursementsAndRecharges />
      </Section>
    </Document>,
  );

function state(data: Obj): RuntimeState {
  return { ctx: { ...data }, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } };
}

function collect(nodes: DocumentNode[], into: DocumentNode[] = []): DocumentNode[] {
  for (const n of nodes) {
    into.push(n);
    const children = (n as { children?: DocumentNode[] }).children;
    if (children) collect(children, into);
  }
  return into;
}

const built = (invoice: Obj) =>
  buildDocument(templateOf(), { invoice }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);

test("a publish build carries 20 arms and one deriver call", async () => {
  assertEquals(branches, 20);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 20);
  const section = all.find((n) => n.id === "disbursements-and-recharges") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  const d = JSON.parse(JSON.stringify(section.derivers[0]));
  assertEquals(d.name, "disbursementsAndRecharges");
  assertEquals(d.inputs.length, 2);
  assertEquals(d.inputs[0], { type: "ref", ref: { scope: "data", path: "invoice" } });

  const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.id, n.kind, (n as { text?: string }).text]);
  const cases: Array<[Obj, (ids: string[]) => void]> = [
    [{ ...UK, expenses: basic }, (ids) => {
      assertEquals(ids.includes("disbursements-and-recharges-recharge-2"), true);
      assertEquals(ids.includes("disbursements-and-recharges-recharges-vat"), true);
    }],
    [{ ...US, expenses: [ex("2026-10-01", "Airfare", 642.18, { taxable: true }), ex("2026-10-03", "Parking", 48)] }, (ids) => {
      assertEquals(ids.includes("disbursements-and-recharges-recharges-vat"), false);
      assertEquals(ids.includes("disbursements-and-recharges-notes"), true);
    }],
    [previewInvoice(), () => {}],
    [{}, (ids) => assertEquals(ids, [])],
  ];
  for (const [invoice, check] of cases) {
    const direct = await built(invoice);
    const resolved = await resolveDocument(published, state({ invoice }), { derivers });
    assertEquals(seen(resolved.nodes).map((r) => r.slice(1)), seen(direct.nodes).map((r) => r.slice(1)));
    check(seen(resolved.nodes).map((r) => r[0] as string).filter((id) => !id.startsWith("table-cell") && id !== "body"));
  }
  const us = JSON.stringify(
    (await resolveDocument(published, state({ invoice: cases[1][0] }), { derivers })).nodes,
  );
  assertStringIncludes(us, "Taxable reimbursed expenses");
  assertStringIncludes(us, "Non-taxable reimbursed expenses");
});
