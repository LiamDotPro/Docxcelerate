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
import { COMPONENTS } from "docxcelerate/registry";
import { registryRoot } from "docxcelerate/registry/install";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import { transformDocumentSource } from "docxcelerate/transform";
import { assertEquals } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/final-invoice-settlement/final-invoice-settlement.node.tsx";
const { FinalInvoiceSettlement } = await loadRegistryModule<
  typeof import("../registry/components/final-invoice-settlement/final-invoice-settlement.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/final-invoice-settlement.deriver.ts")
>("components/derivers/final-invoice-settlement.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ finalInvoiceSettlement: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
type Out = ReturnType<typeof mod.run> extends infer R ? Awaited<R> : never;
const res = (o: unknown) => o as {
  vatShown: boolean;
  bands: Obj[];
  supply: Obj;
  deducted: Obj[];
  overpaid: boolean;
  due: { net: string; vat: string; gross: string };
  correction: string;
  pendingOpen: string;
  pendingDeducted: string;
};
void (null as unknown as Out);

const run = (invoice: Obj, seller: Obj | undefined = {}, locale = "en-GB") =>
  res(mod.run([invoice, seller, locale], {} as RuntimeState));
const throws = (invoice: Obj, message: string, seller: Obj = {}) => {
  let got = "";
  try {
    run(invoice, seller);
  } catch (error) {
    got = error instanceof Error ? error.message : String(error);
  }
  if (!got.includes(message)) throw new Error(`expected "${message}", got "${got}"`);
};
const adv = (number: string, issueDate: string, net: number, vat: number, vatRate: number, received = true, more: Obj = {}) => ({
  number, issueDate, net, vat, vatRate, received, ...more,
});
const gbp = (lines: Obj[], advances?: Obj[]) => ({ jurisdiction: "UK", currency: "GBP", lines, advances });
const eur = (jurisdiction: string, lines: Obj[], advances?: Obj[]) => ({ jurisdiction, currency: "EUR", lines, advances });
const row = (b: Obj) => [b.rate, b.net, b.vat, b.gross];
const due = (o: ReturnType<typeof run>) => [o.due.net, o.due.vat, o.due.gross];
const filled = (rows: Obj[], key: string) => rows.filter((r) => r[key]);

test("UK deposit, no rate change", () => {
  const o = run(gbp([{ net: 1000000, vatRate: 20 }], [adv("D1", "2026-03-01", 200000, 40000, 20)]));
  assertEquals(row(o.bands[0]), ["20%", "£10,000.00", "£2,000.00", "£12,000.00"]);
  assertEquals(row(o.deducted[0]).slice(1), ["−£2,000.00".replace("−", "-"), "-£400.00", "-£2,400.00"]);
  assertEquals(due(o), ["£8,000.00", "£1,600.00", "£9,600.00"]);
  assertEquals([o.correction, o.pendingOpen, o.pendingDeducted], ["", "", ""]);
});

test("UK rate change carves the deposit out at its own rate", () => {
  const o = run(gbp([{ net: 5000000, vatRate: 20 }], [adv("D1", "2026-03-01", 1000000, 175000, 17.5)]));
  assertEquals(filled(o.bands, "rate").map(row), [
    ["20%", "£40,000.00", "£8,000.00", "£48,000.00"],
    ["17.5%", "£10,000.00", "£1,750.00", "£11,750.00"],
  ]);
  assertEquals(Object.values(o.supply), ["£50,000.00", "£9,750.00", "£59,750.00"]);
  assertEquals(due(o), ["£40,000.00", "£8,000.00", "£48,000.00"]);
  assertEquals(o.correction, "");
});

test("DE, FR and EU re-rate the whole supply and print the difference", () => {
  for (const j of ["DE", "FR", "EU"]) {
    const o = run(eur(j, [{ net: 10000000, vatRate: 19 }], [adv("A1", "2026-03-01", 5000000, 800000, 16)]));
    assertEquals(filled(o.bands, "rate").map(row), [["19%", "€100,000.00", "€19,000.00", "€119,000.00"]]);
    assertEquals(row(o.deducted[0]).slice(1), ["-€50,000.00", "-€8,000.00", "-€58,000.00"]);
    assertEquals(due(o), ["€50,000.00", "€11,000.00", "€61,000.00"]);
    assertEquals(o.correction, "€1,500.00");
  }
});

test("DE unpaid advance stays in the amount due", () => {
  const lines = [{ net: 10000000, vatRate: 19 }];
  const o = run(eur("DE", lines, [
    adv("A1", "2026-04-01", 3000000, 570000, 19),
    adv("A2", "2026-05-01", 2000000, 380000, 19, false),
  ]));
  assertEquals(filled(o.deducted, "number").length, 1);
  assertEquals(due(o), ["€70,000.00", "€13,300.00", "€83,300.00"]);
  assertEquals([o.pendingOpen, o.pendingDeducted], ["A2 (1 May 2026, €23,800.00)", ""]);
});

test("UK unpaid advance is deducted and listed as still payable", () => {
  const o = run(gbp([{ net: 10000000, vatRate: 20 }], [
    adv("A1", "2026-04-01", 3000000, 600000, 20),
    adv("A2", "2026-05-01", 2000000, 400000, 20, false),
  ]));
  assertEquals(filled(o.deducted, "number").length, 2);
  assertEquals(due(o), ["£50,000.00", "£10,000.00", "£60,000.00"]);
  assertEquals([o.pendingDeducted, o.pendingOpen], ["A2 (1 May 2026, £24,000.00)", ""]);
});

test("overpaid is not a negative amount due", () => {
  const o = run(eur("DE", [{ net: 1000000, vatRate: 19 }], [adv("A1", "2026-03-01", 1200000, 228000, 19)]));
  assertEquals([o.overpaid, ...due(o)], [true, "€2,000.00", "€380.00", "€2,380.00"]);
  throws(gbp([{ net: 1000000, vatRate: 20 }], [
    adv("A1", "2026-03-01", 1000000, 200000, 20),
    adv("A2", "2026-03-02", 200000, 40000, 20, false),
  ]), "cancel the unpaid advance invoice with a credit note first");
});

test("settled exactly is zero, not overpaid", () => {
  const o = run(eur("DE", [{ net: 1000000, vatRate: 19 }], [adv("A1", "2026-03-01", 1000000, 190000, 19)]));
  assertEquals([o.overpaid, ...due(o)], [false, "€0.00", "€0.00", "€0.00"]);
});

test("small business and US print no VAT", () => {
  const sb = run(eur("DE", [{ net: 500000 }], [{ number: "A1", issueDate: "2026-03-01", net: 200000, received: true }]), {
    smallBusinessExempt: true,
  });
  assertEquals([sb.vatShown, sb.supply.net, row(sb.deducted[0]), sb.due.net], [
    false, "€5,000.00", ["", "-€2,000.00", "", "-€2,000.00"], "€3,000.00",
  ]);
  assertEquals([filled(sb.bands, "rate").length, sb.correction], [0, ""]);
  const exempt = { smallBusinessExempt: true };
  throws(eur("DE", [{ net: 500000, vatRate: 19 }]), "no VAT is shown", exempt);
  throws(eur("DE", [{ net: 500000 }], [{ ...adv("A1", "2026-03-01", 200000, 38000, 19), vatRate: undefined }]), "no VAT is shown", exempt);
  const us = run({ jurisdiction: "US", currency: "USD", lines: [{ net: 250000 }], advances: [{ number: "A", issueDate: "2026-01-01", net: 100000, received: true }] });
  assertEquals([us.vatShown, us.due.net], [false, "US$1,500.00"]);
  throws({ jurisdiction: "US", currency: "USD", lines: [{ net: 250000, vatRate: 8.25 }] }, "no VAT is shown");
});

test("no advances", () => {
  for (const advances of [undefined, []]) {
    const o = run(gbp([{ net: 1000000, vatRate: 20 }], advances));
    assertEquals(due(o), ["£10,000.00", "£2,000.00", "£12,000.00"]);
    assertEquals([filled(o.deducted, "number").length, o.deducted.length, o.pendingOpen], [0, 8, ""]);
  }
});

test("three rates including 0% print in order", () => {
  const o = run(gbp([20, 5, 0].map((vatRate) => ({ net: 100000, vatRate }))));
  assertEquals(filled(o.bands, "rate").map((b) => [b.rate, b.vat]), [["20%", "£200.00"], ["5%", "£50.00"], ["0%", "£0.00"]]);
});

test("rounding is half away from zero, once per band", () => {
  const vat = (lines: Obj[]) => {
    const o = run(gbp(lines));
    return Number(String(o.supply.vat).replace(/[£.,]/g, "")) - 0;
  };
  const cases: Array<[Obj[], number]> = [
    [[{ net: 5, vatRate: 10 }], 1],
    [[{ net: 4, vatRate: 10 }], 0],
    [[{ net: 10, vatRate: 5.5 }], 1],
    [[{ net: 9, vatRate: 5.5 }], 0],
    [[{ net: 1000, vatRate: 5.5 }], 55],
    [[{ net: 5, vatRate: 10 }, { net: 5, vatRate: 10 }], 1],
  ];
  for (const [lines, want] of cases) assertEquals(vat(lines), want);
});

test("advance VAT is deducted as printed", () => {
  const o = run(eur("DE", [{ net: 1000000, vatRate: 19 }], [adv("A1", "2026-03-01", 100000, 19001, 19)]));
  assertEquals(o.deducted[0].vat, "-€190.01");
});

test("the target band of a re-rated advance", () => {
  const two = [{ net: 1000000, vatRate: 19 }, { net: 1000000, vatRate: 7 }];
  const a = (more: Obj = {}) => [adv("A1", "2026-03-01", 100000, 16000, 16, true, more)];
  throws(eur("DE", two, a()), "give its supplyRate");
  assertEquals(run(eur("DE", two, a({ supplyRate: 19 }))).correction, "€30.00");
  throws(eur("DE", two, a({ supplyRate: 10 })), "supplyRate");
  assertEquals(run(eur("DE", [{ net: 1000000, vatRate: 19 }], a())).correction, "€30.00");
});

test("a UK carve-out larger than its band throws", () => {
  throws(gbp([{ net: 100000, vatRate: 20 }], [adv("D", "2026-03-01", 150000, 26250, 17.5)]), "exceeds the supply it was paid towards");
});

test("discount lines net within a band", () => {
  const o = run(gbp([{ net: 1000000, vatRate: 20 }, { net: -100000, vatRate: 20 }]));
  assertEquals(o.bands[0].net, "£9,000.00");
  throws(gbp([{ net: 1000000, vatRate: 20 }, { net: -2000000, vatRate: 5 }]), "credit note");
  throws(gbp([{ net: 1000, vatRate: 20 }, { net: -1000, vatRate: 20 }]), "nets to zero");
});

test("slots: eight, nine, five rates, and nine unpaid", () => {
  const many = (n: number, received = true) =>
    Array.from({ length: n }, (_, i) => adv("A" + i, "2026-03-01", 1000, 190, 19, received));
  const lines = [{ net: 10000000, vatRate: 19 }];
  assertEquals(filled(run(eur("DE", lines, many(8))).deducted, "number").length, 8);
  throws(eur("DE", lines, many(9)), "more than eight advances");
  throws(gbp([1, 2, 3, 4, 5].map((vatRate) => ({ net: 100, vatRate }))), "more than four VAT rates");
  const o = run(eur("DE", lines, many(9, false)));
  assertEquals([filled(o.deducted, "number").length, o.pendingOpen.split(" (").length - 1], [0, 9]);
});

test("advances are deducted by issue date, ties in input order", () => {
  const o = run(eur("DE", [{ net: 1000000, vatRate: 19 }], [
    adv("B", "2026-06-01", 1000, 190, 19),
    adv("A1", "2026-03-01", 1000, 190, 19),
    adv("A2", "2026-03-01T10:00:00Z", 1000, 190, 19),
  ]));
  assertEquals(filled(o.deducted, "number").map((d) => d.number), ["A1", "A2", "B"]);
});

test("invalid input throws", () => {
  const l = [{ net: 1000000, vatRate: 19 }];
  const a = adv("A1", "2026-03-01", 1000, 190, 19);
  throws({ ...eur("IE", l) }, "invoice.jurisdiction");
  throws({ ...eur("DE", l), currency: "EURO" }, "invoice.currency");
  throws({ jurisdiction: "DE", currency: "EUR" }, "invoice.lines");
  throws(eur("DE", []), "invoice.lines");
  throws(eur("DE", [{ net: 12.5, vatRate: 19 }]), "net");
  for (const vatRate of [19.123, 100, -1]) throws(eur("DE", [{ net: 100, vatRate }]), "vatRate");
  throws(eur("DE", l, [{ ...a, number: "  " }]), "number");
  throws(eur("DE", l, [a, { ...a }]), "used twice");
  throws(eur("DE", l, [{ ...a, issueDate: "31/03/2026" }]), "issueDate");
  throws(eur("DE", l, [{ ...a, received: undefined }]), "received");
  throws(eur("DE", l, [{ ...a, net: 0 }]), "net");
  throws(eur("DE", l, [{ ...a, vat: undefined }]), "vat");
});

test("amounts and percents follow the locale", () => {
  const o = run(eur("DE", [{ net: 487654321, vatRate: 19 }]), {}, "de-DE");
  assertEquals([o.bands[0].rate, o.bands[0].net], ["19 %", "4.876.543,21 €"]);
});

// ---------------------------------------------------------------------------
// The node
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Invoice">
      <Section id="body" title="Body">
        <FinalInvoiceSettlement />
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

const built = (data: Obj) =>
  buildDocument(templateOf(), data, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);
const idsOf = async (data: Obj) =>
  collect((await built(data)).nodes).map((n) => n.id).filter((id) => id.startsWith("final-settlement"));
const entry = COMPONENTS.find((c) => c.id === "final-invoice-settlement")!;
const preview = entry.previewData as Obj;
const smallBusiness = {
  invoice: eur("DE", [{ net: 500000 }], [{ number: "A1", issueDate: "2026-03-01", net: 200000, received: true }]),
  seller: { smallBusinessExempt: true },
};

test("the DE preview prints the VAT table with a correction and a pending note", async () => {
  assertEquals(await idsOf(preview), [
    "final-settlement",
    "final-settlement-vat",
    "final-settlement-vat-header",
    "final-settlement-band-1",
    "final-settlement-band-2",
    "final-settlement-vat-supply",
    "final-settlement-vat-advance-1",
    "final-settlement-vat-advance-2",
    "final-settlement-vat-due",
    "final-settlement-vat-correction",
    "final-settlement-pending-open",
  ]);
  const o = run((preview as { invoice: Obj }).invoice);
  assertEquals(due(o), ["€2,299,999.99", "€309,419.75", "€2,609,419.74"]);
  assertEquals([o.correction, o.pendingOpen], ["-€45,000.00", "AR-2020-000171/BAU-NORD-03 (30 September 2020, €928,000.00)"]);
});

test("the plain table prints alone for a seller with no VAT", async () => {
  assertEquals(await idsOf(smallBusiness), [
    "final-settlement",
    "final-settlement-plain",
    "final-settlement-plain-header",
    "final-settlement-plain-supply",
    "final-settlement-plain-advance-1",
    "final-settlement-plain-due",
  ]);
});

test("overpaid prints its row and the credit note instead of the due row", async () => {
  const ids = await idsOf({
    invoice: eur("DE", [{ net: 1000000, vatRate: 19 }], [adv("A1", "2026-03-01", 1200000, 228000, 19)]),
    seller: {},
  });
  assertEquals(ids.includes("final-settlement-vat-overpaid") && ids.includes("final-settlement-credit-note"), true);
  assertEquals(ids.includes("final-settlement-vat-due"), false);
});

test("eight advances print eight rows", async () => {
  const many = Array.from({ length: 8 }, (_, i) => adv("A" + i, "2026-03-01", 1000, 190, 19));
  const ids = await idsOf({ invoice: eur("DE", [{ net: 10000000, vatRate: 19 }], many), seller: {} });
  for (let i = 1; i <= 8; i++) assertEquals(ids.includes(`final-settlement-vat-advance-${i}`), true);
});

// ---------------------------------------------------------------------------
// Publishing
// ---------------------------------------------------------------------------

function state(data: Obj): RuntimeState {
  return { ctx: { ...data }, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } };
}

test("a publish build carries the arms its doc comment states, and resolves like a direct build", async () => {
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  const arms = all.filter((n) => (n as { when?: unknown }).when !== undefined);
  assertEquals(arms.length, 30);
  assertEquals(branches, 27);

  const cases: Obj[] = [
    { invoice: gbp([{ net: 1000000, vatRate: 20 }], [adv("D1", "2026-03-01", 200000, 40000, 20)]), seller: {} },
    { invoice: eur("DE", [{ net: 10000000, vatRate: 19 }], [adv("A1", "2026-03-01", 5000000, 800000, 16)]), seller: {} },
    { invoice: eur("DE", [{ net: 1000000, vatRate: 19 }], [adv("A1", "2026-03-01", 1200000, 228000, 19)]), seller: {} },
    smallBusiness,
    preview,
  ];
  const seen = (nodes: DocumentNode[]) =>
    collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);
  for (const data of cases) {
    const direct = await built(data);
    const resolved = await resolveDocument(published, state(data), { derivers });
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
  }
});
