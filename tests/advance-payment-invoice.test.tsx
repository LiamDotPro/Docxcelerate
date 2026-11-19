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
import { COMPONENTS } from "docxcelerate/registry";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/advance-payment-invoice.deriver.ts")
>("components/derivers/advance-payment-invoice.deriver.ts");
const { AdvancePaymentInvoice } = await loadRegistryModule<
  typeof import("../registry/components/advance-payment-invoice/advance-payment-invoice.node.tsx")
>("components/advance-payment-invoice/advance-payment-invoice.node.tsx");

type Obj = Record<string, unknown>;
type Figures = import("../registry/components/derivers/advance-payment-invoice.deriver.ts").AdvancePaymentInvoiceFigures;

interface Input {
  order?: Obj;
  deposit?: Obj;
  previous?: unknown;
  invoice?: Obj;
}

const base = (): Required<Pick<Input, "order" | "deposit" | "invoice">> & { previous: unknown } => ({
  order: { reference: " SO-1 ", currency: "GBP", lines: [{ gross: 1000, vatRate: 20 }] },
  deposit: { percent: 30, kind: "prepayment" },
  previous: undefined,
  invoice: { jurisdiction: "UK", issueDate: "2026-11-30" },
});

function inputs(i: Input) {
  const b = base();
  return {
    order: i.order ? { ...b.order, ...i.order } : b.order,
    deposit: i.deposit ? { ...b.deposit, ...i.deposit } : b.deposit,
    previous: "previous" in i ? i.previous : b.previous,
    invoice: i.invoice ? { ...b.invoice, ...i.invoice } : b.invoice,
  };
}

function run(i: Input = {}): Figures {
  const x = inputs(i);
  return mod.run([x.order, x.deposit, x.previous, x.invoice], {} as RuntimeState) as Figures;
}

const lines = (...l: Obj[]) => ({ lines: l });
const fails = (i: Input, text: string) => assertRejects(() => Promise.resolve().then(() => run(i)), Error, text);

test("the percentage is taken of the gross, not net plus VAT", () => {
  const r = run({ order: lines({ gross: 100.05, vatRate: 20 }) });
  assertEquals(r.depositMinor, 3002);
  assertEquals([r.bands[0].vatMinor, r.bands[0].netMinor], [500, 2502]);
  assertEquals(r.orderReference, "SO-1");
});

test("one band, round figures", () => {
  const r = run();
  assertEquals([r.depositMinor, r.bands[0].vatMinor, r.bands[0].netMinor, r.leftMinor, r.bands.length], [
    30000,
    5000,
    25000,
    70000,
    1,
  ]);
  assertEquals(r.depositLabel, "Deposit requested (30% of the order total)");
  assertEquals(r.title, "Advance payment");
});

test("largest remainder spreads the deposit, ties to the earlier band", () => {
  const r = run({
    order: lines({ gross: 100, vatRate: 20 }, { gross: 100, vatRate: 5 }, { gross: 100, vatRate: 0 }),
    deposit: { percent: undefined, amount: 100 },
  });
  assertEquals(r.bands.map((b) => b.shareMinor), [3334, 3333, 3333]);
  assertEquals(r.bands.map((b) => b.vatMinor), [556, 159, 0]);
  assertEquals(r.bands.map((b) => b.netMinor), [2778, 3174, 3333]);
  assertEquals(r.bands.reduce((s, b) => s + b.shareMinor, 0), 10000);
  assertEquals(r.depositLabel, "Deposit requested");
});

test("VAT on net lines is worked once per band", () => {
  const r = run({ order: lines({ net: 0.03, vatRate: 20 }, { net: 0.03, vatRate: 20 }), deposit: { percent: 100 } });
  assertEquals([r.orderTotalMinor, r.depositMinor, r.leftMinor], [7, 7, 0]);
});

test("exactly 100% is allowed and prints zero left", () => {
  const r = run({ previous: [{ number: "A", amount: 700 }] });
  assertEquals(r.leftMinor, 0);
  assertEquals(r.leftToInvoice, "£0.00");
});

test("over 100% names both figures", async () => {
  await fails({ previous: [{ number: "A", amount: 700.01 }] }, "deposits requested (£1,000.01) exceed the order total (£1,000.00)");
  await fails({ previous: [{ number: "A", amount: 1200 }] }, "exceed the order total");
  // A security deposit is not added to the cap, so only the P > G check can refuse this.
  await fails({ previous: [{ number: "A", amount: 1200 }], deposit: { kind: "security" } }, "exceed the order total");
});

test("a zero-exponent currency", () => {
  const r = run({
    order: { currency: "JPY", ...lines({ gross: 100001, vatRate: 10 }, { gross: 50001, vatRate: 8 }) },
    deposit: { percent: 33.3 },
  });
  assertEquals([r.orderTotalMinor, r.depositMinor], [150002, 49951]);
  assertEquals(r.bands.map((b) => b.shareMinor), [33301, 16650]);
  assertEquals(r.bands.map((b) => b.vatMinor), [3027, 1233]);
  assertEquals(r.depositLabel, "Deposit requested (33.3% of the order total)");
});

const PREVIEW = COMPONENTS.find((c) => c.id === "advance-payment-invoice")!.previewData as unknown as {
  order: Obj;
  deposit: Obj;
  previousDeposits: unknown;
  invoice: Obj;
};

test("the preview data prints the figures the catalog promises", () => {
  const r = mod.run([PREVIEW.order, PREVIEW.deposit, PREVIEW.previousDeposits, PREVIEW.invoice], {} as RuntimeState) as Figures;
  assertEquals(r.orderTotal, "£1,814,110.07");
  assertEquals(r.previous, [
    { number: "ADV-2026-11-000417-A", amount: "£250,000.00" },
    { number: "ADV-2026-11-000418-B", amount: "£125,000.00" },
  ]);
  assertEquals([r.depositLabel, r.deposit, r.leftToInvoice], [
    "Deposit requested (30% of the order total)",
    "£544,233.02",
    "£894,877.05",
  ]);
  assertEquals([r.paymentDateLabel, r.paymentDate], ["Date of payment", "28 November 2026"]);
  assertEquals(r.bands, [
    { label: "20%", net: "£321,187.50", vat: "£64,237.50", gross: "£385,425.00", shareMinor: 38542500, netMinor: 32118750, vatMinor: 6423750 },
    { label: "5%", net: "£123,700.00", vat: "£6,185.00", gross: "£129,885.00", shareMinor: 12988500, netMinor: 12370000, vatMinor: 618500 },
    { label: "0%", net: "£28,923.02", vat: "£0.00", gross: "£28,923.02", shareMinor: 2892302, netMinor: 2892302, vatMinor: 0 },
  ]);
  assertEquals([r.netTotal, r.vatTotal], ["£473,810.52", "£70,422.50"]);
  assertEquals(r.previousMinor + r.depositMinor + r.leftMinor, r.orderTotalMinor);
});

test("many previous deposits keep their order", () => {
  const previous = Array.from({ length: 8 }, (_, i) => ({ number: "P" + i, amount: 1 }));
  const r = run({ previous, deposit: { percent: 10 } });
  assertEquals(r.previous.map((p) => p.number), previous.map((p) => p.number));
  assertEquals([r.previous.length, r.leftMinor], [8, 89200]);
})

test("more than 8 previous deposits or 4 rates are refused, not dropped", async () => {
  await fails({ previous: Array.from({ length: 9 }, (_, i) => ({ number: "P" + i, amount: 1 })) }, "at most 8");
  await fails({ order: lines(...[0, 1, 2, 3, 4].map((r) => ({ gross: 100, vatRate: r }))) }, "at most 4");
});

test("Germany, with a payment date", () => {
  const r = run({
    order: lines({ net: 10000, vatRate: 19 }, { net: 2500, vatRate: 7 }),
    deposit: { percent: 40, receivedDate: "2026-12-02" },
    invoice: { jurisdiction: "DE", issueDate: "2026-11-30" },
  });
  assertEquals([r.orderTotalMinor, r.depositMinor, r.leftMinor], [1457500, 583000, 874500]);
  assertEquals(r.bands.map((b) => [b.shareMinor, b.vatMinor]), [[47600 * 10, 76000], [10700 * 10, 7000]]);
  assertEquals([r.paymentDateLabel, r.title], ["Zeitpunkt der Vereinnahmung", "Anzahlung"]);
});

test("the payment date prints only on a different day", () => {
  const same = run({ deposit: { receivedDate: "2026-11-30T09:00:00Z" } });
  assertEquals([same.paymentDate, same.paymentDateLabel], [null, null]);
  assertEquals(run().paymentDate, null);
  assertEquals(run({ deposit: { receivedDate: "" } }).paymentDate, null);
});

test("US: no VAT, no analysis, no payment date", () => {
  const r = run({
    order: { currency: "USD", ...lines({ net: 12500 }, { net: 12500 }) },
    deposit: { percent: 25, receivedDate: "2026-12-02" },
    invoice: { jurisdiction: "US" },
  });
  assertEquals([r.depositMinor, r.leftMinor, r.vatApplies, r.bands, r.paymentDate], [625000, 1875000, false, [], null]);
  assertEquals(r.deposit, "$6,250.00");
});

test("a security deposit carries no VAT and does not reduce what is left", () => {
  const r = run({
    order: lines({ gross: 10000, vatRate: 20 }),
    deposit: { percent: undefined, amount: 2500, kind: "security" },
    previous: [{ number: "A", amount: 3000 }],
  });
  assertEquals([r.vatApplies, r.depositLabel, r.leftMinor, r.title], [false, "Security deposit", 700000, "Security deposit"]);
  const big = run({
    order: lines({ gross: 100, vatRate: 20 }),
    deposit: { percent: undefined, amount: 5000, kind: "security" },
  });
  assertEquals(big.leftMinor, 10000);
  assertEquals(run({ deposit: { kind: "security" } }).depositLabel, "Security deposit (30% of the order total)");
});

test("bands: zero dropped, negatives netted, a negative band refused", () => {
  const dropped = run({ order: lines({ gross: 0, vatRate: 5 }, { gross: 100, vatRate: 20 }) });
  assertEquals(dropped.bands.map((b) => b.label), ["20%"]);
  assertEquals(run({ order: lines({ gross: 100, vatRate: 20 }, { gross: -20, vatRate: 20 }) }).orderTotalMinor, 8000);
  return fails({ order: lines({ gross: -10, vatRate: 5 }, { gross: 100, vatRate: 20 }) }, "5% band");
});

test("deposit amount and percentage are validated", async () => {
  for (const percent of [0, -5, 100.01]) await fails({ deposit: { percent } }, "deposit.percent");
  await fails({ order: lines({ gross: 0.01, vatRate: 20 }), deposit: { percent: 1 } }, "rounds to zero");
  await fails({ deposit: { amount: 10 } }, "exactly one of percent or amount");
  await fails({ deposit: { percent: undefined } }, "exactly one of percent or amount");
  await fails({ deposit: { percent: undefined, amount: 0 } }, "deposit.amount");
});

test("order lines are validated", async () => {
  await fails({ order: lines({ gross: 1, net: 1, vatRate: 20 }) }, "order.lines[0]");
  await fails({ order: lines({ vatRate: 20 }) }, "order.lines[0]");
  await fails({ order: lines({ gross: 1 }) }, "order.lines[0].vatRate");
  await fails({ order: lines({ gross: 1, vatRate: -1 }) }, "order.lines[0].vatRate");
  await fails({ order: { lines: undefined } }, "order.lines");
  await fails({ order: { lines: [] } }, "order.lines");
});

test("a missing order date prints nothing; missing previous deposits mean none", () => {
  const r = run();
  assertEquals([r.orderDate, r.previous], [null, []]);
  assertEquals(run({ order: { date: "2026-10-14" } }).orderDate, "14 October 2026");
});

test("bad previous entries, enums, currency and dates throw naming the field", async () => {
  await fails({ previous: [{ number: " ", amount: 1 }] }, "previousDeposits[0].number");
  await fails({ previous: [{ number: "A", amount: 0 }] }, "previousDeposits[0].amount");
  await fails({ previous: [{ number: "A", amount: -1 }] }, "previousDeposits[0].amount");
  await fails({ deposit: { kind: "gift" } }, "deposit.kind");
  await fails({ invoice: { jurisdiction: "AU" } }, "invoice.jurisdiction");
  await fails({ order: { currency: "pounds" } }, "order.currency");
  await fails({ order: { reference: " " } }, "order.reference");
  await fails({ invoice: { issueDate: "30/11/2026" } }, "invoice.issueDate");
  await fails({ order: { date: "soon" } }, "order.date");
  await fails({ deposit: { receivedDate: "soon" } }, "deposit.receivedDate");
});

// The node.
const nodeData = (over: Input = {}) => {
  const x = inputs(over);
  return { order: x.order, deposit: x.deposit, previousDeposits: x.previous, invoice: x.invoice };
};
type Data = ReturnType<typeof nodeData>;

function documentOf(arrangement?: "table" | "stacked") {
  return template(
    <Document title="Deposit">
      <AdvancePaymentInvoice arrangement={arrangement} />
    </Document>,
  );
}

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[]; rows?: DocumentNode[]; cells?: DocumentNode[] };
    return [...(own.id ? [own.id] : []), ...ids([...(own.children ?? []), ...(own.rows ?? []), ...(own.cells ?? [])])];
  });
}

function kinds(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { children?: DocumentNode[]; rows?: DocumentNode[]; cells?: DocumentNode[] };
    return [node.kind, ...kinds([...(own.children ?? []), ...(own.rows ?? []), ...(own.cells ?? [])])];
  });
}

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { title?: string; text?: string; children?: DocumentNode[]; rows?: DocumentNode[]; cells?: DocumentNode[] };
    return [
      ...(node.kind === "section" && own.title ? [own.title] : []),
      ...(own.text ? [own.text] : []),
      ...texts([...(own.children ?? []), ...(own.rows ?? []), ...(own.cells ?? [])]),
    ];
  });
}

const buildLocal = (data: Data, arrangement?: "table" | "stacked") =>
  buildDocument(documentOf(arrangement), data, { dynamicMode: "placeholder" });

test("exactly one tax arm prints, and it is the right one", async () => {
  const cases: Array<[Input, string]> = [
    [{ deposit: { kind: "security" } }, "security"],
    [{}, "uk"],
    [{ invoice: { jurisdiction: "EU" } }, "eu"],
    [{ invoice: { jurisdiction: "FR" } }, "eu"],
    [{ invoice: { jurisdiction: "DE" } }, "de"],
    [{ order: lines({ gross: 100 }), invoice: { jurisdiction: "US" } }, "us"],
  ];
  for (const [input, arm] of cases) {
    const built = await buildLocal(nodeData(input));
    assertEquals(ids(built.nodes).filter((id) => id.startsWith("advance-payment-tax-")), ["advance-payment-tax-" + arm]);
  }
});

test("both arrangements print the same line ids, in different nodes", async () => {
  const data = nodeData({ previous: [{ number: "A", amount: 100 }], order: { date: "2026-10-14" }, deposit: { receivedDate: "2026-12-02" } });
  const t = await buildLocal(data, "table");
  const s = await buildLocal(data, "stacked");
  const lineIds = (n: DocumentNode[]) => [...new Set(ids(n))].filter((id) => id.startsWith("advance-payment") && id !== "advance-payment-figures" && id !== "advance-payment-vat").sort();
  assertEquals(lineIds(t.nodes), lineIds(s.nodes));
  assertEquals(ids(t.nodes).includes("advance-payment-order-date"), true);
  assertEquals(kinds(t.nodes).includes("table"), true);
  assertEquals(kinds(s.nodes).includes("table"), false);
  assertEquals((await buildLocal(data)).nodes, t.nodes);
});

test("a missing order date and no previous deposits leave their nodes out", async () => {
  const built = await buildLocal(nodeData());
  assertEquals(ids(built.nodes).includes("advance-payment-order-date"), false);
  assertEquals(ids(built.nodes).includes("advance-payment-previous-1"), false);
});

const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ advancePaymentInvoice: mod.run }));

for (const arrangement of ["table", "stacked"] as const) {
  test(`${arrangement}: published, bundled and resolved matches a local build`, async () => {
    const published = await buildDocument(documentOf(arrangement), createPublishData() as Data, {
      branchMode: "publish",
      deriverMode: "preserve",
      dynamicMode: "resolve",
      aiClient: { generateParagraph: () => "" },
    });
    const serialised = JSON.stringify(published.nodes);
    for (const arm of ["security", "uk", "eu", "de", "us"]) {
      assertEquals(serialised.includes("advance-payment-tax-" + arm), true);
    }
    assertEquals(serialised.includes("{{derived.advancePaymentInvoice"), true);

    const data = {
      order: PREVIEW.order,
      deposit: PREVIEW.deposit,
      previousDeposits: PREVIEW.previousDeposits,
      invoice: PREVIEW.invoice,
    } as unknown as Data;
    const state: RuntimeState = {
      ctx: { ...data },
      derived: {},
      dataProvider: new InMemoryDataProvider(data as unknown as Record<string, unknown>),
      aiClient: { generateParagraph: () => "" },
    };
    const resolved = await resolveDocument(published, state, { derivers: registry });
    const direct = await buildLocal(data, arrangement);

    assertEquals(texts(resolved.nodes), texts(direct.nodes));
    assertEquals(texts(direct.nodes).some((t) => t.startsWith("VAT is charged on this payment now")), true);
  });
}
