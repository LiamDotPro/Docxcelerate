import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { findRegistryEntry } from "docxcelerate/registry";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/deposit-cancellation-outcome.deriver.ts")
>("components/derivers/deposit-cancellation-outcome.deriver.ts");
const { DepositCancellationOutcome } = await loadRegistryModule<
  typeof import("../registry/components/deposit-cancellation-outcome/deposit-cancellation-outcome.node.tsx")
>("components/deposit-cancellation-outcome/deposit-cancellation-outcome.node.tsx");
// Through a bundle, so the publish path proves the function survives being
// shipped as source text, which is how a deriver fails.
const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ depositCancellationOutcome: mod.run }),
);

type Obj = Record<string, unknown>;
const preview = (findRegistryEntry("deposit-cancellation-outcome") as unknown as { previewData: { order: Obj } })
  .previewData.order;
const base: Obj = { reference: "BK-1", currency: "GBP", jurisdiction: "UK", cancelledOn: "2026-11-03" };

type Result = ReturnType<NonNullable<typeof mod.__types>>;

async function run(order: Obj = {}, whole = false): Promise<Result> {
  return await mod.run([whole ? order : { ...base, ...order }], {} as RuntimeState) as Result;
}
const refundable = (gross: number, rate: number, net: number, vat: number, invoiceNumber?: string) => ({
  kind: "refundable-advance",
  gross,
  bands: [{ rate, net, vat }],
  invoiceNumber,
});
const security = (gross: number) => ({ kind: "security", gross });

test("preview data: every figure and sentence", async () => {
  const r = await run(preview, true);
  assertEquals(r.totals, {
    paid: 135775000,
    security: 2500000,
    deductions: 1445040,
    refundable: 128475000,
    retention: 32118750,
    nonRefundable: 4800000,
    refunded: 97411210,
    retained: 38363790,
    refundedVat: 15302678,
    retainedVat: 5900893,
  });
  assertEquals(r.totals.refunded + r.totals.retained, r.totals.paid);
  assertEquals(r.lines.paid, "£1,357,750.00");
  assertEquals(r.lines.refunded, "£974,112.10");
  assertEquals(r.lines.retained, "£383,637.90");
  assertEquals(r.lines.retention, "£321,187.50");
  assertEquals(r.lines.retentionLabel, "Retained: cancellation less than 90 days before the event, clause 14.2");
  assertEquals(r.lines.deductions[0], {
    description: "Specialist cleaning of the main hall carpets after the rehearsal dinner (including £308.40 VAT)",
    amount: "£1,850.40",
    line: "Specialist cleaning of the main hall carpets after the rehearsal dinner: £1,850.40, including £308.40 VAT",
  });
  assertEquals(
    r.lines.deductions[1].description,
    "Replacement of two damaged antique Venetian glass chandeliers in the Long Gallery (compensation, outside the scope of VAT)",
  );
  assertEquals(r.lines.deductions[2], { description: null, amount: null, line: null });
  assertEquals(r.cancelledOn, "3 November 2026");
  assertEquals(
    r.creditNote,
    "Credit note CN-2026-11-000913/BK-0048817 issued for £963,562.50, including £153,026.78 VAT against invoice SI-2026-07-004817/DEP-1.",
  );
  assertEquals(r.refundPaid, "The refund of £974,112.10 has not been paid yet.");
  assertEquals(
    r.retainedVat,
    "The £369,187.50 retained includes £59,008.93 VAT, accounted for when you paid. It remains due, so no credit note is issued for it.",
  );
  assertEquals(r.securityVat, "The £25,000.00 security deposit carried no VAT.");
  assertEquals(r.vatCorrection, null);
  assertEquals(r.salesTax, null);
});

test("a retained cent is taken from the refunded VAT, never worked out on its own", async () => {
  const r = await run({
    payments: [refundable(120, 20, 100, 20)],
    retention: { amount: 60.03, reason: "fee" },
  });
  assertEquals(r.totals.retainedVat, 1001);
  assertEquals(r.totals.refundedVat, 999);
  assertEquals(r.totals.refundable - r.totals.retention, 5997);
  assertEquals(r.totals.retainedVat + r.totals.refundedVat, 2000);
});

test("a percentage rounds half up on the pooled gross", async () => {
  const r = await run({
    payments: [refundable(100.01, 20, 83.34, 16.67)],
    retention: { percent: 50, reason: "fee" },
  });
  assertEquals(r.totals.retention, 5001);
  assertEquals(r.totals.retainedVat, 834);
  assertEquals(r.totals.refundedVat, 833);
});

test("nothing paid prints the intro and nothing else", async () => {
  for (const payments of [undefined, [], [security(0), refundable(0, 20, 0, 0)]]) {
    const r = await run({ payments });
    assertEquals(r.nothingPaid, true);
    assertEquals(Object.values(r.lines).filter((v) => v !== null && !Array.isArray(v)), []);
    assertEquals(r.lines.deductions.every((d) => d.line === null), true);
    for (const note of [r.refundPaid, r.creditNote, r.vatCorrection, r.retainedVat, r.securityVat, r.salesTax]) {
      assertEquals(note, null);
    }
    assertEquals(r.reference, "BK-1");
  }
});

test("a security deposit with no deductions comes back whole", async () => {
  const uk = await run({ payments: [security(500)] });
  assertEquals(uk.lines.refunded, "£500.00");
  assertEquals(uk.lines.retained, null);
  assertEquals(uk.securityVat, "The £500.00 security deposit carried no VAT.");
  assertEquals(uk.creditNote, null);
  const us = await run({ jurisdiction: "US", currency: "USD", payments: [security(500)] });
  assertEquals(us.securityVat, null);
  assertEquals(us.lines.refunded, "$500.00");
});

test("a deposit fully deducted refunds a printed zero, not a negative or a blank", async () => {
  const r = await run({ payments: [security(100)], deductions: [{ label: "Damage", amount: 100, isSupply: false }] });
  assertEquals(r.lines.refunded, "£0.00");
  assertEquals(r.lines.retained, "£100.00");
  assertEquals(r.refundPaid, null);
});

test("deductions: six fill the slots, seven or an uncovered one throw", async () => {
  const items = (n: number) => Array.from({ length: n }, (_, i) => ({ label: `Item ${i + 1}`, amount: 1, isSupply: false }));
  const six = await run({ payments: [security(100)], deductions: items(6) });
  assertEquals(six.lines.deductions.map((d) => d.amount), Array(6).fill("£1.00"));
  await assertRejects(() => run({ payments: [security(100)], deductions: items(7) }), Error, "order.deductions");
  await assertRejects(() => run({ deductions: items(1) }), Error, "no security deposit");
  await assertRejects(() => run({ payments: [security(5)], deductions: items(6) }), Error, "exceed");
});

test("deduction treatment: UK and DE must say, US says nothing", async () => {
  const pay = { payments: [security(100)] };
  await assertRejects(() => run({ ...pay, deductions: [{ label: "x", amount: 1 }] }), Error, "order.deductions[0].isSupply");
  await assertRejects(
    () => run({ ...pay, deductions: [{ label: "x", amount: 1, isSupply: true }] }),
    Error,
    "order.deductions[0].vatRate",
  );
  await assertRejects(() => run({ ...pay, deductions: [{ label: "x", amount: 0, isSupply: false }] }), Error, "amount");
  const us = await run({ jurisdiction: "US", currency: "USD", ...pay, deductions: [{ label: "Cleaning", amount: 12.5 }] });
  assertEquals(us.lines.deductions[0], { description: "Cleaning", amount: "$12.50", line: "Cleaning: $12.50" });
});

test("a non-refundable advance is kept whole with its own VAT", async () => {
  const r = await run({
    payments: [{ kind: "non-refundable-advance", gross: 120, bands: [{ rate: 20, net: 100, vat: 20 }] }],
  });
  assertEquals(r.totals.retainedVat, 2000);
  assertEquals(r.totals.refundedVat, 0);
  assertEquals(r.lines.refunded, "£0.00");
  assertEquals(r.lines.nonRefundable, "£120.00");
  assertEquals(r.creditNote, null);
  assertEquals(r.retainedVat !== null, true);
});

test("a refundable advance with no retention is refunded whole, with a credit note", async () => {
  const pay = [refundable(120, 20, 100, 20, "A")];
  const r = await run({ payments: pay });
  assertEquals(r.lines.refunded, "£120.00");
  assertEquals(r.lines.retention, null);
  assertEquals(r.creditNote, "A credit note for £120.00, including £20.00 VAT against invoice A, will follow.");
  assertEquals((await run({ payments: pay, creditNoteNumber: " CN1 " })).creditNote?.startsWith("Credit note CN1 issued"), true);
  assertEquals((await run({ payments: [refundable(120, 20, 100, 20)] })).creditNote?.includes("against"), false);
  const two = await run({ payments: [refundable(60, 20, 50, 10, "A"), refundable(60, 20, 50, 10, "B")] });
  assertEquals(two.creditNote?.includes("against invoices A and B"), true);
  const three = await run({
    payments: [refundable(60, 20, 50, 10, "A"), refundable(60, 20, 50, 10, "B"), refundable(60, 20, 50, 10, "C")],
  });
  assertEquals(three.creditNote?.includes("against invoices A, B and C"), true);
  // An advance with nothing refunded is not named.
  const zero = await run({ payments: [refundable(120, 20, 100, 20, "A"), refundable(0, 0, 0, 0, "Z")] });
  assertEquals(zero.creditNote?.endsWith(" against invoice A, will follow."), true);
});

test("retention bounds", async () => {
  const pay = [refundable(120, 20, 100, 20)];
  const all = await run({ payments: pay, retention: { percent: 100, reason: "r" } });
  assertEquals(all.totals.retainedVat, 2000);
  assertEquals(all.totals.refundedVat, 0);
  assertEquals(all.lines.refunded, "£0.00");
  // The band's own VAT is kept exactly, not re-derived from its rate: 99.99 + 20.01 is not 20%.
  const odd = await run({
    payments: [refundable(120, 20, 99.99, 20.01)],
    retention: { percent: 100, reason: "r" },
  });
  assertEquals(odd.totals.retainedVat, 2001);
  assertEquals((await run({ payments: pay, retention: { percent: 0, reason: "r" } })).lines.retention, null);
  await assertRejects(() => run({ payments: pay, retention: { percent: 100.5, reason: "r" } }), Error, "order.retention.percent");
  await assertRejects(() => run({ payments: pay, retention: { amount: 120.01, reason: "r" } }), Error, "order.retention.amount");
  await assertRejects(() => run({ retention: { amount: 1, reason: "r" } }), Error, "order.retention.amount");
  await assertRejects(() => run({ payments: pay, retention: { percent: 5, amount: 5, reason: "r" } }), Error, "exactly one");
  await assertRejects(() => run({ payments: pay, retention: { reason: "r" } }), Error, "exactly one");
  await assertRejects(() => run({ payments: pay, retention: { percent: 5, reason: " " } }), Error, "order.retention.reason");
});

test("a zero-rated advance has nothing to credit or to say about VAT", async () => {
  const r = await run({ payments: [{ kind: "refundable-advance", gross: 50 }, { kind: "non-refundable-advance", gross: 10 }] });
  assertEquals(r.lines.refunded, "£50.00");
  assertEquals(r.creditNote, null);
  assertEquals(r.retainedVat, null);
});

test("mixed rates: the spare minor unit goes to the largest remainder, the earlier band on a tie", async () => {
  // Tie: T = 5 over grosses 6 and 6 gives 2 and 2 with equal remainders, so the spare cent goes to band 1.
  // Keeping 3 puts its retained VAT at round(3 * 2000 / 12000) = 1; keeping 2 would give 0.
  const r = await run({
    payments: [{
      kind: "refundable-advance",
      gross: 0.12,
      bands: [{ rate: 20, net: 0.05, vat: 0.01 }, { rate: 0, net: 0.06, vat: 0 }],
    }],
    retention: { amount: 0.05, reason: "r" },
  });
  assertEquals(r.totals.retention, 5);
  assertEquals(r.totals.retainedVat, 1);
  assertEquals(r.totals.refundedVat, 0);
  // Unequal remainders: T = 4 over grosses 1, 7, 3 gives floors 0, 2, 1 and remainders 4, 6, 1.
  // The spare cent goes to the middle band, which keeps 3 and so retains 1 of its VAT; any other order gives 0.
  const s = await run({
    payments: [{
      kind: "refundable-advance",
      gross: 0.11,
      bands: [{ rate: 0, net: 0.01, vat: 0 }, { rate: 20, net: 0.06, vat: 0.01 }, { rate: 0, net: 0.03, vat: 0 }],
    }],
    retention: { amount: 0.04, reason: "r" },
  });
  assertEquals(s.totals.retention, 4);
  assertEquals(s.totals.retainedVat, 1);
  assertEquals(s.totals.refundedVat, 0);
});

test("Germany: VAT is corrected on repayment, the fee is consideration", async () => {
  const order = {
    jurisdiction: "DE",
    currency: "EUR",
    payments: [refundable(119, 19, 100, 19, "R-1")],
    retention: { percent: 50, reason: "Stornogebühr" },
  };
  const open = await run(order);
  assertEquals(
    open.vatCorrection,
    "The €9.50 VAT on the refunded amount is corrected when the refund is paid, not on cancellation.",
  );
  assertEquals(open.retainedVat, "The €59.50 retained is the cancellation fee agreed in your contract and includes €9.50 VAT.");
  assertEquals(open.creditNote?.startsWith("A credit note for €59.50"), true);
  const paid = await run({ ...order, refundPaidOn: "2026-12-01" });
  assertEquals(
    paid.vatCorrection,
    "The €9.50 VAT on the refunded amount is corrected in the period the refund was paid, 1 December 2026.",
  );
  assertEquals(paid.refundPaid, "The refund of €59.50 was paid on 1 December 2026.");
});

test("US: the state decides, and the component asks rather than guesses", async () => {
  const us = { jurisdiction: "US", currency: "USD", payments: [{ kind: "non-refundable-advance", gross: 200 }] };
  assertEquals(
    (await run({ ...us, usState: "California", usForfeitTaxable: true })).salesTax,
    "Under California rules, sales tax is due on the $200.00 retained. It is not included in the amounts above.",
  );
  assertEquals(
    (await run({ ...us, usState: "OR", usForfeitTaxable: false })).salesTax,
    "OR charges no sales tax on the $200.00 retained.",
  );
  await assertRejects(() => run({ ...us, usState: "CA" }), Error, "order.usForfeitTaxable");
  await assertRejects(() => run({ ...us, usForfeitTaxable: true }), Error, "order.usState");
  const none = await run({ ...us, payments: [refundable(100, 0, 100, 0)] });
  assertEquals(none.salesTax, null);
  await assertRejects(
    () => run({ ...us, payments: [refundable(106, 6, 100, 6)] }),
    Error,
    "order.payments[0].bands[0].vat",
  );
  assertEquals(none.lines.refunded, "$100.00");
});

test("bad input throws, naming the field", async () => {
  const bad = (order: Obj, field: string) => assertRejects(() => run(order), Error, "depositCancellationOutcome: ").then(
    () => assertRejects(() => run(order), Error, field),
  );
  await bad({ payments: [{ kind: "security", gross: -1 }] }, "order.payments[0].gross");
  await bad({ payments: [{ kind: "security", gross: 1.005 }] }, "order.payments[0].gross");
  await bad({ currency: "JPY", payments: [{ kind: "security", gross: 1.5 }] }, "order.payments[0].gross");
  await bad({ payments: [security(1), { kind: "refundable-advance", gross: 120, bands: [{ rate: 20, net: 100, vat: 10 }] }] }, "order.payments[1].bands");
  await bad({ payments: [{ kind: "security", gross: 12, bands: [{ rate: 20, net: 10, vat: 2 }] }] }, "order.payments[0].bands[0].vat");
  await bad({ payments: [{ kind: "gift", gross: 1 }] }, "order.payments[0].kind");
  await bad({ jurisdiction: "FR" }, "order.jurisdiction");
  await bad({ currency: "POUNDS" }, "order.currency");
  await bad({ reference: " " }, "order.reference");
  await bad({ cancelledOn: "03/11/2026" }, "order.cancelledOn");
  await bad({ cancelledOn: "2019-02-28" }, "order.cancelledOn");
  await bad({ cancelledOn: "2026-13-01" }, "order.cancelledOn");
  await bad({ refundPaidOn: "2026-99-99" }, "order.refundPaidOn");
  await bad({ refundPaidOn: "soon" }, "order.refundPaidOn");
  assertEquals((await run({ cancelledOn: "2019-03-01" })).cancelledOn, "1 March 2019");
  assertEquals((await run({ cancelledOn: "2019-02-28", jurisdiction: "DE" })).cancelledOn, "28 February 2019");
  assertEquals((await run({ currency: "JPY", payments: [security(1500)] })).lines.paid, "JP¥1,500");
});

// The node, built and published.
const ARRANGEMENTS = ["table", "stacked"] as const;

function documentOf(arrangement?: (typeof ARRANGEMENTS)[number]) {
  return template(
    <Document title="Cancellation">
      <DepositCancellationOutcome arrangement={arrangement} />
    </Document>,
  );
}

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { title?: string; text?: string; children?: DocumentNode[] };
    return [...(own.text ? [own.text] : []), ...texts(own.children ?? [])];
  });
}

/** Ids of every node that only exists under a condition. */
function conditional(value: unknown, found: string[] = []): string[] {
  if (Array.isArray(value)) value.forEach((item) => conditional(item, found));
  else if (value && typeof value === "object") {
    const node = value as { id?: string; when?: unknown };
    if (node.id && node.when) found.push(node.id);
    Object.values(value).forEach((item) => conditional(item, found));
  }
  return found;
}

const data = { order: preview };

test("every arrangement builds, they differ, and no prop is the table", async () => {
  const build = (arrangement?: (typeof ARRANGEMENTS)[number]) =>
    buildDocument(documentOf(arrangement), data, { dynamicMode: "placeholder" });
  const table = (await build("table")).nodes;
  const stacked = (await build("stacked")).nodes;
  assertEquals(JSON.stringify(table) === JSON.stringify(stacked), false);
  assertEquals(await build().then((built) => built.nodes), table);
  assertEquals(texts(stacked).includes("Refunded to you: £974,112.10"), true);
  assertEquals(texts(table).includes("Refunded to you"), true);
  assertEquals(texts(table).includes("£974,112.10"), true);
});

const publishOptions = {
  branchMode: "publish",
  deriverMode: "preserve",
  dynamicMode: "resolve",
  aiClient: { generateParagraph: () => "" },
} as const;

for (const arrangement of ARRANGEMENTS) {
  test(`${arrangement}: published, it is 19 branches and resolves to what a direct build prints`, async () => {
    const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof data, publishOptions);
    const ids = conditional(published.nodes);
    const notes = ["refund-paid", "credit-note", "vat-correction", "retained-vat", "security-vat", "sales-tax"];
    for (const note of notes) assertEquals(ids.includes(`deposit-cancellation-${note}`), true);
    assertEquals(ids.includes("deposit-cancellation-nothing-paid"), true);
    assertEquals(ids.includes("deposit-cancellation-amounts") || ids.includes("deposit-cancellation-paid"), true);
    for (const row of ["security", "refundable", "retention", "non-refundable", "retained", "deduction-1", "deduction-6"]) {
      assertEquals(ids.includes(`deposit-cancellation-${row}`), true);
    }
    // Counted exactly: one more branch than this would fail, one fewer would too.
    await buildDocument(documentOf(arrangement), createPublishData() as typeof data, { ...publishOptions, branchLimit: 19 });
    await assertRejects(
      () => buildDocument(documentOf(arrangement), createPublishData() as typeof data, { ...publishOptions, branchLimit: 18 }),
      Error,
      "branches",
    );

    const resolve = async (order: Obj) => {
      const d = { order };
      const state: RuntimeState = {
        ctx: { ...d },
        derived: {},
        dataProvider: new InMemoryDataProvider(d),
        aiClient: { generateParagraph: () => "" },
      };
      return texts((await resolveDocument(published, state, { derivers: registry })).nodes);
    };
    const full = await resolve(preview);
    const direct = texts((await buildDocument(documentOf(arrangement), data, { dynamicMode: "placeholder" })).nodes);
    assertEquals(full, direct);
    assertEquals(full.some((line) => line.includes("£974,112.10")), true);
    assertEquals(await resolve({ ...base }), [
      "Order BK-1 was cancelled on 3 November 2026.",
      "Nothing was paid in advance on this order, so nothing is refunded or retained.",
    ]);
  });
}
