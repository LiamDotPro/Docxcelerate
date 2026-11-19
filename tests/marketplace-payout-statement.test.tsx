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
  typeof import("../registry/components/derivers/marketplace-payout-statement.deriver.ts")
>("components/derivers/marketplace-payout-statement.deriver.ts");
const { MarketplacePayoutStatement } = await loadRegistryModule<
  typeof import("../registry/components/marketplace-payout-statement/marketplace-payout-statement.node.tsx")
>("components/marketplace-payout-statement/marketplace-payout-statement.node.tsx");
// Through a bundle, so the test proves the function survives being shipped as source text.
const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ marketplacePayoutStatement: mod.run }),
);

type Obj = Record<string, unknown>;
interface Row {
  label: string;
  amount: string;
  paidIn: string;
  takenOut: string;
  minor: number;
}
interface Result {
  rows: Record<string, Row>;
  outcome: string;
  outcomeAmount: string;
  payTo: string;
  reverseChargeNote: string;
  closingMinor: number;
}

const state = (data: Obj): RuntimeState => ({
  ctx: { ...data },
  derived: {},
  dataProvider: new InMemoryDataProvider(data),
  aiClient: { generateParagraph: () => "" },
});

async function run(payout: Obj): Promise<Result> {
  return await registry.run(
    { name: "marketplacePayoutStatement", output: "s", inputs: [dataRef("payout")] } as never,
    state({ payout }),
  ) as Result;
}

const base: Obj = { currency: "GBP", openingBalance: 0 };
const printed = (r: Result) =>
  Object.entries(r.rows).filter(([, row]) => row.amount !== "").map(([key]) => key);

const worked: Obj = {
  ...base,
  orders: [{ id: "A", itemTotal: 10000, shipping: 499, facilitatorTax: 2100 }],
  refunds: [{ orderId: "A", amount: 2500, commissionRefunded: 375, adminFee: 75 }],
  fees: [
    { label: "Referral", amount: 1575, vatPercent: 20 },
    { label: "Storage", amount: 333, vatPercent: 20 },
    { label: "Ads", amount: 1000, reverseCharge: true },
  ],
  chargebacks: [{ orderId: "A", amount: 1999, disputeFee: 1500, outcome: "won" }],
  reserve: { held: 1000, released: 400 },
  accountLast4: "4417",
};

test("worked small case", async () => {
  const r = await run(worked);
  const amounts = Object.fromEntries(Object.entries(r.rows).map(([k, v]) => [k, v.amount]));
  assertEquals(amounts, {
    broughtForward: "",
    sales: "+£104.99",
    taxCollected: "+£21.00",
    taxRemitted: "-£21.00",
    refunds: "-£25.00",
    commissionReturned: "+£3.75",
    refundAdminFees: "-£0.75",
    fees: "-£29.08",
    feeVat: "-£3.82",
    chargebacks: "-£19.99",
    disputeFees: "-£15.00",
    disputesWon: "+£19.99",
    reserveHeld: "-£10.00",
    reserveReleased: "+£4.00",
    closing: "£29.09",
  });
  assertEquals(r.rows.refunds.takenOut, "£25.00");
  assertEquals(r.rows.refunds.paidIn, "");
  assertEquals(r.outcome, "payout");
  assertEquals(r.payTo, "your account ending 4417");
  assertEquals(
    r.reverseChargeNote,
    "No VAT has been charged on fees of £10.00, which are supplied under the reverse charge. Account for the VAT on them in your own VAT return.",
  );
});

test("the preview data ties to the platform's figure and prints every row", async () => {
  const { COMPONENTS } = await import("../src/registry/catalog.ts");
  const entry = COMPONENTS.find((c) => c.id === "marketplace-payout-statement")!;
  const r = await run((entry.previewData as { payout: Obj }).payout);
  assertEquals(printed(r).length, 15);
  assertEquals(r.rows.closing.amount, "£190,392.59");
  assertEquals(r.outcome, "payout");
  assertEquals(r.reverseChargeNote.includes("£8,450.00"), true);
});

test("VAT rounds half away from zero, once per fee line", async () => {
  assertEquals((await run({ ...base, fees: [{ label: "a", amount: 1234, vatPercent: 8.1 }] })).rows.feeVat.minor, -100);
  assertEquals((await run({ ...base, fees: [{ label: "a", amount: -333, vatPercent: 20 }] })).rows.feeVat.minor, 67);
});

test("zero: only sales and closing print, and nothing is owed", async () => {
  const r = await run(base);
  assertEquals(printed(r), ["sales", "closing"]);
  assertEquals(r.rows.sales.amount, "£0.00");
  assertEquals(r.rows.sales.paidIn, "£0.00");
  assertEquals(r.rows.closing.paidIn, "£0.00");
  assertEquals(r.outcome, "nil");
  assertEquals(r.reverseChargeNote, "");
});

test("a negative balance is carried, or debited, never paid", async () => {
  const r = await run({ ...base, openingBalance: -5000 });
  assertEquals(r.rows.broughtForward.amount, "-£50.00");
  assertEquals(r.rows.closing.amount, "-£50.00");
  assertEquals(r.rows.closing.takenOut, "£50.00");
  assertEquals(r.outcome, "carry");
  assertEquals(r.outcomeAmount, "£50.00");
  assertEquals((await run({ ...base, openingBalance: -5000, negativeBalanceMethod: "debit" })).outcome, "debit");
});

test("a positive opening balance is added", async () => {
  const r = await run({ ...base, openingBalance: 1200 });
  assertEquals(r.rows.broughtForward.amount, "+£12.00");
  assertEquals(r.closingMinor, 1200);
});

test("chargebacks raised earlier only count when won", async () => {
  const cb = (outcome: string) => ({ chargebacks: [{ orderId: "A", amount: 900, disputeFee: 100, outcome, raisedEarlier: true }] });
  const won = await run({ ...base, ...cb("won") });
  assertEquals(printed(won), ["sales", "disputesWon", "closing"]);
  assertEquals(won.rows.disputesWon.minor, 900);
  for (const outcome of ["lost", "open"]) assertEquals(printed(await run({ ...base, ...cb(outcome) })), ["sales", "closing"]);
});

test("only the last four digits of an account are used", async () => {
  const pay = (accountLast4?: string) => run({ ...base, openingBalance: 100, accountLast4 });
  assertEquals((await pay("GB29 NWBK 6016 1331 9268 19")).payTo, "your account ending 6819");
  assertEquals((await pay("12")).payTo, "your bank account");
  assertEquals((await pay()).payTo, "your bank account");
});

test("closingBalance must tie", async () => {
  await run({ ...base, openingBalance: 100, closingBalance: 100 });
  await assertRejects(
    () => run({ ...base, openingBalance: 100, closingBalance: 101 }),
    Error,
    "payout.closingBalance is 101 but the rows sum to 100",
  );
});

test("bad input throws, naming the field", async () => {
  const bad: Array<[Obj, string]> = [
    [{ ...base, currency: undefined }, "payout.currency"],
    [{ ...base, currency: "XXX1" }, "payout.currency"],
    [{ ...base, openingBalance: undefined }, "payout.openingBalance"],
    [{ ...base, orders: [{ id: "A", itemTotal: 10.5 }] }, "payout.orders[0].itemTotal"],
    [{ ...base, refunds: [{ orderId: "A", amount: 1, commissionRefunded: 10, adminFee: 11 }] }, "payout.refunds[0].adminFee"],
    [{ ...base, refunds: [{ orderId: "A", amount: 1, commissionRefunded: 10 }] }, "payout.refunds[0].adminFee"],
    [{ ...base, fees: [{ label: "a", amount: 1 }, { label: "b", amount: 1, vatPercent: 120 }] }, "payout.fees[1].vatPercent"],
    [{ ...base, fees: [{ label: " ", amount: 1 }] }, "payout.fees[0].label"],
    [{ ...base, chargebacks: [{ orderId: "A", amount: 1, outcome: "pending" }] }, "payout.chargebacks[0].outcome"],
    [{ ...base, negativeBalanceMethod: "invoice" }, "payout.negativeBalanceMethod"],
  ];
  for (const [payout, field] of bad) await assertRejects(() => run(payout), Error, field);
});

test("a zero-decimal currency has no minor unit", async () => {
  const r = await run({ currency: "JPY", openingBalance: 0, orders: [{ id: "A", itemTotal: 12000 }] });
  assertEquals(r.rows.sales.amount, "+¥12,000");
});

test("one order or five hundred still make at most fifteen rows, summed exactly", async () => {
  const orders = (n: number) => Array.from({ length: n }, (_, i) => ({ id: "o" + i, itemTotal: 1999, shipping: 1 }));
  const one = await run({ ...base, orders: orders(1) });
  const many = await run({ ...base, orders: orders(500) });
  assertEquals(one.rows.sales.minor, 2000);
  assertEquals(many.rows.sales.minor, 1_000_000);
  assertEquals(printed(many).length <= 15, true);
});

// The node, published and resolved by an engine that only has the bundle.
const documentOf = (arrangement?: "signed" | "columns") =>
  template(
    <Document title="Statement">
      <MarketplacePayoutStatement arrangement={arrangement} />
    </Document>,
  );

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[] };
    // Cell ids are counted as they are built, so only the authored ids are stable.
    return [...(own.id && /^(payout|marketplace)-/.test(own.id) ? [own.id] : []), ...ids(own.children ?? [])];
  });
}

const publish = (arrangement?: "signed" | "columns") =>
  buildDocument(documentOf(arrangement), createPublishData() as { payout: Obj }, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });

const ROWS = [
  "brought-forward", "sales", "tax-collected", "tax-remitted", "refunds", "commission-returned",
  "refund-admin-fees", "fees", "fee-vat", "chargebacks", "dispute-fees", "disputes-won", "reserve-held",
  "reserve-released", "closing",
].map((id) => "payout-" + id);

for (const arrangement of ["signed", "columns"] as const) {
  test(`${arrangement}: every row, the note and each outcome are published as arms`, async () => {
    const published = await publish(arrangement);
    const tree = JSON.stringify(published.nodes);
    for (const id of [...ROWS, "payout-reverse-charge", "payout-paid", "payout-nil", "payout-carry", "payout-debit"]) {
      assertEquals(tree.includes(`"${id}"`), true);
    }
  });

  test(`${arrangement}: resolving the published tree matches building directly`, async () => {
    const published = await publish(arrangement);
    for (const payout of [worked, base, { ...base, openingBalance: -5000 }]) {
      const data = { payout };
      const resolved = await resolveDocument(published, state(data), { derivers: registry });
      const direct = await buildDocument(documentOf(arrangement), data, { dynamicMode: "placeholder" });
      assertEquals(ids(resolved.nodes), ids(direct.nodes));
    }
  });
}

test("the worked case resolves to its rows and the one paid outcome; zero to nil", async () => {
  const published = await publish();
  const resolve = async (payout: Obj) =>
    ids((await resolveDocument(published, state({ payout }), { derivers: registry })).nodes);
  const w = await resolve(worked);
  assertEquals(w.filter((id) => ROWS.includes(id)), ROWS.filter((id) => id !== "payout-brought-forward"));
  assertEquals(w.filter((id) => /payout-(paid|nil|carry|debit)$/.test(id)), ["payout-paid"]);
  assertEquals(w.includes("payout-reverse-charge"), true);
  const z = await resolve(base);
  assertEquals(z.filter((id) => ROWS.includes(id)), ["payout-sales", "payout-closing"]);
  assertEquals(z.filter((id) => /payout-(paid|nil|carry|debit)$/.test(id)), ["payout-nil"]);
  assertEquals(z.includes("payout-reverse-charge"), false);
});
