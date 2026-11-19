import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  dataRef,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, ParagraphNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";
import { COMPONENTS } from "../src/registry/catalog.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/invoice-billing-periods.deriver.ts")
>("components/derivers/invoice-billing-periods.deriver.ts");
const { InvoiceBillingPeriods } = await loadRegistryModule<
  typeof import("../registry/components/invoice-billing-periods/invoice-billing-periods.node.tsx")
>("components/invoice-billing-periods/invoice-billing-periods.node.tsx");
// Run through a bundle, so the test proves the function survives being shipped
// as source text, which is how a deriver fails.
const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ invoiceBillingPeriods: mod.run }),
);

type Obj = Record<string, unknown>;
const base: Obj = {
  anchor: "2026-01-01",
  interval: "month",
  invoiceDate: "2026-12-01",
  timeZone: "Europe/London",
  hasAdvanceItems: true,
  hasArrearsItems: false,
};

interface Period {
  start: string;
  end: string;
  part: boolean;
}
interface Result {
  statement: string;
  headerLine: string | null;
  header: { start: string; end: string };
  advance: Period | null;
  arrears: Period | null;
}

function stateOf(data: Obj): RuntimeState {
  return {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
}

// Every case runs both directly and through the shipped bundle, and the two
// must agree.
async function run(billing: Obj): Promise<Result> {
  const direct = mod.run([billing], stateOf({ billing })) as Result;
  const shipped = await registry.run(
    { name: "invoiceBillingPeriods", output: "periods", inputs: [dataRef("billing")] } as never,
    stateOf({ billing }),
  ) as Result;
  assertEquals(shipped, direct);
  return direct;
}

const statement = async (over: Obj) => (await run({ ...base, ...over })).statement;
const both = { hasAdvanceItems: true, hasArrearsItems: true };
const rejects = (over: Obj, message: string) => assertRejects(() => run({ ...base, ...over }), Error, message);

test("advance and arrears state two periods and a header that contains both", async () => {
  const r = await run({ ...base, ...both });
  assertEquals(r.statement, "Subscription 1–31 Dec 2026, billed in advance. Usage 1–30 Nov 2026, billed in arrears.");
  assertEquals(r.header, { start: "2026-11-01", end: "2026-12-31" });
  assertEquals(r.headerLine, "Invoicing period: 1 Nov – 31 Dec 2026.");
  assertEquals([r.advance?.part, r.arrears?.part], [false, false]);
});

test("one period states no header line", async () => {
  const a = await run({ ...base, hasAdvanceItems: false, hasArrearsItems: true });
  assertEquals(a.statement, "Usage 1–30 Nov 2026, billed in arrears.");
  assertEquals(a.header, { start: "2026-11-01", end: "2026-11-30" });
  assertEquals([a.headerLine, a.advance], [null, null]);

  const b = await run(base);
  assertEquals(b.statement, "Subscription 1–31 Dec 2026, billed in advance.");
  assertEquals([b.headerLine, b.arrears], [null, null]);
});

test("month-end anchors are clamped without being forgotten", async () => {
  const a = { anchor: "2026-01-31" };
  assertEquals(await statement({ ...a, invoiceDate: "2026-02-28" }), "Subscription 28 Feb – 30 Mar 2026, billed in advance.");
  assertEquals(await statement({ ...a, invoiceDate: "2026-03-28" }), "Subscription 28 Feb – 30 Mar 2026, billed in advance.");
  assertEquals(await statement({ ...a, invoiceDate: "2026-03-31" }), "Subscription 31 Mar – 29 Apr 2026, billed in advance.");
  assertEquals(
    await statement({ anchor: "2028-01-31", invoiceDate: "2028-02-29" }),
    "Subscription 29 Feb – 30 Mar 2028, billed in advance.",
  );
});

test("a 29 Feb annual anchor falls on 28 Feb in a common year", async () => {
  const a = { anchor: "2024-02-29", interval: "year" };
  assertEquals(
    await statement({ ...a, invoiceDate: "2026-02-28" }),
    "Subscription 28 Feb 2026 – 27 Feb 2027, billed in advance.",
  );
  assertEquals(
    await statement({ ...a, invoiceDate: "2027-02-28" }),
    "Subscription 28 Feb 2027 – 28 Feb 2028, billed in advance.",
  );
});

test("quarterly, fortnightly and year-crossing periods", async () => {
  assertEquals(
    await statement({ anchor: "2026-01-31", intervalCount: 3, invoiceDate: "2026-05-01" }),
    "Subscription 30 Apr – 30 Jul 2026, billed in advance.",
  );
  assertEquals(
    await statement({ anchor: "2026-11-02", interval: "week", intervalCount: 2, invoiceDate: "2026-11-30", ...both }),
    "Subscription 30 Nov – 13 Dec 2026, billed in advance. Usage 16–29 Nov 2026, billed in arrears.",
  );
  assertEquals(
    await statement({ anchor: "2026-12-15", invoiceDate: "2026-12-15" }),
    "Subscription 15 Dec 2026 – 14 Jan 2027, billed in advance.",
  );
});

test("part periods", async () => {
  const a = { anchor: "2026-12-01", signupDate: "2026-11-17" };
  const one = await run({ ...base, ...a, invoiceDate: "2026-11-17" });
  assertEquals(one.statement, "Subscription 17–30 Nov 2026 (part period), billed in advance.");
  assertEquals(one.advance, { start: "2026-11-17", end: "2026-11-30", part: true });

  const two = await run({ ...base, ...a, ...both, invoiceDate: "2026-12-01" });
  assertEquals(
    two.statement,
    "Subscription 1–31 Dec 2026, billed in advance. Usage 17–30 Nov 2026 (part period), billed in arrears.",
  );
  assertEquals(two.header, { start: "2026-11-17", end: "2026-12-31" });

  assertEquals(
    await statement({ ...a, signupDate: "2026-11-30", invoiceDate: "2026-11-30" }),
    "Subscription 30 Nov 2026 (part period), billed in advance.",
  );
  const onAnchor = await run({ ...base, anchor: "2026-12-01", signupDate: "2026-12-01", invoiceDate: "2026-12-01" });
  assertEquals([onAnchor.advance?.part, onAnchor.statement.includes("part")], [false, false]);
});

test("an anchor after the invoice is counted backwards", async () => {
  const r = await run({ ...base, anchor: "2026-12-01", invoiceDate: "2026-11-20" });
  assertEquals(r.statement, "Subscription 1–30 Nov 2026, billed in advance.");
  assertEquals(r.advance?.part, false);
});

test("an instant is a calendar date in the billing time zone", async () => {
  const instant = { invoiceDate: "2026-12-01T03:00:00Z" };
  assertEquals(await statement({ ...instant, timeZone: "America/New_York" }), "Subscription 1–30 Nov 2026, billed in advance.");
  assertEquals(await statement({ ...instant, timeZone: "Europe/London" }), "Subscription 1–31 Dec 2026, billed in advance.");
  assertEquals(
    await statement({ anchor: "2026-02-01T00:00:00Z", timeZone: "America/New_York", invoiceDate: "2026-02-15" }),
    "Subscription 31 Jan – 27 Feb 2026, billed in advance.",
  );
});

test("missing flags and counts read as the documented defaults", async () => {
  assertEquals(await run({ ...base, intervalCount: undefined }), await run({ ...base, intervalCount: 1 }));
  const arrearsOnly = await run({ ...base, hasAdvanceItems: undefined, hasArrearsItems: true });
  assertEquals(arrearsOnly.statement, "Usage 1–30 Nov 2026, billed in arrears.");
});

test("bad input throws an error naming the field", async () => {
  for (const intervalCount of [0, -1, 1.5]) await rejects({ intervalCount }, "billing.intervalCount");
  await rejects({ interval: "day" }, "billing.interval");
  await rejects({ timeZone: undefined }, "billing.timeZone");
  await rejects({ timeZone: "Mars/Olympus_Mons" }, "billing.timeZone");
  await rejects({ anchor: undefined }, "billing.anchor");
  await rejects({ anchor: "2026-02-30" }, "billing.anchor");
  await rejects({ invoiceDate: "2026-12-01T03:00:00" }, "billing.invoiceDate");
  await rejects({ invoiceDate: 1796000000000 }, "billing.invoiceDate");
  await rejects({ invoiceDate: new Date() }, "billing.invoiceDate");
  await rejects({ signupDate: "2026-12-02" }, "billing.signupDate");
  await rejects({ hasAdvanceItems: false, hasArrearsItems: false }, "billing.hasAdvanceItems");
  await rejects({ hasAdvanceItems: undefined, hasArrearsItems: undefined }, "billing.hasAdvanceItems");
  await rejects({ hasArrearsItems: true, signupDate: "2026-12-01" }, "no earlier period");
});

test("the preview data gives the documented output", async () => {
  const entry = COMPONENTS.find((c) => c.id === "invoice-billing-periods");
  const r = await run((entry!.previewData as { billing: Obj }).billing);
  assertEquals(
    r.statement,
    "Subscription 31 Dec 2026 – 30 Jan 2027, billed in advance. Usage 30 Nov – 30 Dec 2026, billed in arrears.",
  );
  assertEquals(r.headerLine, "Invoicing period: 30 Nov 2026 – 30 Jan 2027.");
});

// The node, published and resolved by an engine that only has the bundle.
function documentOf() {
  return template(
    <Document title="Invoice">
      <InvoiceBillingPeriods />
    </Document>,
  );
}

function texts(nodes: DocumentNode[]): Record<string, string> {
  const section = nodes[0];

  if (section?.kind !== "section") throw new Error("expected the section");

  return Object.fromEntries(
    [[section.id, section.title ?? ""], ...section.children.map((c) => [c.id, (c as ParagraphNode).text ?? ""])],
  );
}

function publish() {
  return buildDocument(documentOf(), createPublishData() as never, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
}

test("publishing carries one invocation and does not run the deriver", async () => {
  const built = await publish();
  const section = built.nodes[0] as unknown as { derivers?: unknown; children: ParagraphNode[] };

  assertEquals(section.derivers, [{
    name: "invoiceBillingPeriods",
    output: "invoiceBillingPeriods",
    inputs: [{ type: "ref", ref: { scope: "data", path: "billing" } }],
  }]);
  // The text is still a reference: nothing ran.
  assertEquals(section.children[0].text, "{{derived.invoiceBillingPeriods.statement}}");
});

for (
  const [label, over, header] of [
    ["two periods", both, "Invoicing period: 1 Nov – 31 Dec 2026."],
    ["one period", {}, undefined],
  ] as const
) {
  test(`publishing then resolving matches building directly (${label})`, async () => {
    const data = { billing: { ...base, ...over } };
    const resolved = await resolveDocument(await publish(), stateOf(data), { derivers: registry });
    const direct = await buildDocument(documentOf(), data, { dynamicMode: "placeholder" });
    const got = texts(resolved.nodes);

    assertEquals(got, texts(direct.nodes));
    assertEquals(got["billing-periods-header"], header);
    assertEquals(got["billing-periods"], "Billing period");
    assertEquals(got["billing-periods-statement"], (await run(data.billing)).statement);
  });
}

test("building locally with the preview data prints the same text", async () => {
  const entry = COMPONENTS.find((c) => c.id === "invoice-billing-periods")!;
  const built = await buildDocument(documentOf(), entry.previewData as never, { dynamicMode: "placeholder" });
  const got = texts(built.nodes);

  assertEquals(
    got["billing-periods-statement"],
    "Subscription 31 Dec 2026 – 30 Jan 2027, billed in advance. Usage 30 Nov – 30 Dec 2026, billed in arrears.",
  );
  assertEquals(got["billing-periods-header"], "Invoicing period: 30 Nov 2026 – 30 Jan 2027.");
});
