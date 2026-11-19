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

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/renewal-and-price-change-notice.deriver.ts")
>("components/derivers/renewal-and-price-change-notice.deriver.ts");
const { RenewalAndPriceChangeNotice } = await loadRegistryModule<
  typeof import("../registry/components/renewal-and-price-change-notice/renewal-and-price-change-notice.node.tsx")
>("components/renewal-and-price-change-notice/renewal-and-price-change-notice.node.tsx");

type Obj = Record<string, unknown>;
const base: Obj = {
  jurisdiction: "US-CA",
  customerType: "consumer",
  contractStart: "2025-09-01",
  initialTermMonths: 12,
  renewalTermMonths: 12,
  renewalDate: "2026-09-01",
  noticePeriod: { value: 1, unit: "months" },
  sendDate: "2026-08-01",
  currency: "USD",
  currentAmount: 120,
  cancelInstructions: "Visit account settings",
};

interface Result {
  regime: string;
  heading: string;
  statement: string;
  renewal: string;
  trial: string | null;
  priceChange: string | null;
  previousPayment: string | null;
  cancelBy: string | null;
  consequence: string | null;
  warning: string | null;
  timing: "early" | "ok" | "late" | null;
  cancelByIso: string | null;
  windows: Array<{ anchor: string; from: string; to: string; status: string }>;
}

const run = async (over: Obj = {}) => await mod.run([{ ...base, ...over }], undefined as never) as Result;
const fails = (over: Obj, field: string) => assertRejects(() => run(over), Error, field);
const day = (iso: string, n: number) => new Date(Date.parse(iso) + n * 86_400_000).toISOString().slice(0, 10);

test("base case is California", async () => {
  const r = await run();
  assertEquals(r.regime, "us-ca-arl");
  assertEquals(r.heading, "Renewal notice");
  assertEquals(r.renewal, "September 1, 2026, for $120.00 every year");
  assertEquals(r.cancelBy, "On or before August 1, 2026");
});

test("month arithmetic clamps to the end of the month", async () => {
  const iso = async (renewalDate: string, noticePeriod: Obj) =>
    (await run({ renewalDate, noticePeriod, sendDate: "2026-08-01", jurisdiction: "other" })).cancelByIso;
  assertEquals(await iso("2027-03-31", { value: 1, unit: "months" }), "2027-02-28");
  assertEquals(await iso("2028-03-31", { value: 1, unit: "months" }), "2028-02-29");
  assertEquals(await iso("2027-05-31", { value: 3, unit: "months" }), "2027-02-28");
  assertEquals(await iso("2027-03-31", { value: 31, unit: "days" }), "2027-02-28");
  assertEquals(await iso("2027-03-31", { value: 0, unit: "days" }), "2027-03-30");
  assertEquals(await iso("2027-03-31", { value: 0, unit: "months" }), "2027-03-30");
});

test("California renewal window", async () => {
  const at = (n: number) => run({ sendDate: day("2026-09-01", n) });
  for (const n of [-45, -15]) assertEquals((await at(n)).timing, "ok");
  const early = await at(-46);
  assertEquals(early.timing, "early");
  assertEquals(early.warning?.includes("it is 1 day early"), true);
  const late = await at(-14);
  assertEquals([late.timing, late.cancelBy], ["late", null]);
  assertEquals(late.consequence?.startsWith("California law requires notice 15 to 45 days before the renewal on"), true);
  const short = await run({ initialTermMonths: 6 });
  assertEquals([short.timing, short.windows], [null, []]);
});

test("California only applies to consumer contracts from July 2025", async () => {
  assertEquals((await run({ contractStart: "2025-06-30", renewalDate: "2026-06-30", sendDate: "2026-06-01" })).regime, "contractual");
  const amended = await run({ contractStart: "2025-06-30", amendedOn: "2025-07-01", renewalDate: "2026-06-30", sendDate: "2026-06-01" });
  assertEquals(amended.regime, "us-ca-arl");
  assertEquals((await run({ customerType: "business" })).regime, "contractual");
});

test("California trial", async () => {
  const long = await run({ trialEnds: "2025-10-03", sendDate: "2025-09-20", renewalDate: "2026-09-01" });
  assertEquals(long.windows.map((w) => w.anchor), ["renewal", "trial-end"]);
  const short = await run({ trialEnds: "2025-10-02", sendDate: "2025-09-20" });
  assertEquals(short.windows.map((w) => w.anchor), ["renewal"]);
  assertEquals(short.trial, "October 2, 2025; from then $120.00 every year");
  assertEquals((await run({ trialEnds: "2025-09-01", sendDate: "2025-09-20" })).trial, null);
  const late = await run({ trialEnds: "2025-10-03", sendDate: "2025-10-01" });
  assertEquals(late.windows.find((w) => w.anchor === "trial-end")?.status, "late");
  assertEquals(late.consequence?.includes("3 to 21 days before your introductory period ends on October 3, 2025"), true);
});

test("California price change", async () => {
  const eff = "2026-09-01";
  const at = async (n: number) => (await run({ priceChange: { newAmount: 130, effectiveDate: eff }, sendDate: day(eff, n) }));
  const status = async (n: number) => (await at(n)).windows.find((w) => w.anchor === "price-change")?.status;
  assertEquals(await status(-7), "ok");
  assertEquals(await status(-6), "late");
  assertEquals(await status(-31), "early");
  // Renewal is early at -31 too; late wins over early.
  const mixed = await run({
    priceChange: { newAmount: 130, effectiveDate: "2026-08-10" },
    sendDate: "2026-08-05",
    renewalDate: "2026-09-30",
  });
  assertEquals(mixed.windows.map((w) => w.status), ["early", "late"]);
  assertEquals(mixed.timing, "late");
});

test("each window counts from its own anchor", async () => {
  const r = await run({ priceChange: { newAmount: 130, effectiveDate: "2026-12-01" }, sendDate: "2026-11-01", renewalDate: "2027-01-01" });
  const w = r.windows.find((x) => x.anchor === "price-change");
  assertEquals([w?.from, w?.to], ["2026-11-01", "2026-11-24"]);
});

const ny: Obj = {
  jurisdiction: "US-NY",
  customerType: "business",
  renewalDate: "2027-05-31",
  contractStart: "2024-05-31",
  noticePeriod: { value: 3, unit: "months" },
};

test("New York measures from the cancellation deadline", async () => {
  const ok = await run({ ...ny, sendDate: "2027-02-01" });
  assertEquals(ok.regime, "us-ny-5-903");
  assertEquals(ok.windows, [{ anchor: "cancel-by", from: "2027-01-29", to: "2027-02-13", status: "ok" }]);
  const late = await run({ ...ny, sendDate: "2027-02-14" });
  assertEquals(late.timing, "late");
  assertEquals(late.cancelBy, null);
  assertEquals(late.consequence?.includes("cannot be enforced"), true);
  assertEquals(late.statement, "Your contract was set to renew automatically on May 31, 2027.");
  assertEquals((await run({ ...ny, sendDate: "2027-01-28" })).timing, "early");
  assertEquals((await run({ ...ny, renewalTermMonths: 1, sendDate: "2027-02-01" })).regime, "contractual");
});

const de: Obj = {
  jurisdiction: "DE",
  contractStart: "2024-01-15",
  initialTermMonths: 24,
  renewalDate: "2026-01-15",
  noticePeriod: { value: 1, unit: "months" },
  sendDate: "2025-11-01",
  currency: "EUR",
};

test("German consumer contract past its minimum term", async () => {
  const r = await run({ ...de, sendDate: "2026-03-01" });
  assertEquals(r.regime, "de-consumer");
  assertEquals(r.cancelBy, "None. You can cancel at any time; the subscription then ends one month later.");
  assertEquals(r.cancelByIso, null);
  assertEquals(r.renewal.startsWith("Continues until you cancel, at"), true);
  assertEquals((await run({ ...de, sendDate: "2026-03-01", noticePeriod: { value: 14, unit: "days" } })).cancelBy?.endsWith("14 days later."), true);
});

test("German first term caps a long notice period", async () => {
  const r = await run({ ...de, noticePeriod: { value: 3, unit: "months" } });
  assertEquals(r.cancelByIso, "2025-12-15");
  assertEquals(r.warning?.includes("notice period of 3 months"), true);
  const ok = await run({ ...de, noticePeriod: { value: 14, unit: "days" } });
  assertEquals(ok.warning, null);
  assertEquals((await run({ ...de, contractStart: "2022-02-28", initialTermMonths: 12, renewalDate: "2026-01-15" })).regime, "contractual");
  assertEquals((await run({ ...de, customerType: "business" })).regime, "contractual");
});

test("UK subscription rules start at the commencement date", async () => {
  const uk: Obj = { jurisdiction: "UK", currency: "GBP", contractStart: "2026-11-19", renewalDate: "2027-11-19", sendDate: "2027-10-01" };
  assertEquals((await run(uk)).regime, "contractual");
  const started = await run({ ...uk, contractStart: "2027-01-01" });
  assertEquals(started.regime, "uk-dmcc");
  assertEquals(started.warning?.includes("previous payment"), true);
  assertEquals((await run({ ...uk, commencementDate: "2026-06-01" })).regime, "uk-dmcc");
  const prev = await run({ ...uk, contractStart: "2027-01-01", previousAmount: 100, currentAmount: 110 });
  assertEquals(prev.previousPayment, "£100.00. This renewal is £10.00 (10.0%) more");
  assertEquals(prev.warning, null);
});

test("price changes and percentages", async () => {
  const change = async (from: number, to: number, currency = "USD") =>
    (await run({ currency, currentAmount: from, priceChange: { newAmount: to, effectiveDate: "2026-09-01" } })).priceChange;
  assertEquals(await change(120, 135), "From September 1, 2026: $120.00 to $135.00, an increase of $15.00 (12.5%)");
  assertEquals((await change(135, 120))?.endsWith("a decrease of $15.00 (11.1%)"), true);
  assertEquals((await change(0, 10))?.endsWith("an increase of $10.00"), true);
  assertEquals((await change(10, 0))?.endsWith("(100.0%)"), true);
  assertEquals((await change(100000, 100000.05))?.endsWith("(less than 0.1%)"), true);
  assertEquals((await change(1, 1.005, "GBP"))?.endsWith("an increase of £0.01 (1.0%)"), true);
  assertEquals((await change(1000, 1100, "JPY"))?.endsWith("¥1,000 to ¥1,100, an increase of ¥100 (10.0%)"), true);
  const same = await run({ currentAmount: 120, priceChange: { newAmount: 120.001, effectiveDate: "2026-09-01" } });
  assertEquals([same.priceChange, same.heading], [null, "Renewal notice"]);
  const r = await run({ priceChange: { newAmount: 130, effectiveDate: "2026-09-01" } });
  assertEquals(r.renewal, "September 1, 2026, for $130.00 every year");
  const later = await run({ priceChange: { newAmount: 130, effectiveDate: "2026-09-02" }, sendDate: "2026-08-20" });
  assertEquals(later.renewal, "September 1, 2026, for $120.00 every year");
});

test("previous payment, zero amounts and frequency", async () => {
  const r = await run({ previousAmount: 100 });
  assertEquals(r.previousPayment, "$100.00. This renewal is $20.00 (20.0%) more");
  assertEquals((await run({ previousAmount: 120 })).previousPayment, "$120.00");
  assertEquals((await run({ currentAmount: 0 })).renewal.includes("$0.00"), true);
  for (const [n, want] of [[1, "every month"], [3, "every 3 months"], [12, "every year"], [24, "every 2 years"]] as const) {
    assertEquals((await run({ renewalTermMonths: n })).renewal.endsWith("for $120.00 " + want), true);
  }
});

test("a passed deadline drops the date and says the contract will renew", async () => {
  const r = await run({ jurisdiction: "other", sendDate: "2026-08-15" });
  assertEquals([r.regime, r.cancelBy, r.timing], ["contractual", null, null]);
  assertEquals(r.consequence?.startsWith("The last day to cancel before the renewal on 1 September 2026 was 1 August 2026"), true);
});

test("headings", async () => {
  const both = await run({ trialEnds: "2026-08-20", priceChange: { newAmount: 130, effectiveDate: "2026-09-01" } });
  assertEquals(both.heading, "Price change notice");
  assertEquals((await run({ trialEnds: "2026-08-20" })).heading, "Your introductory period is ending");
});

test("invalid input throws, naming the field", async () => {
  for (const field of ["jurisdiction", "customerType", "contractStart", "initialTermMonths", "renewalTermMonths", "renewalDate", "noticePeriod", "sendDate", "currency", "currentAmount", "cancelInstructions"]) {
    await fails({ [field]: undefined }, "renewal." + field);
  }
  await fails({ jurisdiction: "FR" }, "renewal.jurisdiction");
  await fails({ customerType: "consumer " }, "renewal.customerType");
  await fails({ initialTermMonths: 0 }, "renewal.initialTermMonths");
  await fails({ initialTermMonths: 1.5 }, "renewal.initialTermMonths");
  await fails({ noticePeriod: { value: -1, unit: "days" } }, "renewal.noticePeriod");
  await fails({ noticePeriod: { value: 1, unit: "weeks" } }, "renewal.noticePeriod");
  await fails({ currency: "XXXX" }, "renewal.currency");
  await fails({ currentAmount: -1 }, "renewal.currentAmount");
  await fails({ currentAmount: NaN }, "renewal.currentAmount");
  await fails({ renewalDate: "2027-02-30" }, "renewal.renewalDate");
  await fails({ renewalDate: "2025-08-31", sendDate: "2025-08-01" }, "renewal.contractStart");
  await fails({ sendDate: "2026-09-02" }, "renewal.sendDate");
  await fails({ cancelInstructions: "  " }, "renewal.cancelInstructions");
  await fails({ priceChange: { newAmount: -1, effectiveDate: "2026-09-01" } }, "renewal.priceChange.newAmount");
  await fails({ trialEnds: "soon" }, "renewal.trialEnds");
});

// The node, published and resolved by an engine that only has the bundle.
const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ renewalAndPriceChangeNotice: mod.run }),
);
const ARRANGEMENTS = ["stacked", "table"] as const;
const documentOf = (arrangement?: (typeof ARRANGEMENTS)[number]) =>
  template(
    <Document title="Notice">
      <RenewalAndPriceChangeNotice arrangement={arrangement} />
    </Document>,
  );

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[] };
    return [...(own.id ? [own.id] : []), ...ids(own.children ?? [])];
  });
}
function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { text?: string; children?: DocumentNode[] };
    return [...(own.text ? [own.text] : []), ...texts(own.children ?? [])];
  });
}

async function resolved(renewal: Obj, arrangement?: (typeof ARRANGEMENTS)[number]) {
  const data = { renewal: { ...base, ...renewal } };
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

for (const arrangement of ARRANGEMENTS) {
  test(`${arrangement}: a late New York notice keeps the consequence and drops the deadline`, async () => {
    const nodes = await resolved({ ...ny, sendDate: "2027-02-14" }, arrangement);
    const found = ids(nodes);
    assertEquals(found.includes("renewal-notice-consequence"), true);
    assertEquals(found.includes("renewal-notice-cancel-by"), false);
    assertEquals(texts(nodes).some((t) => t.includes("cannot be enforced")), true);
  });

  test(`${arrangement}: an on-time California notice keeps the deadline`, async () => {
    const nodes = await resolved({ sendDate: "2026-08-01" }, arrangement);
    assertEquals(ids(nodes).includes("renewal-notice-cancel-by"), true);
    assertEquals(texts(nodes).some((t) => t.includes("On or before August 1, 2026")), true);
  });
}

test("both arrangements print the same lines, placed differently", async () => {
  const preview = {
    ...ny,
    sendDate: "2027-02-01",
    currency: "USD",
    currentAmount: 100,
    previousAmount: 90,
    trialEnds: "2027-03-01",
    priceChange: { newAmount: 110, effectiveDate: "2027-05-31" },
  };
  const built = (a: (typeof ARRANGEMENTS)[number]) => resolved(preview, a);
  const [s, t] = [await built("stacked"), await built("table")];
  const lines = (n: DocumentNode[]) => ids(n).filter((i) => i.startsWith("renewal-notice-") && i !== "renewal-notice-details").sort();
  assertEquals(lines(s), lines(t));
  assertEquals(JSON.stringify(s) === JSON.stringify(t), false);
  assertEquals(ids(t).includes("renewal-notice-details"), true);
  assertEquals(JSON.stringify(await resolved(preview)), JSON.stringify(s));
});
