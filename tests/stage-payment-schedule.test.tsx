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

const nodeSource = "components/stage-payment-schedule/stage-payment-schedule.node.tsx";
const { StagePaymentSchedule } = await loadRegistryModule<
  typeof import("../registry/components/stage-payment-schedule/stage-payment-schedule.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/stage-payment-schedule.deriver.ts")
>("components/derivers/stage-payment-schedule.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ stagePaymentSchedule: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

// deno-lint-ignore no-explicit-any
type Obj = Record<string, any>;
interface Line {
  description: string;
  status: string;
  amount: string;
  vat: string;
  taxPoint: string;
  grossMinor: number;
  deductionMinor: number;
  netMinor: number;
  vatMinor: number | null;
}
interface Out {
  intro: string;
  taxPoints: boolean;
  taxNote: string;
  deposit: Line;
  stages: Line[];
  totals: Obj;
}
const run = (contract: Obj) => mod.run([contract], {} as RuntimeState) as Out;
const gross = (contract: Obj) => run(contract).stages.slice(0, contract.stages.length).map((l) => l.grossMinor);
const pct = (...p: number[]) => p.map((percent) => ({ percent, state: "pending", label: "S" }));
const sum = (S: number, stages: Obj[], more: Obj = {}) => ({ contractSumMinor: S, stages, ...more });
const st = (more: Obj) => ({ label: "S", state: "pending", ...more });
const throws = (contract: Obj, text: string) =>
  assertRejects(() => Promise.resolve().then(() => run(contract)), Error, text);

test("rounding: lost pennies go on the last percentage stage", () => {
  assertEquals(gross(sum(10000, pct(33.33, 33.33, 33.34))), [3333, 3333, 3334]);
  assertEquals(gross(sum(10001, pct(33.33, 33.33, 33.34))), [3333, 3333, 3335]);
  assertEquals(gross(sum(10004, [st({ amountMinor: 2501 }), st({ percent: 75 })])), [2501, 7503]);
  assertEquals(gross(sum(10000, [...pct(30, 30, 30), st({ balance: true })])), [3000, 3000, 3000, 1000]);
  assertEquals(gross(sum(10000, pct(100))), [10000]);
});

test("allocation errors", async () => {
  await throws(sum(10000, pct(30, 30, 30)), "£10.00 short");
  await throws(sum(10000, [...pct(30), st({ balance: true }), ...pct(10)]), "last stage only");
  await throws(sum(10000, [...pct(100), st({ balance: true })]), "no balance");
  await throws(sum(10000, pct(60, 50)), "more than 100");
  await throws(sum(10000, [st({ amountMinor: 6000 }), st({ amountMinor: 5000 })]), "fixed amounts");
  await throws(sum(10000, [st({ amountMinor: 4000 }), st({ percent: 70 })]), "more than the contract sum");
  await throws(sum(10000, [st({ amountMinor: 2500 }), st({ percent: 37.5 }), st({ amountMinor: 5000 })]), "more than the contract sum");
  await throws(sum(10001, [st({ amountMinor: 2501 }), st({ percent: 75 })]), "contract sum");
});

test("a percentage stage that floors to nothing still states the excess", async () => {
  await throws(sum(100, [st({ amountMinor: 100 }), st({ percent: 0.5 })]), "£0.01 more than the contract sum");
});

test("status strings, intro and tax note", () => {
  const o = run(sum(10000, [
    st({ percent: 30, state: "paid", paidDate: "2026-03-10" }),
    st({ percent: 30, state: "reached" }),
    st({ percent: 40, state: "invoiced", invoiceDate: "2026-08-07" }),
  ], { deposit: { amountMinor: 1500, receivedDate: "2026-02-16", application: "first" } }));
  assertEquals([o.deposit.status, o.stages[1].status, o.stages[2].status], ["Received 16 Feb 2026", "Reached", "Invoiced 7 Aug 2026"]);
  assertEquals(o.intro, "The contract sum of £100.00 is paid in 3 stages, after a deposit of £15.00.");
  assertEquals(run(sum(10000, pct(100))).intro, "The contract sum of £100.00 is paid in one stage.");
  const vat = (deposit?: Obj) => run(sum(10000, halves, { vat: uk, deposit })).taxNote;
  assertStringIncludes(vat(), "Reaching a stage is not a tax point");
  assertEquals(vat().includes("VAT on the deposit"), false);
  assertStringIncludes(vat({ amountMinor: 2000, receivedDate: "2026-01-15", application: "first" }), "VAT on the deposit");
});

test("eight stages print, nine throw, none throw", async () => {
  const eight = sum(8000, Array.from({ length: 8 }, () => st({ amountMinor: 1000 })));
  assertEquals(run(eight).stages.filter((l) => l.description).length, 8);
  await throws(sum(9000, Array.from({ length: 9 }, () => st({ amountMinor: 1000 }))), "more than 8");
  await throws(sum(10000, []), "contract.stages");
  await throws({ contractSumMinor: 10000 }, "contract.stages");
});

test("stage and sum validation names the field", async () => {
  await throws(sum(10000, [st({ percent: 50, amountMinor: 5000 })]), "contract.stages[0]");
  await throws(sum(10000, [st({})]), "exactly one");
  for (const percent of [12.345, 0, -5, 101]) await throws(sum(10000, [st({ percent })]), "contract.stages[0].percent");
  for (const S of [0, -1, 1.5]) await throws(sum(S, pct(100)), "contract.contractSumMinor");
  await throws(sum(10000, [st({ percent: 100, state: "done" })]), "contract.stages[0].state");
  await throws(sum(10000, [st({ percent: 100, reachedDate: "1 May" })]), "contract.stages[0].reachedDate");
  await throws(sum(10000, [st({ percent: 100, label: " " })]), "label");
});

const tri = (deposit: Obj) => sum(10000, pct(10, 40, 50), { deposit: { receivedDate: "2026-01-15", ...deposit } });
const split = (contract: Obj) => {
  const o = run(contract);
  const rows = o.stages.slice(0, contract.stages.length);
  return [rows.map((l) => l.deductionMinor), rows.map((l) => l.netMinor)];
};

test("deposit netting", () => {
  assertEquals(split(tri({ amountMinor: 1500, application: "first" })), [[1000, 500, 0], [0, 3500, 5000]]);
  assertEquals(run(tri({ amountMinor: 1500, application: "first" })).stages[0].description, "S (£10.00 less £10.00 of the deposit)");
  assertEquals(split(tri({ amountMinor: 6000, application: "final" })), [[0, 1000, 5000], [1000, 3000, 0]]);
  assertEquals(split(tri({ amountMinor: 1000, application: "proRata" })), [[100, 400, 500], [900, 3600, 4500]]);
  assertEquals(
    split(sum(10001, [st({ percent: 33.33 }), st({ percent: 33.33 }), st({ balance: true })], {
      deposit: { amountMinor: 1000, receivedDate: "2026-01-15", application: "proRata" },
    })),
    [[333, 333, 334], [3000, 3000, 3001]],
  );
  assertEquals(
    split(sum(10000, [9990, 5, 5].map((amountMinor) => st({ amountMinor })), {
      deposit: { amountMinor: 9999, receivedDate: "2026-01-15", application: "proRata" },
    })),
    [[9989, 5, 5], [1, 0, 0]],
  );
});

test("a deposit equal to the contract sum, and no deposit", () => {
  const o = run(tri({ amountMinor: 10000, application: "first" }));
  assertEquals(o.stages.slice(0, 3).map((l) => l.netMinor), [0, 0, 0]);
  assertEquals([o.totals.paidMinor, o.totals.invoicedMinor, o.totals.notInvoicedMinor], [10000, 0, 0]);
  const zero = run(sum(10000, pct(10, 40, 50), { deposit: { amountMinor: 0 } }));
  assertEquals(zero, run(sum(10000, pct(10, 40, 50))));
  assertEquals(zero.deposit.description, "");
});

test("deposit validation", async () => {
  await throws(tri({ amountMinor: 10001, application: "first" }), "deposit");
  await throws(tri({ amountMinor: -1, application: "first" }), "deposit");
  await throws(tri({ amountMinor: 1.5, application: "first" }), "deposit");
  await throws(tri({ amountMinor: 100, application: "first", receivedDate: undefined }), "receivedDate");
  await throws(tri({ amountMinor: 100, application: "last" }), "application");
  await throws(tri({ amountMinor: 100 }), "application");
});

const uk = { jurisdiction: "UK", ratePercent: 20 };
const halves = [
  st({ percent: 50, state: "invoiced", invoiceDate: "2026-03-10" }),
  st({ percent: 50, state: "reached", reachedDate: "2026-03-01" }),
];

test("tax points under regulation 93", () => {
  const o = run(sum(10000, halves, { vat: uk, deposit: { amountMinor: 2000, receivedDate: "2026-01-15", application: "first" } }));
  assertEquals([o.deposit.vat, o.deposit.taxPoint], ["£4.00", "15 Jan 2026"]);
  assertEquals([o.stages[0].amount, o.stages[0].vat, o.stages[0].taxPoint], ["£30.00", "£6.00", "10 Mar 2026"]);
  assertEquals([o.stages[1].vat, o.stages[1].taxPoint, o.stages[1].status], ["Not yet due", "Not yet", "Reached 1 Mar 2026"]);
  const t = o.totals;
  assertEquals([t.paidMinor, t.paidVatMinor, t.invoicedMinor, t.invoicedVatMinor, t.notInvoicedMinor, t.notInvoicedVat], [2000, 400, 3000, 600, 5000, "Not yet due"]);
  assertStringIncludes(o.taxNote, "VAT on the deposit");
  assertEquals(o.taxPoints, true);
});

test("a paid stage's tax point is the earlier of invoice and payment", () => {
  const paid = (more: Obj) => run(sum(10000, [st({ percent: 100, state: "paid", paidDate: "2026-05-13", ...more })], { vat: uk })).stages[0];
  assertEquals(paid({ invoiceDate: "2026-05-15" }).taxPoint, "13 May 2026");
  assertEquals(paid({ invoiceDate: "2026-05-01" }).taxPoint, "1 May 2026");
  assertEquals(paid({}).taxPoint, "13 May 2026");
});

test("missing dates throw", async () => {
  await throws(sum(10000, [st({ percent: 100, state: "invoiced" })]), "invoiceDate");
  await throws(sum(10000, [st({ percent: 100, state: "paid" })]), "paidDate");
});

test("VAT rounds half up once per line, and zero rate keeps the columns", async () => {
  const vat = (ratePercent: number, S: number) => run(sum(S, [st({ percent: 100, state: "invoiced", invoiceDate: "2026-05-01" })], { vat: { jurisdiction: "UK", ratePercent } })).stages[0];
  assertEquals(vat(5, 12345).vatMinor, 617);
  assertEquals(vat(5, 12350).vatMinor, 618);
  assertEquals([vat(0, 10000).vat, run(sum(10000, pct(100), { vat: { jurisdiction: "UK", ratePercent: 0 } })).taxPoints], ["£0.00", true]);
  for (const ratePercent of [-1, 101, 17.555]) await throws(sum(10000, pct(100), { vat: { jurisdiction: "UK", ratePercent } }), "ratePercent");
  await throws(sum(10000, pct(100), { vat: { jurisdiction: "DE", ratePercent: 19 } }), "jurisdiction");
});

test("no VAT: plain table, no note, blank VAT strings", async () => {
  const o = run(sum(10000, pct(100)));
  assertEquals([o.taxPoints, o.taxNote, o.stages[0].vat, o.stages[0].taxPoint, o.totals.notInvoicedVat], [false, "", "", "", ""]);
  const ids = await ids_of(sum(10000, pct(100)));
  assertEquals(ids.some((i) => i.includes("-vat-")) || ids.includes("stage-payment-schedule-tax-note"), false);
  assertEquals(ids.includes("stage-payment-schedule-table"), true);
});

test("single payment sits outside regulation 93", async () => {
  const contract = sum(10000, pct(100), { singlePayment: true, vat: uk, deposit: { amountMinor: 2000, receivedDate: "2026-01-15", application: "first" } });
  const o = run(contract);
  assertEquals(o.taxPoints, false);
  assertStringIncludes(o.taxNote, "VAT of £16.00");
  assertStringIncludes(o.taxNote, "on the £80.00 left after the deposit");
  assertStringIncludes(o.taxNote, "received on 15 Jan 2026");
  assertStringIncludes(o.intro, "a single sum on completion, after a deposit of £20.00");
  const ids = await ids_of(contract);
  assertEquals(ids.includes("stage-payment-schedule-table") && ids.includes("stage-payment-schedule-tax-note"), true);
  assertEquals(run(sum(10000, pct(100), { singlePayment: true })).taxNote, "");
  await throws(sum(10000, pct(50, 50), { singlePayment: true }), "exactly one stage");
});

test("currency and ordering", async () => {
  assertEquals(run(sum(10000, pct(100))).totals.contractSum, "£100.00");
  assertEquals(run({ currency: "JPY", ...sum(1000000, pct(100)) }).totals.contractSum, new Intl.NumberFormat("en-GB", { style: "currency", currency: "JPY" }).format(1000000));
  await throws({ currency: "XXX1", ...sum(10000, pct(100)) }, "contract.currency");
  const o = run(sum(10000, [st({ percent: 50, state: "pending" }), st({ percent: 50, state: "paid", paidDate: "2026-05-13" })]));
  assertEquals([o.totals.paidMinor, o.totals.notInvoicedMinor, o.stages[0].status, o.stages[1].status], [5000, 5000, "Not yet reached", "Paid 13 May 2026"]);
});

const preview = (COMPONENTS.find((c) => c.id === "stage-payment-schedule")!.previewData as { contract: Obj }).contract;

test("the catalog's preview data", () => {
  const o = run(preview);
  assertEquals(o.stages.map((l) => l.grossMinor), [350000000, 609406254, 853168755, 731287504, 731287504, 609406254, 487525003, 503168759]);
  assertEquals([o.stages[0].deductionMinor, o.stages[0].netMinor, o.stages[1].deductionMinor, o.stages[1].netMinor], [350000000, 0, 137525003, 471881251]);
  assertEquals([o.deposit.vatMinor, o.stages[0].vatMinor, o.stages[1].vatMinor, o.stages[2].vatMinor, o.stages[3].vat], [97505001, 0, 94376250, 170633751, "Not yet due"]);
  assertEquals(o.stages[1].taxPoint, "13 May 2026");
  const t = o.totals;
  assertEquals([t.paidMinor, t.paidVatMinor, t.invoicedMinor, t.invoicedVatMinor, t.notInvoicedMinor], [959406254, 191881251, 853168755, 170633751, 3062675024]);
  assertEquals([t.paid, t.notInvoiced], ["£9,594,062.54", "£30,626,750.24"]);
  assertEquals(o.stages.every((l) => l.description), true);
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Contract">
      <Section id="body" title="Body">
        <StagePaymentSchedule />
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

const built = (contract: Obj) =>
  buildDocument(templateOf(), { contract }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);

async function ids_of(contract: Obj): Promise<string[]> {
  const body = (await built(contract)).nodes[0] as { children: DocumentNode[] };
  return collect(body.children).map((n) => n.id).filter((id) => !id.startsWith("table-cell"));
}

test("the preview prints every slot in the VAT table", async () => {
  const ids = await ids_of(preview);
  assertEquals(ids.includes("stage-payment-schedule-vat-deposit"), true);
  for (let i = 1; i <= 8; i++) assertEquals(ids.includes(`stage-payment-schedule-vat-stage-${i}`), true);
  assertEquals(ids.includes("stage-payment-schedule-tax-note"), true);
  assertEquals(ids.includes("stage-payment-schedule-table"), false);
});

test("a publish build carries every arm and one deriver call", async () => {
  // The transform counts the table ternary once; the publish build carries both arms.
  assertEquals(branches, 20);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 21);

  const section = all.find((n) => n.id === "stage-payment-schedule") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  const d = JSON.parse(JSON.stringify(section.derivers[0]));
  assertEquals(d.name, "stagePaymentSchedule");
  assertEquals(d.inputs, [{ type: "ref", ref: { scope: "data", path: "contract" } }]);

  const cases: Obj[] = [
    preview,
    sum(10000, pct(100)),
    sum(10000, pct(100), { singlePayment: true, vat: uk, deposit: { amountMinor: 2000, receivedDate: "2026-01-15", application: "first" } }),
    sum(10000, halves, { vat: uk }),
  ];
  for (const contract of cases) {
    const direct = await built(contract);
    const resolved = await resolveDocument(published, state({ contract }), { derivers });
    // A resolved tree keeps the conditions it was published with and numbers its
    // cells differently, so compare what a reader sees: the ids and the text.
    const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.kind, n.id, (n as { text?: string }).text]);
    assertEquals(seen(resolved.nodes).filter((r) => !String(r[1]).startsWith("table-cell")), seen(direct.nodes).filter((r) => !String(r[1]).startsWith("table-cell")));
  }
});
