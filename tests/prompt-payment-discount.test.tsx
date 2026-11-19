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
import { assertEquals } from "./assert.ts";
import assert from "node:assert/strict";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/prompt-payment-discount/prompt-payment-discount.node.tsx";
const { PromptPaymentDiscount } = await loadRegistryModule<
  typeof import("../registry/components/prompt-payment-discount/prompt-payment-discount.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/prompt-payment-discount.deriver.ts")
>("components/derivers/prompt-payment-discount.deriver.ts");
type Result = import("../registry/components/derivers/prompt-payment-discount.deriver.ts").PromptPaymentDiscountResult;
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ promptPaymentDiscount: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
const run = (invoice: Obj): Result => mod.run([invoice], {} as RuntimeState) as Result;
// Intl puts no-break spaces in currency output; compare with plain ones.
const plain = (s: string) => s.replace(/[  ]/g, " ");
const cells = (r: Result) => r.rows.filter((x) => x.label).map((x) => [x.label, x.first, x.second, x.third].map(plain));
const throws = (invoice: Obj, text: string) =>
  assert.throws(() => run(invoice), (e: Error) => e.message.includes(text), text);

const preview = (COMPONENTS.find((c) => c.id === "prompt-payment-discount")!.previewData as { invoice: Obj }).invoice;
const band = (treatment: string, rate: number | null, net: number, vat: number) => ({ treatment, rate, net, vat });
const uk = (more: Obj = {}, disc: Obj = { percent: 2, days: 14 }): Obj => ({
  issueDate: "2026-11-30",
  currency: "GBP",
  jurisdiction: "UK",
  bands: [band("standard", 20, 100, 20)],
  promptPaymentDiscount: disc,
  ...more,
});
const de = (jurisdiction = "DE"): Obj => ({
  issueDate: "2026-11-30",
  currency: "EUR",
  jurisdiction,
  bands: [band("standard", 19, 333.33, 63.33), band("reduced", 7, 111.11, 7.78)],
  promptPaymentDiscount: { percent: 3, days: 14 },
});
const us = (more: Obj = {}, disc: Obj = { percent: 2, days: 10 }): Obj => ({
  issueDate: "2026-11-30",
  currency: "USD",
  jurisdiction: "US",
  usState: "NY",
  dueDays: 30,
  bands: [band("standard", 8.875, 10000, 887.5)],
  promptPaymentDiscount: disc,
  ...more,
});

test("UK preview works per band under the post-2015 rule", () => {
  const r = run(preview);
  assertEquals(cells(r)[1], ["5%", "£47,010.01", "£2,350.50", "£49,360.51"]);
  assertEquals(cells(r)[0], ["20%", "£1,252,631.25", "£250,526.25", "£1,503,157.50"]);
  assertEquals(cells(r)[2], ["0%", "£12,168.00", "£0.00", "£12,168.00"]);
  assertEquals(cells(r)[3], ["Reverse charge", "£93,892.50", "£0.00", "£93,892.50"]);
  assertEquals([r.payableMinor, r.deadline, r.showTable], [165857851, "2026-12-14", true]);
  assertEquals(r.statement, "No credit note will be issued. Following payment you must ensure you have only recovered the VAT actually paid.");
  assertEquals(r.payable, "Total payable if paid by 14 December 2026: £1,658,578.51.");
  assertEquals(r.terms, "A discount of 2.5% of the full price applies if payment is made within 14 days of the invoice date.");
});

test("UK with a credit note changes only the statement", () => {
  const inv = { ...preview, promptPaymentDiscount: { percent: 2.5, days: 14, issuesCreditNote: true } };
  const r = run(inv);
  assertEquals(r.statement, "If payment is made within 14 days, we will issue a credit note for the discount and the VAT on it.");
  assertEquals(r.payableMinor, 165857851);
});

test("UK single band still shows the table, and the other slots are empty", () => {
  const r = run(uk());
  assertEquals(cells(r), [["20%", "£98.00", "£19.60", "£117.60"]]);
  assertEquals(r.showTable, true);
  assertEquals(r.rows.length, 6);
  assertEquals(r.rows.slice(1).every((x) => x.label === ""), true);
});

test("reverse-charge VAT stays 0 and a stated VAT on it throws", () => {
  const r = run(uk({ bands: [band("reverse-charge", null, 100, 0)] }));
  assertEquals(cells(r), [["Reverse charge", "£98.00", "£0.00", "£98.00"]]);
  throws(uk({ bands: [band("standard", 20, 100, 20), band("zero", 0, 5, 0), band("reverse-charge", null, 100, 5)] }), "invoice.bands[2].vat must be 0 on a reverse-charge band");
});

test("DE splits the reduction per rate by largest remainder", () => {
  const r = run(de());
  assertEquals(cells(r), [
    ["19 %", "396,66 €", "11,90 €", "384,76 €"],
    ["7 %", "118,89 €", "3,57 €", "115,32 €"],
  ]);
  assertEquals(r.terms, "3 % Skonto bei Zahlung bis 14.12.2026.");
  assertEquals(plain(r.payable), "Zahlbetrag abzüglich Skonto: 500,08 €.");
  assertEquals([r.heading, r.statement, r.showTable, r.payableMinor], ["Skonto", "", true, 50008]);
});

test("DE ties go to the earliest band", () => {
  const r = run({
    ...de(),
    bands: [band("standard", 19, 0.84, 0.16), band("reduced", 7, 0.93, 0.07), band("zero", 0, 1, 0)],
    promptPaymentDiscount: { percent: 0.5, days: 14 },
  });
  assertEquals(r.rows.slice(0, 3).map((x) => plain(x.second)), ["0,01 €", "0,01 €", "0,00 €"]);
});

test("DE with one band prints terms and payable but no table", () => {
  const r = run({ ...de(), bands: [band("standard", 19, 100, 19)] });
  assertEquals([r.show, r.showTable, cells(r)], [true, false, []]);
  assertEquals(r.terms, "3 % Skonto bei Zahlung bis 14.12.2026.");
  assertEquals(plain(r.payable), "Zahlbetrag abzüglich Skonto: 115,43 €.");
});

test("EU and FR use the same figures in their own words", () => {
  const eu = run(de("EU"));
  assertEquals(cells(eu)[1], ["7%", "€118.89", "€3.57", "€115.32"]);
  assertEquals(eu.terms, "A discount of 3% applies if payment is made by 14 December 2026.");
  assertEquals(eu.payable, "Amount payable if paid by 14 December 2026: €500.08.");
  const fr = run(de("FR"));
  assertEquals(fr.terms, "Escompte de 3 % pour paiement au plus tard le 14/12/2026.");
  assertEquals(cells(fr)[0][0], "19 %");
  assertEquals(plain(fr.payable), "Net à payer après escompte : 500,08 €.");
  assertEquals(fr.heading, "Escompte pour paiement anticipé");
});

test("exempt and reverse-charge labels follow the jurisdiction", () => {
  const bands = [band("standard", 19, 100, 19), band("exempt", null, 10, 0), band("reverse-charge", null, 10, 0)];
  assertEquals(cells(run({ ...de(), bands })).map((c) => c[0]), ["19 %", "steuerfrei", "Reverse-Charge"]);
  assertEquals(cells(run({ ...de("FR"), bands })).map((c) => c[0]), ["19 %", "Exonéré", "Autoliquidation"]);
  assertEquals(cells(run({ ...uk(), bands })).map((c) => c[0]), ["19%", "Exempt", "Reverse charge"]);
});

test("US NY keeps the sales tax on the full price", () => {
  const r = run(us());
  assertEquals(r.terms, "Terms: 2/10, net 30. A 2% discount applies if payment is made by December 10, 2026.");
  assertEquals(r.payable, "Amount due if paid by December 10, 2026: $10,687.50, including $887.50 sales tax on the full price.");
  assertEquals([r.showTable, r.statement, cells(r)], [false, "", []]);
});

test("US CA taxes the discounted price", () => {
  const r = run(us({ usState: "CA", bands: [band("standard", 7.25, 10000, 725)] }));
  assertEquals(r.payable, "Amount due if paid by December 10, 2026: $10,510.50, including $710.50 sales tax on the discounted price.");
});

test("US with no tax has no tax clause, and no dueDays drops net", () => {
  const r = run(us({ usState: "TX", dueDays: undefined, bands: [band("exempt", null, 1000, 0)] }));
  assertEquals(r.terms, "Terms: 2/10. A 2% discount applies if payment is made by December 10, 2026.");
  assertEquals(r.payable, "Amount due if paid by December 10, 2026: $980.00.");
});

test("US state is required, trimmed, upper-cased and looked up", () => {
  throws(us({ usState: undefined }), "invoice.usState");
  throws(us({ usState: "OR" }), "invoice.usState OR is not in the prompt-payment sales-tax table; add it to the deriver");
  assertEquals(run(us({ usState: " ny " })).show, true);
});

test("prints nothing, and does not throw, without a usable discount", () => {
  const empty = run(uk({}, { percent: 0, days: 14 }));
  assertEquals(empty.rows.length, 6);
  const cases: Obj[] = [
    uk({ promptPaymentDiscount: undefined }),
    uk({}, { percent: 0, days: 14 }),
    uk({}, { percent: null, days: 14 }),
    uk({}, { percent: 2 }),
    uk({}, { percent: 2, days: 0 }),
    uk({}, { percent: 2, days: -5 }),
    uk({ bands: [band("standard", 20, 0, 0)] }),
    { promptPaymentDiscount: { percent: 0, days: 5 }, bands: [], jurisdiction: "XX" },
  ];
  for (const c of cases) {
    const r = run(c);
    assertEquals(r.show, false);
    assertEquals([r.heading, r.terms, r.statement, r.payable, r.deadline, r.showTable], ["", "", "", "", "", false]);
    assertEquals(r.rows.length, 6);
    assertEquals(r.rows.every((x) => !x.label && !x.first && !x.second && !x.third), true);
  }
});

test("bad input throws and names the field", () => {
  throws(uk({}, { percent: -1, days: 5 }), "percent");
  throws(uk({}, { percent: 100, days: 5 }), "percent");
  throws(uk({}, { percent: 2.555, days: 5 }), "percent");
  throws(uk({}, { percent: 2, days: 10.5 }), "days");
  throws(uk({ dueDays: 30 }, { percent: 2, days: 45 }), "invoice.promptPaymentDiscount.days (45) is after the due date (dueDays 30)");
  assertEquals(run(uk({ dueDays: 30 }, { percent: 2, days: 30 })).show, true);
  throws(uk({ dueDays: 0 }), "invoice.dueDays");
  throws(uk({ bands: [] }), "invoice.bands");
  throws(uk({ bands: [1, 2, 3, 4, 5, 6, 7].map((r) => band("standard", r, 10, 1)) }), "6");
  throws(uk({ bands: [band("standard", 20, 10, 2), band("standard", 20, 10, 2)] }), "invoice.bands[1]");
  throws(uk({ bands: [band("standard", 20, -1, 0)] }), "invoice.bands[0].net");
  throws(uk({ bands: [band("standard", 20, 1.005, 0)] }), "invoice.bands[0].net");
  throws(uk({ bands: [band("standard", null, 10, 2)] }), "invoice.bands[0].rate");
  throws(uk({ bands: [band("mystery", 20, 10, 2)] }), "invoice.bands[0].treatment");
  throws(uk({ issueDate: "soon" }), "invoice.issueDate");
  throws(uk({ issueDate: undefined }), "invoice.issueDate");
  throws(uk({ currency: "XYZ" }), "invoice.currency");
  throws(uk({ jurisdiction: "XX" }), "invoice.jurisdiction");
});

test("JPY has no minor units", () => {
  const r = run(uk({ currency: "JPY", bands: [band("standard", 10, 100000, 10000)] }));
  assertEquals(plain(cells(r)[0].join("|")), "10%|JP¥98,000|JP¥9,800|JP¥107,800");
  assertEquals(r.payableMinor, 107800);
});

test("a billion-pound net has no float drift", () => {
  const r = run(uk({ bands: [band("standard", 20, 1_000_000_000, 200_000_000)] }, { percent: 2.5, days: 14 }));
  assertEquals(cells(r)[0], ["20%", "£975,000,000.00", "£195,000,000.00", "£1,170,000,000.00"]);
  assertEquals(r.payableMinor, 117_000_000_000);
  const odd = run(uk({ bands: [band("standard", 17.5, 999_999_999.99, 0)] }, { percent: 0.01, days: 1 }));
  // 99999999999 - 10000000 (half up of 9999999.9999) = 99989999999; VAT 17.5% of that, once.
  assertEquals(odd.payableMinor, 99989999999 + Math.round((99989999999 * 175) / 1000));
});

test("a valid large two-decimal net is not rejected for float error", () => {
  run(uk({ bands: [band("standard", 20, 656870881.93, 0)] }));
});

// ---------------------------------------------------------------------------
// Publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Invoice">
      <Section id="body" title="Body">
        <PromptPaymentDiscount />
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

test("a publish build carries 9 arms and resolves like a direct build", async () => {
  assertEquals(branches, 9);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  const armIds = all.filter((n) => (n as { when?: unknown }).when !== undefined).map((n) => n.id);
  assertEquals(armIds.length, 9);
  for (const id of [
    "prompt-payment-discount",
    "prompt-payment-discount-statement",
    "prompt-payment-discount-figures",
    ...[1, 2, 3, 4, 5, 6].map((i) => "prompt-payment-discount-band-" + i),
  ]) {
    assertEquals(armIds.includes(id), true);
  }
  const section = all.find((n) => n.id === "prompt-payment-discount") as unknown as { derivers: unknown[] };
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).name, "promptPaymentDiscount");

  const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);
  for (const invoice of [preview, de(), us(), uk({}, { percent: 0, days: 1 })]) {
    const direct = await buildDocument(templateOf(), { invoice }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);
    const resolved = await resolveDocument(published, state({ invoice }), { derivers });
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
    const text = JSON.stringify(resolved.nodes);
    const r = run(invoice);
    if (r.show) {
      assert.ok(text.includes(JSON.stringify(r.terms).slice(1, -1)));
      assertEquals(collect(resolved.nodes).filter((n) => n.id.startsWith("prompt-payment-discount-band-")).length, r.showTable ? cells(r).length : 0);
    } else {
      assertEquals(text.includes("prompt-payment-discount-terms"), false);
    }
  }
});
