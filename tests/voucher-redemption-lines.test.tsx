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
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/voucher-redemption-lines/voucher-redemption-lines.node.tsx";
const { VoucherRedemptionLines } = await loadRegistryModule<
  typeof import("../registry/components/voucher-redemption-lines/voucher-redemption-lines.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/voucher-redemption-lines.deriver.ts")
>("components/derivers/voucher-redemption-lines.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ voucherRedemptionLines: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
type Slot = { description: string; amount: string };
interface Out {
  subtotal: string;
  discounts: Slot[];
  bands: Slot[];
  total: string;
  tenders: Slot[];
  amountDue: string;
  spvNote: boolean;
  mpvNote: boolean;
  manufacturerNote: boolean;
  subtotalMinor: number;
  taxMinor: number;
  totalMinor: number;
  amountDueMinor: number;
}

const run = (sale: Obj): Out => mod.run([sale], {} as RuntimeState) as Out;
const L = (key: string, netMinor: number, ratePercent: number) => ({ key, netMinor, ratePercent });
const uk = (lines: Obj[], redemptions?: Obj[]) => ({ currency: "GBP", jurisdiction: "UK", lines, redemptions });
const us = (lines: Obj[], redemptions?: Obj[], more: Obj = {}) => ({
  currency: "USD",
  jurisdiction: "US",
  lines,
  redemptions,
  ...more,
});
const rows = (slots: Slot[]) => slots.filter((s) => s.description).map((s) => `${s.description} | ${s.amount}`);
const one = [L("a", 10000, 20)];

test("no redemptions prints plain totals", () => {
  const o = run(uk(one));
  assertEquals(o.subtotal, "£100.00");
  assertEquals(rows(o.bands), ["VAT at 20% on £100.00 | £20.00"]);
  assertEquals([o.total, o.amountDue, rows(o.discounts), rows(o.tenders)], ["£120.00", "£120.00", [], []]);
  assertEquals([o.spvNote, o.mpvNote, o.manufacturerNote], [false, false, false]);
  assertEquals([o.discounts.length, o.bands.length, o.tenders.length], [4, 4, 4]);
});

test("largest remainder keeps every penny of a coupon", () => {
  const o = run(uk([L("a", 3333, 20), L("b", 3333, 5), L("c", 3334, 0)], [{ kind: "discount-coupon", amountMinor: 1000 }]));
  assertEquals(rows(o.bands), ["VAT at 20% on £30.00 | £6.00", "VAT at 5% on £30.00 | £1.50", "VAT at 0% on £30.00 | £0.00"]);
  assertEquals(o.total, "£97.50");
});

test("a tie gives the extra penny to the first line", () => {
  const o = run(uk([L("a", 1000, 20), L("b", 1000, 5), L("c", 1000, 0)], [{ kind: "discount-coupon", amountMinor: 100 }]));
  assertEquals(rows(o.bands), ["VAT at 20% on £9.66 | £1.93", "VAT at 5% on £9.67 | £0.48", "VAT at 0% on £9.67 | £0.00"]);
  assertEquals(o.total, "£31.41");
});

test("a coupon is capped at the lines it applies to", () => {
  const o = run(uk([L("a", 6000, 20), L("b", 3000, 5)], [{ kind: "discount-coupon", amountMinor: 5000, code: "X", appliesToLineKeys: ["b"] }]));
  assertEquals(rows(o.discounts), ["Coupon X (applied £30.00 of £50.00) | -£30.00"]);
  assertEquals(rows(o.bands), ["VAT at 20% on £60.00 | £12.00"]);
  assertEquals(o.total, "£72.00");
});

test("a single-purpose voucher is paid, and left out of the VAT analysis", () => {
  const o = run(uk(one, [{ kind: "spv", amountMinor: 6000, code: "V1" }]));
  assertEquals(rows(o.bands), ["VAT at 20% on £50.00 | £10.00"]);
  assertEquals([o.total, rows(o.tenders), o.amountDue, o.spvNote], ["£120.00", ["Paid by voucher V1 | -£60.00"], "£60.00", true]);

  const big = run(uk(one, [{ kind: "spv", amountMinor: 15000, code: "V1" }]));
  assertEquals(rows(big.bands), []);
  assertEquals([big.total, rows(big.tenders), big.amountDue], ["£120.00", ["Paid by voucher V1 (remaining balance £30.00) | -£120.00"], "£0.00"]);
  // case 16: nothing left to pay, so a later payment prints no row
  const later = run(uk(one, [{ kind: "spv", amountMinor: 15000, code: "V1" }, { kind: "mpv", amountMinor: 500 }]));
  assertEquals(rows(later.tenders).length, 1);
  assertEquals(later.mpvNote, false);
});

test("a coupon comes before a voucher whatever the input order", () => {
  for (const r of [
    [{ kind: "discount-coupon", amountMinor: 2000 }, { kind: "spv", amountMinor: 6000 }],
    [{ kind: "spv", amountMinor: 6000 }, { kind: "discount-coupon", amountMinor: 2000 }],
  ]) {
    const o = run(uk(one, r));
    assertEquals([rows(o.bands), o.total, o.amountDue], [["VAT at 20% on £30.00 | £6.00"], "£96.00", "£36.00"]);
  }
});

test("a multi-purpose voucher pays; the VAT stays on the full price", () => {
  const o = run({ currency: "EUR", jurisdiction: "EU", lines: one, redemptions: [{ kind: "mpv", amountMinor: 20000, code: "M" }] });
  assertEquals(rows(o.bands), ["VAT at 20% on €100.00 | €20.00"]);
  assertEquals(rows(o.tenders), ["Paid by voucher M (remaining balance €80.00) | -€120.00"]);
  assertEquals([o.amountDue, o.mpvNote], ["€0.00", true]);
});

test("US coupons: retailer, manufacturer by default and by flag, promo, gift card", () => {
  const l = [L("a", 10000, 8.875)];
  const r = run(us(l, [{ kind: "retailer-coupon", amountMinor: 1000 }]));
  assertEquals([rows(r.bands), r.total], [["Sales tax at 8.875% on $90.00 | $7.99"], "$97.99"]);

  const m = run(us(l, [{ kind: "manufacturer-coupon", amountMinor: 1000 }]));
  assertEquals([rows(m.bands), m.total, rows(m.tenders), m.amountDue, m.manufacturerNote], [
    ["Sales tax at 8.875% on $100.00 | $8.88"],
    "$108.88",
    ["Manufacturer coupon | -$10.00"],
    "$98.88",
    true,
  ]);
  const f = run(us(l, [{ kind: "manufacturer-coupon", amountMinor: 1000 }], { manufacturerCouponsReduceBase: true }));
  assertEquals([rows(f.bands), f.total, rows(f.discounts), f.manufacturerNote], [["Sales tax at 8.875% on $90.00 | $7.99"], "$97.99", ["Manufacturer coupon | -$10.00"], false]);

  const p = run(us([L("a", 2000, 6)], [{ kind: "promo-card", amountMinor: 2500, code: "P" }]));
  assertEquals([rows(p.discounts), rows(p.bands), p.total, p.amountDue], [["Promotional card P (remaining balance $5.00) | -$20.00"], [], "$0.00", "$0.00"]);

  const g = run(us([L("a", 10000, 8.875), L("b", 2500, 0)], [
    { kind: "retailer-coupon", amountMinor: 500, appliesToLineKeys: ["b"] },
    { kind: "gift-card", amountMinor: 5000, code: "GC" },
  ]));
  assertEquals(rows(g.bands), ["Sales tax at 8.875% on $100.00 | $8.88", "Sales tax at 0% on $20.00 | $0.00"]);
  assertEquals([g.total, rows(g.tenders), g.amountDue], ["$128.88", ["Paid by gift card GC | -$50.00"], "$78.88"]);
});

test("a manufacturer coupon is applied only up to what the lines hold", () => {
  const o = run(us([L("a", 1000, 0)], [{ kind: "manufacturer-coupon", amountMinor: 5000 }]));
  assertEquals(rows(o.tenders), ["Manufacturer coupon (applied $10.00 of $50.00) | -$10.00"]);
});

test("zero amounts and zero lines", () => {
  const base = run(uk(one));
  assertEquals(run(uk(one, [{ kind: "discount-coupon", amountMinor: 0 }, { kind: "spv", amountMinor: 0 }, { kind: "mpv", amountMinor: 0 }])), base);
  const o = run(uk([L("a", 0, 20), L("b", 1000, 20)], [{ kind: "discount-coupon", amountMinor: 100 }]));
  assertEquals(rows(o.bands), ["VAT at 20% on £9.00 | £1.80"]);
});

test("invalid input throws, naming the field", async () => {
  const ok = [L("a", 1000, 20)];
  const red = (r: Obj) => uk(ok, [{ kind: "discount-coupon", amountMinor: 100, ...r }]);
  const bad: Array<[Obj, string]> = [
    [{ ...uk(ok), lines: undefined }, "sale.lines"],
    [uk([]), "sale.lines"],
    [uk([L("a", 1, 0), L("a", 1, 0)]), "sale.lines[1].key"],
    [uk([L("a", -1, 0)]), "sale.lines[0].netMinor"],
    [uk([L("a", 1.5, 0)]), "sale.lines[0].netMinor"],
    [uk([L("a", 1, 0), L("b", 1, 8.8751)]), "sale.lines[1].ratePercent"],
    [uk([L("a", 1, 101)]), "sale.lines[0].ratePercent"],
    [{ ...uk(ok), currency: "XX" }, "sale.currency"],
    [{ ...uk(ok), jurisdiction: "GB" }, "sale.jurisdiction"],
    [red({ amountMinor: -1 }), "redemptions[0].amountMinor"],
    [red({ amountMinor: 1.5 }), "redemptions[0].amountMinor"],
    [uk(ok, [{ kind: "gift-card", amountMinor: 1 }]), "redemptions[0].kind"],
    [us(ok, [{ kind: "spv", amountMinor: 1 }]), "redemptions[0].kind"],
    [red({ kind: "nope" }), "redemptions[0].kind"],
    [red({ appliesToLineKeys: [] }), "appliesToLineKeys"],
    [red({ appliesToLineKeys: ["zzz"] }), "appliesToLineKeys[0]"],
    [uk([L("a", 1000, 20), L("b", 1000, 5)], [{ kind: "spv", amountMinor: 100 }]), "single-purpose"],
    [uk(ok, [{ kind: "spv", amountMinor: 100, spvRatePercent: 5 }]), "spvRatePercent"],
    [uk(ok, Array.from({ length: 5 }, () => ({ kind: "discount-coupon", amountMinor: 1 }))), "at most 4"],
    [uk(ok, Array.from({ length: 5 }, () => ({ kind: "mpv", amountMinor: 1 }))), "at most 4"],
    [uk([L("a", 1, 1), L("b", 1, 2), L("c", 1, 3), L("d", 1, 4), L("e", 1, 5)]), "at most 4"],
  ];
  for (const [sale, field] of bad) {
    await assertRejects(() => Promise.resolve().then(() => run(sale)), Error, field);
  }
});

test("four discounts, four bands and four payments fill every slot", () => {
  const o = run(uk(
    [L("a", 100000, 20), L("b", 100000, 10), L("c", 100000, 5), L("d", 100000, 0)],
    [
      ...Array.from({ length: 4 }, (_, i) => ({ kind: "discount-coupon", amountMinor: 1000 + i, code: `C${i}` })),
      { kind: "spv", amountMinor: 1000, appliesToLineKeys: ["a"], spvRatePercent: 20, code: "S" },
      { kind: "mpv", amountMinor: 1000, code: "M1" },
      { kind: "mpv", amountMinor: 1000, code: "M2" },
      { kind: "mpv", amountMinor: 1000, code: "M3" },
    ],
  ));
  assertEquals([rows(o.discounts).length, rows(o.bands).length, rows(o.tenders).length], [4, 4, 4]);
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const entry = COMPONENTS.find((c) => c.id === "voucher-redemption-lines")!;
const preview = (entry.previewData as { sale: Obj }).sale;

test("the catalog's preview data", () => {
  const o = run(preview);
  assertEquals(o.subtotal, "£2,334.48");
  assertEquals(rows(o.discounts), ["Coupon LOYALTY-ANNIVERSARY-25-OFF-2026 | -£250.00"]);
  assertEquals(rows(o.bands), ["VAT at 20% on £1,370.24 | £274.05", "VAT at 5% on £169.24 | £8.46", "VAT at 0% on £45.00 | £0.00"]);
  assertEquals(o.total, "£2,466.99");
  assertEquals(rows(o.tenders), [
    "Paid by voucher SPA-GIFT-VOUCHER-0004817-UK | -£600.00",
    "Paid by voucher HIGH-STREET-GIFT-CARD-9921-3348-7710 (remaining balance £133.01) | -£1,866.99",
  ]);
  assertEquals([o.amountDue, o.spvNote, o.mpvNote, o.manufacturerNote], ["£0.00", true, true, false]);
});

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Sale">
      <Section id="body" title="Body">
        <VoucherRedemptionLines />
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

const built = (sale: Obj) =>
  buildDocument(templateOf(), { sale }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);

test("a build prints a row only for a filled slot", async () => {
  const ids = collect((await built(preview)).nodes).map((n) => n.id);
  for (const id of ["discount-1", "tax-3", "tender-2", "note-spv", "note-mpv", "amount-due"]) {
    assertEquals(ids.includes(`voucher-redemption-${id}`), true);
  }
  for (const id of ["discount-2", "tax-4", "tender-3", "note-manufacturer"]) {
    assertEquals(ids.includes(`voucher-redemption-${id}`), false);
  }
});

test("a publish build carries 15 arms and one deriver call", async () => {
  assertEquals(branches, 15);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 15);
  // The call is attached to the table and to the notes that read it; one call, shared.
  const withDeriver = all.filter((n) => (n as { derivers?: unknown[] }).derivers?.length) as unknown as Array<{ derivers: unknown[] }>;
  const calls = [...new Set(withDeriver.flatMap((n) => n.derivers).map((d) => JSON.stringify(d)))].map((d) => JSON.parse(d));
  assertEquals(calls.length, 1);
  assertEquals(calls[0].name, "voucherRedemptionLines");
  assertEquals(calls[0].inputs, [{ type: "ref", ref: { scope: "data", path: "sale" } }]);

  const cases: Obj[] = [
    preview,
    uk(one),
    uk(one, [{ kind: "spv", amountMinor: 15000, code: "V1" }]),
    us([L("a", 10000, 8.875)], [{ kind: "manufacturer-coupon", amountMinor: 1000 }]),
    us([L("a", 2000, 6)], [{ kind: "promo-card", amountMinor: 2500, code: "P" }]),
  ];
  for (const sale of cases) {
    const direct = await built(sale);
    const resolved = await resolveDocument(published, state({ sale }), { derivers });
    const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
  }
});
