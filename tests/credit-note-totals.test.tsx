import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import { COMPONENTS } from "docxcelerate/registry";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/credit-note-totals.deriver.ts")
>("components/derivers/credit-note-totals.deriver.ts");
const { CreditNoteTotals } = await loadRegistryModule<
  typeof import("../registry/components/credit-note-totals/credit-note-totals.node.tsx")
>("components/credit-note-totals/credit-note-totals.node.tsx");

type Out = ReturnType<NonNullable<typeof mod.__types>>;
// deno-lint-ignore no-explicit-any
type Obj = Record<string, any>;

const run = (note: unknown): Out => mod.run([note], {} as RuntimeState) as Out;
const fails = (note: unknown, text: string) => {
  let message = "";
  try {
    run(note);
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }
  assertStringIncludes(message, "creditNoteTotals: ");
  assertStringIncludes(message, text);
};

const band = (vatRate: number | null, net: number, vat: number, label?: string) => ({
  vatRate,
  net,
  vat,
  ...(label ? { label } : {}),
});
const original = (bands: unknown[], more: Obj = {}) => ({
  number: "INV-1",
  issueDate: "2026-09-30",
  currency: "GBP",
  bands,
  ...more,
});
const full = (bands: unknown[], more: Obj = {}, previousCredits?: unknown[]) => ({
  original: original(bands, more),
  ...(previousCredits ? { previousCredits } : {}),
  credit: { mode: "full" },
});
const lineNote = (
  lines: Array<{ id?: string; qty: number; unitNet: number; vatRate: number | null }>,
  bands: unknown[],
  credit: unknown[],
  previousCredits?: unknown[],
) => ({
  original: original(bands, {
    lines: lines.map((l, i) => ({ id: "ABC"[i], description: "Item " + "ABC"[i], ...l })),
  }),
  ...(previousCredits ? { previousCredits } : {}),
  credit: { mode: "lines", lines: credit },
});
const amountNote = (amount: unknown, bands: unknown[], more: Obj = {}) => ({
  original: original(bands, more),
  credit: { mode: "amount", amount },
});

const preview = COMPONENTS.find((c) => c.id === "credit-note-totals")!.previewData as Obj;
const bands20 = [band(20, 1000, 200)];

test("P1: the preview data credits lines across two bands, with sterling and an earlier credit", () => {
  const o = run(preview.creditNote);
  assertEquals(o.lines.map((l) => [l.quantity, l.amount]), [["300", "€123,825.00"], ["12,000", "€22,470.00"], ["12.5", "€16,875.00"]]);
  assertEquals(o.lines[0].unitPrice, "€412.75");
  assertEquals(o.lines[1].unitPrice, "€1.8725");
  assertEquals(o.bands.map((b) => [b.label, b.netMinor, b.vatMinor]), [["20%", 14070000, 2814000], ["5%", 2247000, 112350]]);
  assertEquals(o.multipleBands, true);
  assertEquals([o.net, o.vat, o.total], ["€163,170.00", "€29,263.50", "€192,433.50"]);
  assertEquals([o.vatGBP, o.vatGBPMinor], ["£25,070.04", 2507004]);
  assertEquals(o.previouslyCredited, "Earlier credit notes against this invoice: €49,530.00.");
  assertEquals(
    o.original,
    "Credited against invoice SI-2026-09-004817/REV-2 of 30 September 2026, at the VAT rates that invoice charged. " +
      "VAT is converted to sterling at that invoice's rate of £0.8567 to 1 EUR.",
  );
  assertEquals(o.basis, "This credit note credits the items listed below from invoice SI-2026-09-004817/REV-2.");
});

test("F1: a full cancellation of one band", () => {
  const o = run(full(bands20));
  assertEquals(o.bands.map((b) => [b.net, b.vat]), [["£1,000.00", "£200.00"]]);
  assertEquals([o.multipleBands, o.hasLines, o.total, o.previouslyCredited, o.vatGBP], [false, false, "£1,200.00", null, null]);
  assertEquals(o.basis, "This credit note cancels invoice INV-1 in full.");
});

test("F2: a full cancellation copies a band the original rounded down", () => {
  const o = run(full([band(20, 10.05, 2)]));
  assertEquals([o.vatMinor, o.total], [200, "£12.05"]);
});

test("F3: earlier credits come off a full cancellation", () => {
  const o = run(full([band(20, 1000, 200), band(5, 100, 5)], {}, [{ bands: [band(20, 250, 50)] }]));
  assertEquals(o.bands.map((b) => [b.label, b.net, b.vat]), [["20%", "£750.00", "£150.00"], ["5%", "£100.00", "£5.00"]]);
  assertEquals(o.total, "£1,005.00");
  assertEquals(o.basis, "This credit note cancels what remained of invoice INV-1 after earlier credits.");
  assertEquals(o.previouslyCredited, "Earlier credit notes against this invoice: £300.00.");
});

test("F4: nothing left to credit throws", () => {
  fails(full(bands20, {}, [{ bands: bands20 }]), "nothing is left to credit on invoice INV-1");
});

test("F5: sterling is converted once, or copied on a first full cancellation", () => {
  const eur = { currency: "EUR", fxToGBP: 0.8567 };
  const bands = [band(20, 1306120, 261224)];
  assertEquals(run(full(bands, eur)).vatGBP, "£223,790.60");
  assertEquals(run(full(bands, { ...eur, vatGBP: 223790.61 })).vatGBP, "£223,790.61");
  // With an earlier credit the printed figure no longer applies.
  const later = run(full(bands, { ...eur, vatGBP: 223790.61 }, [{ bands: [band(20, 6120, 1224)] }]));
  assertEquals(later.vatGBPMinor, Math.round((26122400 - 122400) * 0.8567));
});

test("F6: fxToGBP on a sterling invoice is ignored", () => {
  assertEquals(run(full(bands20, { fxToGBP: 1.2 })).vatGBP, null);
});

test("L1: a partial return rounds the line once and VAT once", () => {
  const o = run(lineNote([{ qty: 3, unitNet: 3.35, vatRate: 20 }], [band(20, 10.05, 2.01)], [{ id: "A", qty: 1 }]));
  assertEquals([o.net, o.vat, o.total, o.hasLines, o.lines.length], ["£3.35", "£0.67", "£4.02", true, 1]);
});

test("L2: a line of 4.9995 rounds up", () => {
  const o = run(lineNote([{ qty: 2, unitNet: 3.333, vatRate: 20 }], [band(20, 6.67, 1.33)], [{ id: "A", qty: 1.5 }]));
  assertEquals([o.lines[0].amount, o.vat], ["£5.00", "£1.00"]);
});

test("L3: half rounds away from zero, on the line and on the VAT", () => {
  const o = run(lineNote([{ qty: 1, unitNet: 0.125, vatRate: 20 }, { qty: 1, unitNet: 1, vatRate: 20 }], [band(20, 1.13, 0.2)], [{ id: "A", qty: 1 }]));
  assertEquals([o.lines[0].amount, o.vat], ["£0.13", "£0.03"]);
});

test("L4: a band returned whole is copied, not recomputed", () => {
  const o = run(lineNote([{ qty: 3, unitNet: 3.35, vatRate: 20 }], [band(20, 10.05, 2)], [{ id: "A", qty: 3 }]));
  assertEquals(o.vatMinor, 200);
});

test("L5: bad lines throw, naming the field", () => {
  const A = { qty: 3, unitNet: 3.35, vatRate: 20 };
  const bands = [band(20, 10.05, 2.01)];
  const ok: Obj = lineNote([A], bands, [{ id: "A", qty: 1 }]);
  fails(lineNote([A], bands, [{ id: "Z", qty: 1 }]), "creditNote.credit.lines[0].id");
  fails(lineNote([A], bands, [{ id: "A", qty: 1 }, { id: "A", qty: 1 }]), "creditNote.credit.lines[1].id");
  fails(lineNote([A], bands, [{ id: "A", qty: 0 }]), "creditNote.credit.lines[0].qty");
  fails(lineNote([A], bands, [{ id: "A", qty: -1 }]), "creditNote.credit.lines[0].qty");
  fails(lineNote([A], bands, [{ id: "A", qty: 4 }]), "creditNote.credit.lines[0].qty");
  fails(lineNote([A], bands, []), "creditNote.credit.lines");
  fails({ ...ok, original: { ...ok.original, lines: undefined } }, "creditNote.original.lines");
  fails(lineNote([{ ...A, vatRate: 5 }], bands, [{ id: "A", qty: 1 }]), "creditNote.original.lines[0].vatRate");
  const twice = [ok.original.lines[0], ok.original.lines[0]];
  fails({ ...ok, original: { ...ok.original, lines: twice } }, "creditNote.original.lines[1].id");
});

test("L6: a line credit over what the band has left names the band", () => {
  fails(
    lineNote([{ qty: 1, unitNet: 20, vatRate: 20 }], [band(20, 100, 20)], [{ id: "A", qty: 1 }], [{ bands: [band(20, 90, 18)] }]),
    "the 20% band can take at most £10.00 net and £2.00 VAT",
  );
});

test("L7: lines that price at nothing credit nothing", () => {
  fails(lineNote([{ qty: 1, unitNet: 0, vatRate: 20 }, { qty: 1, unitNet: 5, vatRate: 20 }], bands20, [{ id: "A", qty: 1 }]), "nothing to credit");
});

test("A1: a goodwill amount is spread by largest remainder", () => {
  const o = run(amountNote(100, [band(20, 1000, 200), band(0, 500, 0)]));
  assertEquals(o.bands.map((b) => [b.label, b.net, b.vat, b.netMinor + b.vatMinor]), [
    ["20%", "£58.82", "£11.77", 7059],
    ["0%", "£29.41", "£0.00", 2941],
  ]);
  assertEquals(o.total, "£100.00");
  assertEquals(o.basis, "This credit note is a credit of £100.00 against invoice INV-1, spread across its VAT rates in proportion to what was charged at each.");
});

test("A2: an amount equal to everything left takes each band exactly", () => {
  const o = run(amountNote(1700, [band(20, 1000, 200), band(0, 500, 0)]));
  assertEquals(o.bands.map((b) => [b.net, b.vat]), [["£1,000.00", "£200.00"], ["£500.00", "£0.00"]]);
});

test("A3: a tie goes to the first band", () => {
  const o = run(amountNote(0.01, [band(20, 100, 20), band(5, 114.29, 5.71)]));
  assertEquals(o.bands.map((b) => [b.label, b.net, b.vat]), [["20%", "£0.01", "£0.00"]]);
  assertEquals(o.multipleBands, false);
});

test("A4: a bad amount throws", () => {
  for (const amount of [0, -5, 10.005, 1200.01, undefined]) fails(amountNote(amount, bands20), "creditNote.credit.amount");
});

test("A5: a currency with no minor unit", () => {
  const o = run(amountNote(5500, [band(10, 10000, 1000)], { currency: "JPY" }));
  assertEquals([o.net, o.vat, o.total], ["JP¥5,000", "JP¥500", "JP¥5,500"]);
});

test("A6: the VAT is clamped to what the band has left", () => {
  const o = run(amountNote(1199.98, [band(20, 1000, 199.99)]));
  assertEquals(o.netMinor + o.vatMinor, 119998);
  assertEquals(o.vatMinor <= 19999 && o.netMinor <= 100000, true);
});

test("A7: shares always add up and never exceed a band", () => {
  const bands = [band(20, 333.33, 66.67), band(5, 12.34, 0.62), band(12.5, 0.07, 0.01)];
  for (const amount of [0.01, 0.5, 7.77, 100, 413.04]) {
    const o = run(amountNote(amount, bands));
    assertEquals(o.totalMinor, Math.round(amount * 100));
  }
});

test("V1: an exempt band keeps its own label", () => {
  assertEquals(run(full([band(null, 300, 0, "Exempt")])).bands[0].label, "Exempt");
  assertEquals(run(full([band(null, 300, 0)])).bands[0].label, "No VAT");
  fails(full([band(null, 300, 1)]), "creditNote.original.bands[0].vat");
});

test("V2: zero-rated is not exempt", () => {
  assertEquals(run(full([band(0, 300, 0)])).bands[0].label, "0%");
});

test("V3: bad input throws, naming the field", () => {
  fails(full([band(20, 1, 0.2), band(20, 1, 0.2)]), "creditNote.original.bands[1].vatRate");
  fails(full(bands20, {}, [{ bands: [band(5, 1, 0.05)] }]), "creditNote.previousCredits[0].bands[0].vatRate");
  fails(full([band(20, -1, 0)]), "creditNote.original.bands[0].net");
  fails({ credit: { mode: "full" } }, "creditNote.original");
  fails({ original: original(bands20) }, "creditNote.credit");
  fails({ ...full(bands20), credit: { mode: "partial" } }, "creditNote.credit.mode");
  fails(full(bands20, { currency: "EURO" }), "creditNote.original.currency");
  for (const fx of [0, -1, NaN]) fails(full(bands20, { currency: "EUR", fxToGBP: fx }), "creditNote.original.fxToGBP");
  fails(full([band(20, 1e21, 0)]), "creditNote.original.bands[0].net");
  fails(full(bands20, { issueDate: "30/09/2026" }), "creditNote.original.issueDate");
  fails(full([band(20, 1.005, 0.2)]), "decimal places");
  fails(full(bands20, { number: " " }), "creditNote.original.number");
});

test("V4: no earlier credits reads the same missing or empty", () => {
  assertEquals(run(full(bands20, {}, [])), run(full(bands20)));
  assertEquals(run(full(bands20, {}, [])).previouslyCredited, null);
});

test("V5: five bands keep the invoice's order", () => {
  const o = run(full([band(20, 10, 2), band(12.5, 10, 1.25), band(5, 10, 0.5), band(0, 10, 0), band(null, 10, 0, "Exempt")]));
  assertEquals(o.bands.map((b) => b.label), ["20%", "12.5%", "5%", "0%", "Exempt"]);
  assertEquals(o.multipleBands, true);
});

// ---------------------------------------------------------------------------
// The node
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Credit note">
      <Section id="body" title="Body">
        <CreditNoteTotals />
      </Section>
    </Document>,
  );
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ creditNoteTotals: mod.run }));

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
const texts = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);
const idsOf = (nodes: DocumentNode[]) => new Set(collect(nodes).map((n) => n.id));

test("N1: every node is there for the preview data, and the optional ones go away", async () => {
  const ids = idsOf((await built(preview)).nodes);
  for (
    const id of [
      "credit-note-totals", "credit-note-totals-basis", "credit-note-totals-original", "credit-note-totals-previous",
      "credit-note-totals-lines", "credit-note-totals-lines-head", "credit-note-totals-analysis",
      "credit-note-totals-analysis-head", "credit-note-totals-summary", "credit-note-totals-net",
      "credit-note-totals-vat", "credit-note-totals-vat-gbp", "credit-note-totals-total",
    ]
  ) assertEquals([id, ids.has(id)], [id, true]);

  const bare = idsOf((await built({ creditNote: full(bands20) })).nodes);
  for (const id of ["lines", "analysis", "vat-gbp", "previous"]) {
    assertEquals(["credit-note-totals-" + id, bare.has("credit-note-totals-" + id)], ["credit-note-totals-" + id, false]);
  }
});

function state(data: Obj): RuntimeState {
  return { ctx: { ...data }, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } };
}

test("N2: published and resolved, the loops run over the deriver's rows", async () => {
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const json = JSON.stringify(published);
  assertStringIncludes(json, '"source":{"scope":"derived","path":"creditNoteTotals.bands"}');
  assertStringIncludes(json, '"source":{"scope":"derived","path":"creditNoteTotals.lines"}');

  const a1 = { creditNote: amountNote(100, [band(20, 1000, 200), band(0, 500, 0)]) };
  for (const data of [preview, a1]) {
    const resolved = await resolveDocument(published, state(data), { derivers });
    const direct = await built(data);
    assertEquals(texts(resolved.nodes), texts(direct.nodes));
  }

  // The row text itself, because an unresolved loop leaves no rows to differ.
  const text = JSON.stringify((await resolveDocument(published, state(preview), { derivers })).nodes);
  for (const row of ["€123,825.00", "€22,470.00", "€16,875.00", "€140,700.00", "€28,140.00"]) {
    assertStringIncludes(text, row);
  }
  const second = JSON.stringify((await resolveDocument(published, state(a1), { derivers })).nodes);
  for (const row of ["£58.82", "£11.77", "£29.41"]) assertStringIncludes(second, row);
});

test("N3: a credit that cannot be issued fails the document", async () => {
  await assertRejects(() => built({ creditNote: full(bands20, {}, [{ bands: bands20 }]) }), Error, "nothing is left to credit");
});
