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

const { CreditsAppliedToInvoice } = await loadRegistryModule<
  typeof import("../registry/components/credits-applied-to-invoice/credits-applied-to-invoice.node.tsx")
>("components/credits-applied-to-invoice/credits-applied-to-invoice.node.tsx");
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/credits-applied-to-invoice.deriver.ts")
>("components/derivers/credits-applied-to-invoice.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ creditsAppliedToInvoice: mod.run }));

type Obj = Record<string, unknown>;
const inv: Obj = { currency: "GBP", issueDate: "2026-11-30", totalMinor: 12000 };
const cn = (id: string, remainingMinor: number, more: Obj = {}): Obj => ({
  id,
  kind: "credit-note",
  issuedOn: "2026-01-01",
  currency: "GBP",
  remainingMinor,
  ...more,
});
const poa = (id: string, remainingMinor: number, more: Obj = {}) => cn(id, remainingMinor, { kind: "payment-on-account", ...more });

interface Row {
  reference: string;
  appliedMinor: number;
  carriedForwardMinor: number;
  applied: string;
}
interface Out {
  show: boolean;
  outcome: string;
  intro: string;
  hasCreditNotes: boolean;
  creditNotes: Row[];
  hasPayments: boolean;
  payments: Row[];
  excluded: Array<{ amount: string; reason: string }>;
  total: string;
  due: string;
  dueMinor: number;
  appliedMinor: number;
  lessCreditNotes: string | null;
  lessPayments: string | null;
  vat: string | null;
  carriedForward: string | null;
  carriedForwardMinor: number;
}
const call = (invoice: Obj, credits: Obj) => mod.run([invoice, credits], {} as RuntimeState) as Out;
const run = (open: Obj[] | undefined, more: Obj = {}, invoice: Obj = inv, allocation?: Obj[]) =>
  call({ ...invoice, ...more }, { open, allocation });
const rejects = (invoice: Obj, credits: Obj, path: string) => {
  let message = "";
  try {
    call(invoice, credits);
  } catch (e) {
    message = (e as Error).message;
  }
  assertStringIncludes(message, path);
};

// ---------------------------------------------------------------------------
// The deriver
// ---------------------------------------------------------------------------

test("nothing open prints nothing", async () => {
  for (const open of [undefined, [], [cn("A", 0)]]) assertEquals(run(open).show, false);
  assertEquals(call(inv, {}).show, false);
  const doc = await build({ invoice: inv, credits: {} });
  assertEquals(ids(doc).includes("credits-applied"), false);
});

test("one credit note, part paid", () => {
  const o = run([cn("A", 5000)]);
  assertEquals([o.appliedMinor, o.dueMinor, o.outcome], [5000, 7000, "part"]);
  assertEquals([o.lessCreditNotes, o.lessPayments, o.carriedForward, o.hasPayments], ["−£50.00", null, null, false]);
});

test("a payment bigger than the total is capped and the rest carried forward", () => {
  const o = run([poa("P", 20000)]);
  assertEquals([o.appliedMinor, o.dueMinor, o.outcome, o.carriedForward, o.hasCreditNotes], [12000, 0, "covered", "£80.00", false]);
});

test("a credit of exactly the total leaves nothing carried forward", () => {
  const o = run([cn("A", 12000)]);
  assertEquals([o.outcome, o.carriedForward], ["covered", null]);
});

test("a zero-value invoice applies nothing", () => {
  const o = run([cn("A", 5000)], { totalMinor: 0 });
  assertEquals([o.appliedMinor, o.outcome, o.carriedForward], [0, "none", "£50.00"]);
  assertStringIncludes(o.intro, "There is nothing on this invoice");
});

test("oldest first, ties by input order, kind ignored", () => {
  const a = cn("A", 3000, { issuedOn: "2026-05-01" });
  const b = cn("B", 3000, { issuedOn: "2026-01-01" });
  const o = run([a, b], { totalMinor: 4000 });
  assertEquals(o.creditNotes.map((r) => r.reference), ["B", "A"]);
  assertEquals(o.creditNotes.map((r) => [r.appliedMinor, r.carriedForwardMinor]), [[3000, 0], [1000, 2000]]);
  assertStringIncludes(o.intro, "oldest first");
  const t = run([cn("X", 3000), poa("Y", 3000)], { totalMinor: 4000 });
  assertEquals([t.creditNotes[0].appliedMinor, t.payments[0].appliedMinor], [3000, 1000]);
});

test("an allocation replaces oldest first", () => {
  const a = cn("A", 3000, { issuedOn: "2026-05-01" });
  const b = cn("B", 3000, { issuedOn: "2026-01-01" });
  const o = run([a, b], {}, { ...inv, totalMinor: 4000 }, [{ creditId: "A", amountMinor: 2500 }]);
  assertEquals(o.creditNotes.map((r) => [r.reference, r.appliedMinor, r.carriedForwardMinor, r.applied]), [
    ["B", 0, 3000, "£0.00"],
    ["A", 2500, 500, "£25.00"],
  ]);
  assertEquals(o.dueMinor, 1500);
  assertStringIncludes(o.intro, "as allocated");
  assertEquals(run([a, b], {}, { ...inv, totalMinor: 4000 }, []).creditNotes[0].appliedMinor, 3000);
});

test("an allocation is capped by the remainder and by what is owed", () => {
  assertEquals(run([cn("A", 3000)], {}, inv, [{ creditId: "A", amountMinor: 9999 }]).appliedMinor, 3000);
  assertEquals(run([cn("A", 3000)], {}, { ...inv, totalMinor: 2000 }, [{ creditId: "A", amountMinor: 9999 }]).appliedMinor, 2000);
});

test("an allocation naming an excluded credit applies nothing", () => {
  const o = run([cn("A", 3000, { expiresOn: "2026-01-01" })], {}, inv, [{ creditId: "A", amountMinor: 3000 }]);
  assertEquals(o.appliedMinor, 0);
  assertEquals(o.excluded[0].reason, "Expired on 1 January 2026");
});

test("expiry is judged on the issue date", () => {
  assertEquals(run([cn("A", 100, { expiresOn: "2026-11-30" })]).appliedMinor, 100);
  const o = run([cn("A", 100, { expiresOn: "2026-11-29" })]);
  assertEquals([o.appliedMinor, o.excluded[0].reason], [0, "Expired on 29 November 2026"]);
});

test("only disputed === true is a dispute", () => {
  assertEquals(run([cn("A", 100, { disputed: true })]).excluded[0].reason, "On hold – disputed");
  assertEquals(run([cn("A", 100, { disputed: "yes" })]).appliedMinor, 100);
});

test("currency is compared trimmed and upper-cased, never converted", () => {
  const o = run([cn("A", 12500, { currency: "USD" }), cn("B", 100, { currency: " gbp " })]);
  assertEquals(o.excluded[0].amount, "US$125.00");
  assertEquals(o.excluded[0].reason, "Not applied – different currency");
  assertEquals(o.appliedMinor, 100);
  assertEquals(o.carriedForwardMinor, 0);
});

test("the expired reason wins over disputed and currency", () => {
  const o = run([cn("A", 100, { expiresOn: "2026-01-01", disputed: true, currency: "USD" })]);
  assertEquals(o.excluded[0].reason, "Expired on 1 January 2026");
});

test("a currency without decimals", () => {
  const o = run([cn("A", 2500, { currency: "JPY" })], { currency: "JPY", totalMinor: 10000 });
  assertEquals([o.dueMinor, o.due], [7500, "JP¥7,500"]);
});

test("VAT is repeated, never used", () => {
  const without = run([cn("A", 5000)]);
  const withVat = run([cn("A", 5000)], { vatTotalMinor: 2000 });
  assertEquals([without.vat, withVat.vat], [null, "£20.00"]);
  assertEquals([without.dueMinor, without.total], [withVat.dueMinor, withVat.total]);
});

test("every credit excluded", async () => {
  const open = [cn("A", 100, { disputed: true }), cn("B", 100, { currency: "USD" })];
  const o = run(open);
  assertEquals([o.outcome, o.dueMinor, o.hasCreditNotes, o.hasPayments], ["none", 12000, false, false]);
  assertStringIncludes(o.intro, "None of the credits");
  const found = ids(await build({ invoice: inv, credits: { open } }));
  assertEquals(found.filter((id) => id.startsWith("credits-applied-")).includes("credits-applied-excluded"), true);
  assertEquals(found.includes("credits-applied-credit-notes"), false);
  assertEquals(found.includes("credits-applied-reconciliation"), true);
});

test("forty credits always tie out", () => {
  let seed = 7;
  const rand = (n: number) => (seed = (seed * 48271) % 2147483647) % n;
  const open = Array.from({ length: 40 }, (_, i) =>
    (i % 2 ? cn : poa)(`C${i}`, rand(900) + 1, { issuedOn: `2026-0${(rand(9)) + 1}-1${rand(9)}` })
  );
  const o = run(open, { totalMinor: 9999 });
  assertEquals(o.appliedMinor + o.dueMinor, 9999);
  assertEquals(o.dueMinor >= 0, true);
  assertEquals(o.creditNotes.length + o.payments.length, 40);
});

test("bad data is rejected naming the field", () => {
  const ok = [cn("A", 100)];
  rejects(inv, { open: [cn("A", -1)] }, "credits.open[0].remainingMinor");
  rejects({ ...inv, totalMinor: 12.5 }, { open: ok }, "invoice.totalMinor");
  rejects({ ...inv, totalMinor: -1 }, { open: ok }, "invoice.totalMinor");
  rejects({ ...inv, vatTotalMinor: 12001 }, { open: ok }, "invoice.vatTotalMinor");
  rejects({ ...inv, currency: "pounds" }, { open: ok }, "invoice.currency");
  rejects({ ...inv, issueDate: "yesterday" }, { open: ok }, "invoice.issueDate");
  rejects(inv, { open: [cn("A", 1, { currency: "pounds" })] }, "credits.open[0].currency");
  rejects(inv, { open: [cn("A", 1, { issuedOn: "soon" })] }, "credits.open[0].issuedOn");
  rejects(inv, { open: [cn("A", 1, { expiresOn: "soon" })] }, "credits.open[0].expiresOn");
  rejects(inv, { open: [cn("A", 1, { kind: "refund" })] }, "credits.open[0].kind");
  rejects(inv, { open: [cn("A", 1), cn("A", 2)] }, "credits.open[1].id");
  rejects(inv, { open: ok, allocation: [{ creditId: "Z", amountMinor: 1 }] }, "credits.allocation[0].creditId");
  rejects(inv, { open: ok, allocation: [{ creditId: "A", amountMinor: 1 }, { creditId: "A", amountMinor: 1 }] }, "credits.allocation[1].creditId");
  rejects(inv, { open: ok, allocation: [{ creditId: "A", amountMinor: -1 }] }, "credits.allocation[0].amountMinor");
  // Validation runs even when nothing would print.
  rejects(inv, { open: [cn("A", 0, { issuedOn: "soon" })] }, "credits.open[0].issuedOn");
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Invoice">
      <Section id="body" title="Body">
        <CreditsAppliedToInvoice />
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
const ids = (doc: { nodes: DocumentNode[] }) => collect(doc.nodes).map((n) => n.id);
const build = (data: Obj) =>
  buildDocument(templateOf(), data, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);

test("the preview covers the invoice and the derived loops resolve on the engine", async () => {
  const preview = COMPONENTS.find((c) => c.id === "credits-applied-to-invoice")!.previewData as Obj;
  const direct = await build(preview);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const arms = collect(published.nodes).filter((n) => (n as { when?: unknown }).when !== undefined);
  assertEquals(arms.length <= 12, true);
  const state: RuntimeState = {
    ctx: { ...preview },
    derived: {},
    dataProvider: new InMemoryDataProvider(preview),
    aiClient: { generateParagraph: () => "" },
  };
  const resolved = await resolveDocument(published, state, { derivers });
  // A resolved tree numbers its cells differently, so compare what a reader sees.
  const seen = (nodes: DocumentNode[]) =>
    collect(nodes).map((n) => [n.id.startsWith("table-cell") ? "" : n.id, (n as { text?: string }).text]);
  assertEquals(seen(resolved.nodes), seen(direct.nodes));
  const found = ids(resolved);
  for (const id of ["credits-applied-credit-note-0", "credits-applied-credit-note-1", "credits-applied-payment-0"]) {
    assertEquals(found.includes(id), true);
  }
  const text = JSON.stringify(resolved.nodes);
  for (const figure of ["£42,500.00", "£18,927.45", "£1,223,322.55", "£26,677.45", "US$12,500.00", "£214,125.00", "Nothing to pay"]) {
    assertStringIncludes(text, figure);
  }
});

test("the deriver rejects through a build", async () => {
  await assertRejects(() => build({ invoice: { ...inv, totalMinor: -1 }, credits: { open: [cn("A", 1)] } }));
});
