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
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/financial-assistance-notice.deriver.ts")
>("components/derivers/financial-assistance-notice.deriver.ts");
const { FinancialAssistanceNotice } = await loadRegistryModule<
  typeof import("../registry/components/financial-assistance-notice/financial-assistance-notice.node.tsx")
>("components/financial-assistance-notice/financial-assistance-notice.node.tsx");

type Obj = Record<string, unknown>;
type Result = Record<
  "notice" | "noticeIncomplete" | "collections" | "apply" | "suspended" | "free" | "agbBasis" | "owed",
  string | null
>;

const facility: Obj = { taxExempt: true, fapUrl: " https://fap.example/policy ", fapPhone: "1-800-555-0142" };
const run = (billing: Obj | undefined, fac: Obj = {}) =>
  mod.run([{ ...facility, ...fac }, billing], {} as RuntimeState) as Result;
const nonNull = (r: Result) => Object.entries(r).filter(([, v]) => v !== null).map(([k]) => k);
const b = (extra: Obj = {}): Obj => ({ firstPostDischargeStatementDate: "2027-01-15", ...extra });
const partial = (extra: Obj = {}): Obj =>
  b({ application: { status: "approved", discount: "partial" }, grossCharges: 1284567.89, agbPercent: 37.85, ...extra });

test("the notice prints the trimmed address and phone", () => {
  assertStringIncludes(run(b()).notice!, "are at https://fap.example/policy, or call 1-800-555-0142 and");
});

test("full timeline with no ECA notice", () => {
  const r = run(b());
  assertStringIncludes(r.collections!, "before May 15, 2027, and not until at least 30 days after");
  assertStringIncludes(r.apply!, "until September 12, 2027.");
});

test("leap year", () => {
  assertStringIncludes(run(b({ firstPostDischargeStatementDate: "2028-01-15" })).collections!, "May 14, 2028");
});

test("an ECA notice moves the date, and the 120-day bar wins when it is earlier", () => {
  const first = "2026-11-02";
  const moved = run(b({ firstPostDischargeStatementDate: first, ecaNoticeDate: "2027-02-25" }));
  assertStringIncludes(moved.collections!, "before March 27, 2027.");
  assertStringIncludes(moved.apply!, "until June 30, 2027.");
  const early = run(b({ firstPostDischargeStatementDate: first, ecaNoticeDate: "2026-10-01" }));
  assertStringIncludes(early.collections!, "before March 2, 2027.");
});

test("a late ECA notice moves the application deadline", () => {
  const r = run(b({ firstPostDischargeStatementDate: "2026-11-02", ecaNoticeDate: "2027-06-20" }));
  assertStringIncludes(r.collections!, "July 20, 2027");
  assertStringIncludes(r.apply!, "until July 20, 2027.");
});

test("year boundary", () => {
  assertStringIncludes(run(b({ firstPostDischargeStatementDate: "2026-12-31" })).collections!, "April 30, 2027");
});

test("pending, denied, free", () => {
  assertEquals(nonNull(run(b({ application: { status: "pending" } }))), ["notice", "suspended"]);
  assertEquals(run(b({ application: { status: "denied" } })), run(b({ application: { status: "none" } })));
  assertEquals(run(b({ application: { status: "denied" } })), run(b()));
  assertEquals(nonNull(run(b({ application: { status: "approved", discount: "free" } }))), ["notice", "free"]);
});

test("partial approval: AGB, and what is owed", () => {
  const r = run(partial());
  assertEquals(nonNull(r), ["notice", "collections", "agbBasis", "owed"]);
  assertStringIncludes(r.agbBasis!, "37.85% of $1,284,567.89 is $486,208.95.");
  assertEquals(r.owed, "Amount you owe for this care: $486,208.95.");
  assertEquals(
    run(partial({ balanceDue: 812450 })).owed,
    "Amount you owe for this care: $486,208.95. This is the AGB limit, not the $812,450.00 that would otherwise have been billed.",
  );
  assertEquals(
    run(partial({ balanceDue: 100 })).owed,
    "Amount you owe for this care: $100.00. This is below the AGB limit, so the limit does not change it.",
  );
});

test("half-up rounding in whole cents", () => {
  assertStringIncludes(run(partial({ grossCharges: 2.01, agbPercent: 50 })).agbBasis!, "is $1.01.");
  assertStringIncludes(run(partial({ grossCharges: 100.1, agbPercent: 25 })).agbBasis!, "is $25.03.");
  assertEquals(run(partial({ grossCharges: 0 })).owed, "Amount you owe for this care: $0.00.");
});

test("a non-USD currency prints its own symbol", () => {
  assertStringIncludes(run(partial({ currency: "EUR" })).owed!, "€486,208.95");
});

test("bad data throws, naming the field", async () => {
  for (const [billing, field] of [
    [partial({ grossCharges: -1 }), "grossCharges"],
    [partial({ grossCharges: NaN }), "grossCharges"],
    [partial({ grossCharges: undefined }), "grossCharges"],
    [partial({ agbPercent: 0 }), "agbPercent"],
    [partial({ agbPercent: 100.5 }), "agbPercent"],
    [partial({ agbPercent: undefined }), "agbPercent"],
    [partial({ balanceDue: -5 }), "balanceDue"],
    [b({ application: { status: "approved" } }), "discount"],
    [b({ application: { status: "maybe" } }), "status"],
    [b({ firstPostDischargeStatementDate: undefined }), "firstPostDischargeStatementDate"],
    [b({ firstPostDischargeStatementDate: "2027-02-30" }), "firstPostDischargeStatementDate"],
    [b({ ecaNoticeDate: "soon" }), "ecaNoticeDate"],
    [undefined, "firstPostDischargeStatementDate"],
  ] as const) {
    await assertRejects(() => Promise.resolve().then(() => run(billing as Obj | undefined)), Error, field);
  }
});

test("a missing phone or blank address prints the warning, and the timeline still prints", () => {
  for (const fac of [{ fapPhone: undefined }, { fapUrl: "   " }]) {
    const r = run(b(), fac);
    assertEquals(r.notice, null);
    assertStringIncludes(r.noticeIncomplete!, "notice incomplete");
    assertEquals(nonNull(r), ["noticeIncomplete", "collections", "apply"]);
  }
});

test("for-profit: nothing, or the notice alone, whatever the billing", () => {
  assertEquals(nonNull(run(undefined, { taxExempt: false })), []);
  assertEquals(nonNull(run({}, { taxExempt: false, stateNoticeRequired: false })), []);
  assertEquals(nonNull(run(undefined, { taxExempt: false, stateNoticeRequired: true })), ["notice"]);
  assertEquals(nonNull(run(b(), { taxExempt: false, stateNoticeRequired: true })), ["notice"]);
});

test("a missing taxExempt is read as tax-exempt", () => {
  assertEquals(nonNull(run(b(), { taxExempt: undefined })), ["notice", "collections", "apply"]);
});

// The node, published and resolved by an engine that only has the bundle.
const documentOf = () =>
  template(
    <Document title="Statement">
      <FinancialAssistanceNotice />
    </Document>,
  );

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((n) => {
    const own = n as { id?: string; children?: DocumentNode[] };
    return [...(own.id && n.kind === "paragraph" ? [own.id] : []), ...ids(own.children ?? [])];
  });
}

const publish = () =>
  buildDocument(documentOf(), createPublishData() as { facility: Obj; billing: Obj }, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });

test("published, the node has its eight guarded paragraphs", async () => {
  const json = JSON.stringify(await publish());
  for (
    const id of [
      "financial-assistance-notice",
      "financial-assistance-notice-incomplete",
      "financial-assistance-suspended",
      "financial-assistance-free",
      "financial-assistance-agb-basis",
      "financial-assistance-owed",
      "financial-assistance-collections",
      "financial-assistance-apply",
    ]
  ) {
    assertStringIncludes(json, `"${id}"`);
  }
  for (const f of ["notice", "noticeIncomplete", "suspended", "free", "agbBasis", "owed", "collections", "apply"]) {
    assertStringIncludes(json, `derived.financialAssistanceNotice.${f}`);
  }
});

test("resolved against the bundle, each patient gets their own lines", async () => {
  const published = await publish();
  const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ financialAssistanceNotice: mod.run }));
  const resolve = async (data: Obj) => {
    const state: RuntimeState = {
      ctx: { ...data },
      derived: {},
      dataProvider: new InMemoryDataProvider(data),
      aiClient: { generateParagraph: () => "" },
    };
    return ids((await resolveDocument(published, state, { derivers })).nodes);
  };

  assertEquals(await resolve({ facility, billing: b({ application: { status: "pending" } }) }), [
    "financial-assistance-notice",
    "financial-assistance-suspended",
  ]);
  assertEquals(await resolve({ facility: { ...facility, taxExempt: false }, billing: {} }), []);
});
