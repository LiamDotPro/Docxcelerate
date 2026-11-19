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
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";
import { COMPONENTS } from "docxcelerate/registry";

const nodeSource = "components/statement-of-account/statement-of-account.node.tsx";
const { StatementOfAccount } = await loadRegistryModule<
  typeof import("../registry/components/statement-of-account/statement-of-account.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/statement-of-account.deriver.ts")
>("components/derivers/statement-of-account.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ statementOfAccount: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Row { date: string; details: string; debit: string; credit: string; balance: string }
interface Out {
  hasEntries: boolean;
  outcome: string;
  rows: Row[];
  openingBalance: string;
  totalDebits: string;
  totalCredits: string;
  closingBalance: string;
  closingAbs: string;
  dueSuffix: string;
  closingMinor: number;
}

const base = { currency: "GBP", periodStart: "2026-10-01", statementDate: "2026-10-31", openingBalance: 0 };
const run = (o: Obj): Out => mod.run([{ ...base, ...o }], {} as RuntimeState) as Out;
const e = (type: string, number: string, amount: number, more: Obj = {}) => ({ type, date: "2026-10-10", number, amount, ...more });

test("empty periods print the opening balance and pick the outcome", () => {
  const owing = run({ openingBalance: 250 });
  assertEquals([owing.hasEntries, owing.rows, owing.outcome, owing.closingBalance], [false, [], "owing", "£250.00"]);
  const clear = run({});
  assertEquals([clear.outcome, clear.closingBalance], ["clear", "£0.00"]);
  const credit = run({ openingBalance: -40 });
  assertEquals([credit.outcome, credit.openingBalance, credit.closingAbs], ["credit", "£40.00 CR", "£40.00"]);
});

test("a debit, a credit and a credit note sit in their own columns", () => {
  const o = run({ entries: [e("invoice", "INV-1", 100)] });
  assertEquals(o.rows[0], { date: "10 Oct 2026", details: "Invoice INV-1", debit: "£100.00", credit: "", balance: "£100.00" });
  assertEquals([o.totalDebits, o.totalCredits], ["£100.00", "£0.00"]);
  const cn = run({ openingBalance: 50, entries: [e("credit-note", "CN-1", 20)] });
  assertEquals([cn.rows[0].debit, cn.rows[0].credit, cn.rows[0].balance], ["", "£20.00", "£30.00"]);
});

test("same-day entries print in a fixed order", () => {
  const o = run({ entries: [e("payment", "P-1", 100), e("invoice", "I-1", 100)] });
  assertEquals(o.rows.map((r) => r.balance), ["£100.00", "£0.00"]);
  assertStringIncludes(o.rows[0].details, "Invoice");
  const all = run({
    entries: [e("refund", "R", 1), e("payment", "P", 1), e("credit-note", "C", 1), e("invoice", "I", 1), e("adjustment", "A", 1)],
  });
  assertEquals(all.rows.map((r) => r.details.split(" ")[0]), ["Invoice", "Credit", "Payment", "Refund", "Adjustment"]);
  const nums = run({ entries: [e("invoice", "INV-10", 1), e("invoice", "INV-9", 1)] });
  assertEquals(nums.rows.map((r) => r.details), ["Invoice INV-9", "Invoice INV-10"]);
});

test("adjustments go by sign, and a negative elsewhere throws", async () => {
  const plus = run({ entries: [e("adjustment", "A", 15)] }).rows[0];
  assertEquals([plus.debit, plus.credit], ["£15.00", ""]);
  const minus = run({ entries: [e("adjustment", "A", -15)] }).rows[0];
  assertEquals([minus.debit, minus.credit, minus.balance], ["", "£15.00", "£15.00 CR"]);
  await assertRejects(() => Promise.resolve().then(() => run({ entries: [e("invoice", "I", -5)] })), Error, "entries[0].amount");
});

test("rounding happens only when printing", () => {
  const o = run({ entries: [e("invoice", "A", 0.333), e("invoice", "B", 0.333), e("invoice", "C", 0.333)] });
  assertEquals(o.rows.map((r) => r.balance), ["£0.33", "£0.67", "£1.00"]);
  assertEquals(o.closingBalance, "£1.00");
  assertEquals(run({ entries: [e("invoice", "A", 1.005)] }).rows[0].debit, "£1.01");
  assertEquals(run({ currency: "JPY", entries: [e("invoice", "A", 1234.5)] }).rows[0].debit, "JP¥1,235");
  assertStringIncludes(run({ currency: "KWD", entries: [e("invoice", "A", 1.2345)] }).rows[0].debit, "1.235");
  const jpy = run({ currency: "JPY", openingBalance: -2.5 });
  assertEquals([jpy.closingMinor, jpy.closingBalance], [-3, "JP¥3 CR"]);
  assertEquals(run({ openingBalance: 0.004 }).outcome, "clear");
});

test("a supplied closing balance is checked", async () => {
  const entries = [e("invoice", "A", 100)];
  assertEquals(run({ entries, closingBalance: 100.004 }).closingBalance, "£100.00");
  await assertRejects(() => Promise.resolve().then(() => run({ entries, closingBalance: 100.006 })), Error, "does not match");
  assertEquals(run({ entries }).closingBalance, "£100.00");
});

test("the period is inclusive and out-of-period entries are left out", async () => {
  const at = (date: string, n: string) => e("invoice", n, 10, { date });
  const o = run({
    openingBalance: 5,
    entries: [at("2026-09-30", "A"), at("2026-10-01", "B"), at("2026-10-31", "C"), at("2026-11-01", "D")],
  });
  assertEquals(o.rows.map((r) => r.details), ["Invoice B", "Invoice C"]);
  assertEquals(o.openingBalance, "£5.00");
  await assertRejects(
    () => Promise.resolve().then(() => run({ entries: [e("invoice", "A", 1, { date: "2025-01-01", currency: "EUR" })] })),
    Error,
    "currency",
  );
});

test("details and the due date", () => {
  assertEquals(run({ entries: [e("invoice", "INV-1", 1)] }).rows[0].details, "Invoice INV-1");
  assertEquals(
    run({ entries: [e("invoice", "INV-1", 1, { reference: "PO 7", description: "Consulting" })] }).rows[0].details,
    "Invoice INV-1 (PO 7) – Consulting",
  );
  const due = { dueBy: "2026-11-30" };
  assertEquals(run({ ...due, openingBalance: 5 }).dueSuffix, ", payable by 30 November 2026");
  assertEquals([run({ ...due }).dueSuffix, run({ ...due, openingBalance: -5 }).dueSuffix], ["", ""]);
});

test("invalid input throws", async () => {
  const bad: Obj[] = [
    { currency: "XX" },
    { statementDate: "2026-09-30" },
    { entries: [e("bill", "A", 1)] },
    { entries: [e("invoice", "", 1)] },
    { entries: [e("invoice", "A", Number.NaN)] },
    { openingBalance: undefined },
  ];
  for (const b of bad) await assertRejects(() => Promise.resolve().then(() => run(b)), Error);
});

const many = Array.from({ length: 60 }, (_, i) => e("invoice", "INV-" + (i + 1), 10, { date: "2026-10-" + String((i % 28) + 1).padStart(2, "0") }));

test("sixty invoices", () => {
  const o = run({ entries: many, closingBalance: 600 });
  assertEquals(o.rows.length, 60);
  assertEquals(o.rows[59].balance, o.closingBalance);
});

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Statement">
      <Section id="body" title="Body">
        <StatementOfAccount />
      </Section>
    </Document>,
  );
const state = (data: Obj): RuntimeState => ({
  ctx: { ...data },
  derived: {},
  dataProvider: new InMemoryDataProvider(data),
  aiClient: { generateParagraph: () => "" },
});
const collect = (nodes: DocumentNode[], into: DocumentNode[] = []): DocumentNode[] => {
  for (const n of nodes) {
    into.push(n);
    const c = (n as { children?: DocumentNode[] }).children;
    if (c) collect(c, into);
  }
  return into;
};

test("publish, bundle and resolve: a map over derived rows becomes the ledger", async () => {
  assertEquals(branches >= 5, true);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const resolveIds = async (statement: Obj) =>
    collect((await resolveDocument(published, state({ statement: { ...base, ...statement } }), { derivers })).nodes);

  const all = await resolveIds({ entries: many, dueBy: "2026-11-30" });
  const ids = all.map((n) => n.id);
  const entries = ids.filter((id) => id.startsWith("statement-of-account-entry"));
  assertEquals(entries.length, 60);
  assertEquals(new Set(entries).size, 60);
  const text = JSON.stringify(all);
  assertEquals(ids.includes("statement-of-account-owing"), true);
  assertEquals(["clear", "credit", "empty"].some((k) => ids.includes("statement-of-account-" + k)), false);
  assertStringIncludes(text, "Amount due: £600.00, payable by 30 November 2026.");
  assertEquals(text.indexOf("£10.00") < text.indexOf("£600.00"), true);
  const direct = run({ entries: many });
  assertEquals(direct.rows[0].balance, "£10.00");
  assertStringIncludes(text, direct.rows[59].balance);

  const empty = await resolveIds({ openingBalance: 0 });
  const eIds = empty.map((n) => n.id);
  assertEquals(eIds.includes("statement-of-account-ledger"), false);
  assertEquals(
    ["empty-opening", "empty", "clear"].every((k) => eIds.includes("statement-of-account-" + k)),
    true,
  );
  assertEquals(eIds.includes("statement-of-account-owing"), false);

  const credit = await resolveIds({ openingBalance: -40 });
  assertEquals(credit.map((n) => n.id).includes("statement-of-account-credit"), true);
});

test("the catalog's preview data reconciles", () => {
  const entry = COMPONENTS.find((c) => c.id === "statement-of-account")!;
  const o = mod.run([(entry.previewData as { statement: Obj }).statement], {} as RuntimeState) as Out;
  assertEquals([o.totalDebits, o.totalCredits, o.closingBalance, o.dueSuffix], ["£249,278.33", "£150,902.75", "£116,828.33", ", payable by 30 November 2026"]);
  assertEquals(o.rows.map((r) => r.balance), ["£167,379.15", "£148,926.40", "£136,476.40", "£137,762.73", "£17,762.73", "£114,512.73", "£116,828.33"]);
});
