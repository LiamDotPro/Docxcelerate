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
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/statement-remittance-reference/statement-remittance-reference.node.tsx";
const { StatementRemittanceReference } = await loadRegistryModule<
  typeof import("../registry/components/statement-remittance-reference/statement-remittance-reference.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/statement-remittance-reference.deriver.ts")
>("components/derivers/statement-remittance-reference.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ statementRemittanceReference: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Out {
  outcome: string;
  net: string;
  accountNumber: string;
  statementDate: string;
  referenceRule: string | null;
  referenceLabel: string | null;
  referenceLines: string[];
  perItemNote: string | null;
  creditsNote: string | null;
  disputedNote: string | null;
  rows: Array<{ item: string; due: string; amount: string; running: string }>;
  tickInstruction: boolean;
}
const run = (statement: Obj, items?: Obj[]) => mod.run([statement, items], {} as RuntimeState) as Out;
const throws = (statement: Obj, items: Obj[] | undefined, text: string) => {
  let message = "";
  try {
    run(statement, items);
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }
  assertStringIncludes(message, text);
};

const stmt = (rail = "bacs", more: Obj = {}) => ({
  currency: { bacs: "GBP", fps: "GBP", sepa: "EUR", "sepa-rf": "EUR", "ach-ccd": "USD", "ach-ctx": "USD", nz: "NZD", kid: "NOK", "se-ocr": "SEK", swift: "GBP" }[rail],
  date: "2026-11-30",
  customerNumber: "C-004821",
  rail,
  ...more,
});
const inv = (number: string, amount: number, more: Obj = {}) => ({ type: "invoice", number, date: "2026-09-01", amount, ...more });
const A = (): Obj[] => [
  inv("INV-1041", 1200, { date: "2026-09-01", dueDate: "2026-10-01" }),
  inv("INV-1042", 800, { date: "2026-09-15", dueDate: "2026-10-15" }),
  inv("INV-1047", 2500, { date: "2026-10-01", dueDate: "2026-10-31" }),
  { type: "credit-note", number: "CN-0042", date: "2026-10-05", amount: 1500 },
  inv("INV-1050", 640, { date: "2026-10-20", dueDate: "2026-11-19", disputed: true }),
];
const five = () => ["INV-1041", "INV-1042", "INV-1047", "INV-1053", "INV-1060"].map((n) => inv(n, 100));

test("fixture A nets oldest first and picks the list", () => {
  const o = run(stmt(), A());
  assertEquals([o.outcome, o.net, o.referenceRule, o.referenceLines], ["pay", "£3,000.00", "list", ["INV-1042 1047"]]);
  assertEquals(o.referenceLabel, "payment reference");
  assertEquals(o.rows, [
    { item: "INV-1042 (part-paid by credit: £500.00 of £800.00 remains)", due: "15 Oct 2026", amount: "£500.00", running: "£500.00" },
    { item: "INV-1047", due: "31 Oct 2026", amount: "£2,500.00", running: "£3,000.00" },
    { item: "INV-1050 (in dispute, not included)", due: "19 Nov 2026", amount: "£640.00", running: "" },
  ]);
  assertEquals(
    o.creditsNote,
    "Credit note CN-0042 (£1,500.00) is set against your oldest invoices first. It clears INV-1041, and leaves £500.00 to pay on INV-1042.",
  );
  assertEquals(o.disputedNote, "INV-1050 (£640.00) is in dispute and is not included in the amount to pay.");
  assertEquals([o.tickInstruction, o.accountNumber, o.statementDate], [true, "C-004821", "30 November 2026"]);
});

test("ordering ignores input order, ties go by date then numeric number, a missing due date uses the date", () => {
  assertEquals(run(stmt(), [...A()].reverse()), run(stmt(), A()));
  const o = run(stmt(), [
    inv("INV-10", 10, { date: "2026-09-01", dueDate: "2026-10-01" }),
    inv("INV-9", 10, { date: "2026-09-01", dueDate: "2026-10-01" }),
    inv("INV-5", 10, { date: "2026-08-01", dueDate: "2026-10-01" }),
    inv("INV-0", 10, { date: "2026-09-20" }),
  ]);
  assertEquals((o.rows as Array<{ item: string }>).map((r) => r.item), ["INV-0", "INV-5", "INV-9", "INV-10"]);
});

test("disputed items are never counted, and a disputed credit is not set off", () => {
  const items = A().map((i) => (i.number === "INV-1047" ? { ...i, disputed: true } : i));
  const o = run(stmt(), items);
  assertEquals([o.net, o.referenceRule, o.referenceLines], ["£500.00", "item", ["INV-1042"]]);
  const c = run(stmt(), A().map((i) => (i.number === "CN-0042" ? { ...i, disputed: true } : i)));
  assertEquals([c.net, c.creditsNote === null], ["£4,500.00", true]);
});

test("one item quotes its own number, folded, or its own reference", () => {
  assertEquals(run(stmt(), [inv("inv-1041", 10)]).referenceLines, ["INV-1041"]);
  assertEquals(run(stmt(), [inv("inv-1041", 10, { reference: "abc_12/x" })]).referenceLines, ["ABC 12/X"]);
});

test("the reference is folded before it is measured", () => {
  const o = run(stmt(), [inv("A", 10, { reference: "Müller–Söhne 12345" })]);
  assertEquals(o.referenceLines, ["MULLER SOHNE 12345"]);
  assertEquals((o.referenceLines as string[])[0].length, 18);
  // 24 raw characters, 18 once folded: measuring before folding would reject it.
  const wide = run(stmt(), [inv("A", 10, { reference: "Müller  –  Söhne   12345" })]);
  assertEquals(wide.referenceLines, ["MULLER SOHNE 12345"]);
  const nfd = run(stmt(), [inv("A", 10, { reference: "Müller–Söhne 12345".normalize("NFD") })]);
  assertEquals(nfd.referenceLines, ["MULLER SOHNE 12345"]);
});

test("a list too long for the rail falls to the RF or the account tag", () => {
  const o = run(stmt(), five());
  assertEquals([o.referenceRule, o.referenceLines], ["account", ["C-004821 261130"]]);
  const rf = run(stmt("bacs", { acceptsRf: true }), five());
  assertEquals([rf.referenceRule, rf.referenceLines], ["rf", ["RF30C004821261130"]]);
});

test("RF check digits follow ISO 11649", () => {
  const o = run(stmt("sepa-rf", { customerNumber: "539007547034" }), [inv("A", 1), inv("B", 1)]);
  // 539007547034 + 261130 is 18 characters, so the date is part of the base here.
  assertEquals((o.referenceLines as string[])[0].startsWith("RF"), true);
  const one = run(stmt("sepa-rf"), [inv("A", 1, { reference: "RF18539007547034" })]);
  assertEquals(one.referenceLines, ["RF18539007547034"]);
  const bad = run(stmt("sepa-rf", { customerNumber: "539007547034" }), [inv("A", 1, { reference: "RF19539007547034" })]);
  assertEquals(bad.referenceRule, "rf");
  assertEquals((bad.referenceLines as string[])[0] === "RF19539007547034", false);
});

test("generated RF check digits match an independent mod 97 of the base", () => {
  // The base 539007547034 alone cannot be generated (the date always joins a short base), so
  // verify the full string against BigInt arithmetic, and the ISO example as a held reference.
  const o = run(stmt("sepa-rf", { customerNumber: "539007547034", date: "2026-11-30" }), [inv("A", 1), inv("B", 1)]);
  const base = "539007547034261130";
  const check = String(98n - (BigInt(base + "271500") % 97n)).padStart(2, "0");
  assertEquals(o.referenceLines, ["RF" + check + base]);
  const iso = run(stmt("sepa-rf"), [inv("A", 1, { reference: "RF18539007547034" })]);
  assertEquals(iso.referenceLines, ["RF18539007547034"]);
  assertEquals(String(98n - (BigInt("539007547034271500") % 97n)), "18");
});

test("credits part-paying one invoice and clearing none open the note with a subject", () => {
  const two = run(stmt(), [
    inv("I1", 1000),
    { type: "credit-note", number: "CN-A", date: "2026-09-01", amount: 10 },
    { type: "credit-note", number: "CN-B", date: "2026-10-01", amount: 10 },
  ]);
  assertEquals(
    two.creditsNote,
    "Credit note CN-A (£10.00) and credit note CN-B (£10.00) are set against your oldest invoices first. They leave £980.00 to pay on I1.",
  );
  const one = run(stmt(), [inv("I1", 1000), { type: "credit-note", number: "CN-A", date: "2026-09-01", amount: 10 }]);
  assertStringIncludes(one.creditsNote as string, "first. It leaves £990.00 to pay on I1.");
});

test("a credit's dueDate is ignored and credits go by date then number", () => {
  const o = run(stmt(), [
    inv("I1", 1000),
    { type: "credit-note", number: "CN-A", date: "2026-09-01", dueDate: "2026-12-31", amount: 10 },
    { type: "credit-note", number: "CN-B", date: "2026-10-01", amount: 10 },
  ]);
  assertStringIncludes(o.creditsNote as string, "Credit note CN-A (£10.00) and credit note CN-B");
  run(stmt(), [inv("I1", 1000), { type: "payment", number: "P1", date: "2026-09-01", dueDate: "nonsense", amount: 10 }]);
});

test("the swift list counts the trailing comma, so an exact 35 character line wraps", () => {
  const o = run(stmt("swift"), [inv("X-111111111111111", 1), inv("Y-222222222222222", 1), inv("Z-333", 1)]);
  assertEquals([o.referenceRule, o.referenceLines], ["list", ["X-111111111111111,", "Y-222222222222222,Z-333"]]);
});

test("sepa-rf: valid RF, statement RF fallback, and no free text", () => {
  const s = stmt("sepa-rf");
  assertEquals(run(s, [inv("A", 1, { reference: "RF18 5390 0754 7034" })]).referenceLines, ["RF18539007547034"]);
  assertEquals(run(s, [inv("A", 1)]).referenceRule, "rf");
  assertEquals(run(s, [inv("A", 1), inv("B", 1)]).referenceRule, "rf");
  throws({ ...s, customerNumber: "ABCDEFGHIJKLMNOPQRSTUVW" }, [inv("A", 1), inv("B", 1)], "sepa-rf");
});

test("a reference that fits nowhere throws rather than truncating", () => {
  throws(stmt("bacs", { customerNumber: "VERY-LONG-CUSTOMER-NUMBER-1" }), five(), "No reference for this statement fits bacs");
});

test("ACH quotes RMR segments", () => {
  const items = A();
  const ccd = run(stmt("ach-ccd"), items);
  assertEquals(ccd.referenceLines, ["RMR*IV*INV-1042*PI*500.00\\RMR*IV*INV-1047*PI*2500.00\\"]);
  assertEquals(ccd.referenceRule, "list");
  const ctx = run(stmt("ach-ctx"), items);
  assertEquals(ctx.referenceLines, ["RMR*IV*INV-1042*PI*500.00\\", "RMR*IV*INV-1047*PI*2500.00\\"]);
  assertEquals(run(stmt("ach-ccd"), [inv("A*B", 5)]).referenceLines, ["RMR*IV*A B*PI*5.00\\"]);
  assertEquals(ccd.net, "$3,000.00");
});

test("SWIFT wraps at the separator, and five lines fall to the next rule", () => {
  const p = (COMPONENTS.find((c) => c.id === "statement-remittance-reference")!.previewData) as Obj;
  const o = run({ ...(p.statement as Obj), rail: "swift", acceptsRf: false }, p.items as Obj[]);
  assertEquals([o.referenceRule, o.referenceLines], ["list", ["INV-2026-0042177,0042318,0042406,", "0042690"]]);
  const many = Array.from({ length: 30 }, (_, i) => inv("INV-" + (1000 + i * 7), 1));
  const s = run({ ...stmt("swift"), acceptsRf: false }, many);
  assertEquals(s.referenceRule, "account");
});

test("NZ spreads over particulars, code and reference", () => {
  const o = run(stmt("nz"), A());
  assertEquals(o.referenceLines, ["Particulars: C-004821", "Code: INV-1042", "Reference: 1047"]);
  assertEquals(o.referenceLabel, "particulars, code and reference");
  throws(stmt("nz", { customerNumber: "C-004821-LONG-X" }), A(), "Particulars");
  assertEquals(run(stmt("nz"), [inv("INV-1", 1)]).referenceLines, ["Particulars: C-004821", "Reference: INV-1"]);
});

test("KID and OCR take one reference per payment", () => {
  const one = run(stmt("kid", { acceptsRf: true }), [inv("A", 5, { reference: "0010 4200 17" })]);
  assertEquals([one.referenceRule, one.referenceLines, one.referenceLabel], ["item", ["0010420017"], "KID"]);
  const many = run(stmt("kid"), [inv("INV-1042", 5, { reference: "12345" }), inv("INV-1047", 5, { reference: "67890" })]);
  assertEquals([many.referenceRule, many.referenceLines, many.referenceLabel], ["per-item", [], null]);
  assertEquals((many.rows as Array<{ item: string }>)[0].item, "INV-1042, KID 12345");
  assertStringIncludes(many.perItemNote as string, "KID payments carry one reference each");
  const ocr = run(stmt("se-ocr"), [inv("A", 5, { reference: "12345" }), inv("B", 5, { reference: "67890" })]);
  assertStringIncludes((ocr.rows as Array<{ item: string }>)[1].item, "B, OCR 67890");
  assertStringIncludes(ocr.perItemNote as string, "OCR payments");
  throws(stmt("kid"), [inv("A", 5)], "items[0].reference");
  throws(stmt("kid"), [inv("A", 5, { reference: "12ab" })], "items[0].reference");
});

test("zero outcomes", () => {
  const items = [inv("A", 100), inv("B", 50), { type: "payment", number: "P1", date: "2026-09-02", amount: 150 }];
  const z = run(stmt(), items);
  assertEquals([z.outcome, z.net, z.rows, z.referenceLines, z.referenceRule], ["clear", "", [], [], null]);
  assertEquals(z.creditsNote, "Payment P1 (£150.00) is set against your oldest invoices first. It clears A and B.");
  const tiny = run(stmt(), [inv("A", 0.004)]);
  assertEquals([tiny.outcome, tiny.creditsNote, tiny.rows], ["clear", null, []]);
  for (const none of [[], undefined]) {
    const o = run(stmt(), none);
    assertEquals([o.outcome, o.creditsNote, o.disputedNote, o.perItemNote], ["clear", null, null, null]);
  }
});

test("credits over the invoices leave the account in credit", () => {
  const o = run(stmt(), [inv("A", 100), { type: "credit-note", number: "C1", date: "2026-09-02", amount: 130 }]);
  assertEquals([o.outcome, o.net, o.rows, o.referenceLines], ["credit", "£30.00", [], []]);
});

test("invalid input throws, naming the field", () => {
  const ok = inv("A", 1);
  throws(stmt(), [{ ...ok, amount: -1 }], "items[0].amount must be zero or more");
  throws(stmt(), [{ ...ok, amount: NaN }], "items[0].amount");
  throws(stmt(), [ok, { ...ok, type: "refund" }], "items[1].type");
  throws({ ...stmt(), rail: "wire" }, [ok], "statement.rail");
  throws({ ...stmt("sepa"), currency: "GBP" }, [ok], "statement.currency");
  throws({ ...stmt(), currency: "XX" }, [ok], "statement.currency");
  throws({ ...stmt("swift"), currency: "ABC" }, [ok], "statement.currency");
  throws(stmt(), [ok, { ...ok }], "duplicates");
  throws({ ...stmt(), customerNumber: "  " }, [ok], "statement.customerNumber");
  throws({ ...stmt(), date: "30/11/2026" }, [ok], "statement.date");
  throws(stmt(), [{ ...ok, date: "2026-13-01" }], "items[0].date");
  throws(stmt(), [{ ...ok, dueDate: "soon" }], "items[0].dueDate");
  throws(stmt(), [{ ...ok, number: " " }], "items[0].number");
});

test("minor units and the balance", () => {
  assertEquals(run({ ...stmt("swift"), currency: "JPY" }, [inv("A", 1500.4)]).net, "JP¥1,500");
  assertEquals(run({ ...stmt("swift"), currency: "BHD" }, [inv("A", 1.2345)]).net.replace(/\s/g, " "), "BHD 1.235");
  assertEquals(run(stmt(), [inv("A", 1.005)]).net, "£1.01");
  const p = (COMPONENTS.find((c) => c.id === "statement-remittance-reference")!.previewData) as Obj;
  const o = run(p.statement as Obj, p.items as Obj[]);
  assertEquals(o.net, "€2,076,121.85");
  assertEquals((o.rows as Array<{ running: string }>).map((r) => r.running), ["€1,250,400.50", "€1,663,381.25", "€1,736,996.45", "€2,076,121.85", ""]);
  assertEquals([o.referenceRule, o.referenceLines], ["list", ["INV-2026-0042177,0042318,0042406,0042690"]]);
  assertEquals((o.rows as Array<{ amount: string }>)[4].amount, "€258,300.00");
});

test("forty invoices keep a running total equal to the net and never truncate", () => {
  const items = Array.from({ length: 40 }, (_, i) => inv("INV-" + (2000 + i), 10 + i));
  const o = run(stmt("sepa"), items);
  const rows = o.rows as Array<{ running: string }>;
  assertEquals([rows.length, rows[39].running, o.net], [40, "€1,180.00", "€1,180.00"]);
  const text = (o.referenceLines as string[]).join("");
  assertEquals(text.length <= 140, true);
  assertEquals(o.referenceRule === "list" ? text.includes("2039") : text.includes("C-004821"), true);
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Statement">
      <Section id="body" title="Body">
        <StatementRemittanceReference />
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

const texts = (nodes: DocumentNode[]) =>
  collect(nodes).map((n) => (n as { text?: string }).text).filter((t): t is string => typeof t === "string" && t !== "");

const fixtureA = { statement: stmt(), items: A() };

test("a publish build loops over the derived rows and reference lines", async () => {
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  const sources = all.filter((n) => n.kind === "repeat").map((n) => JSON.parse(JSON.stringify((n as unknown as { source: unknown }).source)));
  assertEquals(sources, [
    { scope: "derived", path: "statementRemittanceReference.referenceLines" },
    { scope: "derived", path: "statementRemittanceReference.rows" },
  ]);
  assertEquals(branches, 13);
  const ids = new Set(all.map((n) => n.id));
  for (
    const id of [
      "credit", "clear", "credits", "disputed", "pay", "reference-label", "per-item", "cut", "advice-heading",
      "advice-account", "advice-date", "tick", "advice", "advice-total", "advice-paid", "advice-header",
    ]
  ) assertEquals(ids.has("statement-remittance-" + id), true);
});

test("the published document resolves to what the local build prints", async () => {
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const resolved = await resolveDocument(published, state(fixtureA), { derivers });
  const local = await buildDocument(templateOf(), fixtureA, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);
  const want = texts(local.nodes);
  assertEquals(texts(resolved.nodes), want);
  const joined = want.join("\n");
  assertStringIncludes(joined, "The amount to pay on this statement is £3,000.00.");
  assertStringIncludes(joined, "INV-1042 1047");
  assertStringIncludes(joined, "- - - - - - - - - - - - - - - - - - - -  Detach here  - - - - - - - - - - - - - - - - - - - -");
  assertStringIncludes(joined, "INV-1050 (in dispute, not included)");
});

test("the published conditions pick the right arms for each outcome", async () => {
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const ids = async (statement: Obj, items: Obj[]) => {
    const r = await resolveDocument(published, state({ statement, items }), { derivers });
    return new Set(collect(r.nodes).map((n) => n.id));
  };
  const has = (set: Set<string | undefined>, id: string) => set.has("statement-remittance-" + id);
  const credit = await ids(stmt(), [inv("A", 10), { type: "credit-note", number: "C", date: "2026-09-01", amount: 50 }]);
  assertEquals([has(credit, "credit"), has(credit, "clear"), has(credit, "pay"), has(credit, "tick"), has(credit, "credits")], [true, false, false, false, true]);
  const clear = await ids(stmt(), [inv("A", 10), { type: "credit-note", number: "C", date: "2026-09-01", amount: 10 }]);
  assertEquals([has(clear, "credit"), has(clear, "clear"), has(clear, "pay"), has(clear, "advice-heading")], [false, true, false, false]);
  const per = await ids(stmt("kid"), [inv("A", 10, { reference: "1234" }), inv("B", 10, { reference: "5678" })]);
  assertEquals([has(per, "pay"), has(per, "per-item"), has(per, "reference-label"), has(per, "tick")], [true, true, false, true]);
  const single = await ids(stmt(), [inv("A", 10)]);
  assertEquals([has(single, "pay"), has(single, "per-item"), has(single, "reference-label"), has(single, "tick")], [true, false, true, false]);
});
