import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import totalsDeriver from "../registry/components/derivers/invoice-totals.deriver.ts";
import type { InvoiceTotals, InvoiceTotalsData } from "../registry/components/derivers/invoice-totals.deriver.ts";
import { InvoiceTotals as InvoiceTotalsNode } from "../registry/components/invoice-totals/invoice-totals.node.tsx";
import { COMPONENTS } from "docxcelerate/registry";
import { transformDocumentSource } from "docxcelerate/transform";
import { readFile, writeFile } from "node:fs/promises";
import ts from "typescript";

/**
 * The invoice arithmetic, held to the penny.
 *
 * The deriver is the only computation, so most of this calls it directly and
 * asserts on the integer `*Minor` fields rather than on Intl output.
 */
type Invoice = InvoiceTotalsData["invoice"];
type Line = Invoice["lines"][number];

const line = (over: Partial<Line> = {}): Line => ({
  description: "Work",
  quantity: 1,
  unitPrice: 100,
  taxRate: 20,
  ...over,
});

const calc = (invoice: Invoice): InvoiceTotals =>
  totalsDeriver.run([invoice], {} as RuntimeState) as InvoiceTotals;

const gbp = (lines: Line[], pricing?: "exclusive" | "inclusive") =>
  calc({ currency: "GBP", pricing, lines });

const throwsNaming = (invoice: Invoice, path: string) => {
  let message = "";

  try {
    calc(invoice);
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, path);
};

test("an invoice with no lines, or lines missing, is empty", () => {
  assertEquals(gbp([]).empty, true);
  assertEquals(calc({ currency: "GBP" } as Invoice).empty, true);
});

test("one 20% line draws no analysis and labels the VAT row with its rate", () => {
  const t = gbp([line()]);

  assertEquals([t.multipleBands, t.vatLabel, t.vat, t.total], [false, "VAT at 20%", "£20.00", "£120.00"]);
});

test("VAT is worked on the band, not per line", () => {
  assertEquals(gbp([line({ unitPrice: 0.03 }), line({ unitPrice: 0.03 }), line({ unitPrice: 0.03 })]).vatMinor, 2);
});

test("inclusive prices have VAT backed out of the band's gross", () => {
  const t = gbp([line({ unitPrice: 1 }), line({ unitPrice: 1 }), line({ unitPrice: 1 })], "inclusive");

  assertEquals([t.vatMinor, t.subtotalMinor, t.totalMinor], [50, 250, 300]);
  assertEquals(t.subtotalLabel, "Total excluding VAT");
  assertEquals(t.amountHeading, "Amount (incl. VAT)");

  const ten = gbp([line({ unitPrice: 10 })], "inclusive");

  assertEquals([ten.vatMinor, ten.subtotalMinor, ten.totalMinor], [167, 833, 1000]);
});

test("a line is rounded once, half away from zero", () => {
  const t = gbp([line({ quantity: 7.25, unitPrice: 142.5 })]);

  assertEquals(t.lines[0].amount, "£1,033.13");
  assertEquals(gbp([line({ unitPrice: -0.005, taxRate: null })]).subtotalMinor, -1);
});

test("a four-decimal unit price is exact in the sum and printed in full", () => {
  const t = gbp([line({ quantity: 18432.5, unitPrice: 0.2847, taxRate: null })]);

  assertEquals(t.subtotalMinor, 524773);
  assertEquals(t.lines[0].unitPrice, "£0.2847");
});

test("yen print no decimals", () => {
  const t = calc({ currency: "JPY", lines: [line({ quantity: 3, unitPrice: 333.3, taxRate: 10 })] });

  assertEquals([t.lines[0].amount, t.vatMinor], ["JP¥1,000", 100]);
});

test("dinars print three decimals", () => {
  const t = calc({ currency: "BHD", lines: [line({ quantity: 3, unitPrice: 1.2345, taxRate: 10 })] });

  assertEquals([t.subtotalMinor, t.vatMinor], [3704, 370]);
  assertStringIncludes(t.vat, "0.370");
});

test("a discount reduces only its own band", () => {
  const t = gbp([line({ unitPrice: 100 }), line({ unitPrice: 50, taxRate: 5 }), line({ unitPrice: -10 })]);

  assertEquals(t.bands.map((b) => [b.label, b.netMinor, b.vatMinor]), [["20%", 9000, 1800], ["5%", 5000, 250]]);
});

test("a band whose only line is negative nets negative, with negative VAT", () => {
  const t = gbp([line({ unitPrice: 100, taxRate: 5 }), line({ unitPrice: -10 })]);

  assertEquals(t.bands.find((b) => b.label === "20%")?.netMinor, -1000);
  assertEquals(t.bands.find((b) => b.label === "20%")?.vatMinor, -200);
});

test("a credit prints a negative total and keeps the same shape", () => {
  const t = gbp([line({ unitPrice: -50 })]);

  assertEquals([t.total, t.empty, t.multipleBands], ["-£60.00", false, false]);
});

test("a zero quantity is still a line", () => {
  const t = gbp([line({ quantity: 0 })]);

  assertEquals([t.empty, t.lines[0].amount], [false, "£0.00"]);
});

test("one exempt line labels the VAT row with its band and draws a dash", () => {
  const t = gbp([line({ taxRate: null, taxLabel: "Exempt" })]);

  assertEquals([t.multipleBands, t.vatLabel, t.vat, t.totalMinor, t.subtotalMinor], [false, "Exempt", "—", 10000, 10000]);
});

test("zero-rated and unlabelled-null are different bands", () => {
  const t = gbp([line({ taxRate: 0 }), line({ taxRate: null })]);

  assertEquals(t.bands.map((b) => b.label), ["0%", "No VAT"]);
  assertEquals(t.multipleBands, true);
});

test("null-rate lines with different labels are separate bands", () => {
  const t = gbp([line({ taxRate: null, taxLabel: "Exempt" }), line({ taxRate: null, taxLabel: "Outside the scope of VAT" })]);

  assertEquals(t.bands.map((b) => b.label), ["Exempt", "Outside the scope of VAT"]);
});

test("a label on a rated line is ignored", () => {
  assertEquals(gbp([line({ taxLabel: "Exempt" })]).lines[0].tax, "20%");
});

test("a fractional rate is exact", () => {
  const t = gbp([line({ taxRate: 12.5, unitPrice: 33.33 })]);

  assertEquals([t.vatMinor, t.vatLabel], [417, "VAT at 12.5%"]);
});

test("bands run highest rate first, then null bands in order seen", () => {
  const t = gbp([line({ taxRate: 0 }), line({ taxRate: null, taxLabel: "B" }), line({ taxRate: 20 }), line({ taxRate: null, taxLabel: "A" }), line({ taxRate: 5 })]);

  assertEquals(t.bands.map((b) => b.label), ["20%", "5%", "0%", "B", "A"]);
});

test("pricing missing is exclusive, and a missing unit prints the quantity bare", () => {
  const t = gbp([line({ quantity: 2 }), line({ quantity: 1.25, unit: "h" })]);

  assertEquals([t.amountHeading, t.lines[0].quantity, t.lines[1].quantity], ["Amount (excl. VAT)", "2", "1.25 h"]);
});

test("thirty lines keep their input order", () => {
  const lines = Array.from({ length: 30 }, (_, i) => line({ description: `L${i}`, taxRate: i % 2 ? 5 : 20 }));

  assertEquals(calc({ currency: "GBP", lines }).lines.map((l) => l.description), lines.map((l) => l.description));
});

test("bad input throws, naming the path", () => {
  throwsNaming({ currency: "ZZZZ", lines: [] }, "invoice.currency");
  throwsNaming({ currency: undefined as unknown as string, lines: [] }, "invoice.currency");
  throwsNaming({ currency: "GBP", pricing: "gross" as "inclusive", lines: [] }, "invoice.pricing");
  throwsNaming({ currency: "GBP", lines: [line(), line({ taxRate: undefined as unknown as number })] }, "invoice.lines[1].taxRate");
  throwsNaming({ currency: "GBP", lines: [line({ taxRate: -1 })] }, "invoice.lines[0].taxRate");
  throwsNaming({ currency: "GBP", lines: [line({ quantity: NaN })] }, "invoice.lines[0].quantity");
  throwsNaming({ currency: "GBP", lines: [line({ unitPrice: Infinity })] }, "invoice.lines[0].unitPrice");
});

test("the catalog preview data produces the figures the brief checked", () => {
  const entry = COMPONENTS.find((c) => c.id === "invoice-totals")!;
  const t = calc((entry.previewData as unknown as InvoiceTotalsData).invoice);

  assertEquals(t.lines.map((l) => l.amount), ["£51,400.00", "£1,033.13", "£3,487.60", "£5,247.73", "£291.00", "£2,310.00", "-£2,570.00"]);
  assertEquals(t.bands.map((b) => [b.label, b.netMinor, b.vatMinor]), [
    ["20%", 5335073, 1067015],
    ["5%", 524773, 26239],
    ["0%", 29100, 0],
    ["Exempt", 231000, 0],
  ]);
  assertEquals([t.subtotalMinor, t.vatMinor, t.totalMinor], [6119946, 1093254, 7213200]);
});

// ---------------------------------------------------------------------------
// The node
// ---------------------------------------------------------------------------

const doc = template<InvoiceTotalsData>(
  <Document title="Invoice">
    <Section title="Body">
      <InvoiceTotalsNode />
    </Section>
  </Document>,
);

const ids = (nodes: DocumentNode[]): string[] =>
  nodes.flatMap((n) => [n.id ?? "", ..."children" in n ? ids(n.children) : []]).filter((id) => id.startsWith("invoice-totals"));

test("an empty invoice renders only the no-charges paragraph", async () => {
  const built = await buildDocument(doc, { invoice: { currency: "GBP", lines: [] } }, { dynamicMode: "placeholder" });

  assertEquals(ids(built.nodes), ["invoice-totals", "invoice-totals-none"]);
});

test("one band renders lines and summary, several add the analysis", async () => {
  const one = await buildDocument(doc, { invoice: { currency: "GBP", lines: [line()] } }, { dynamicMode: "placeholder" });
  const many = await buildDocument(doc, { invoice: { currency: "GBP", lines: [line(), line({ taxRate: 5 })] } }, { dynamicMode: "placeholder" });

  assertEquals(ids(one.nodes), ["invoice-totals", "invoice-totals-lines", "invoice-totals-summary", "invoice-totals-subtotal", "invoice-totals-vat", "invoice-totals-total"]);
  assertEquals(ids(many.nodes).includes("invoice-totals-analysis"), true);
});

// A publish build needs the compiler's `&&` rewrite, which `tsc` does not apply, so
// the node's own source is compiled the way a project's build would and imported.
async function compiledNode(): Promise<typeof InvoiceTotalsNode> {
  const dir = new URL("../registry/components/invoice-totals/", import.meta.url);
  const source = await readFile(new URL("../../registry/components/invoice-totals/invoice-totals.node.tsx", import.meta.url), "utf8");
  const { code } = transformDocumentSource(source, { fileName: "invoice-totals.node.tsx" });
  const js = ts.transpileModule(code, {
    compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ESNext, jsx: ts.JsxEmit.ReactJSX, jsxImportSource: "docxcelerate/template", verbatimModuleSyntax: true },
  }).outputText.replace(".deriver.ts", ".deriver.js");
  const out = new URL("invoice-totals.compiled.js", dir);

  await writeFile(out, js);

  return (await import(out.href)).InvoiceTotals;
}

test("a publish build carries the invocation, with conditions and loops on derived", async () => {
  const Compiled = await compiledNode();
  const published = template<InvoiceTotalsData>(
    <Document title="Invoice">
      <Section title="Body">
        <Compiled />
      </Section>
    </Document>,
  );
  const built = await buildDocument(published, createPublishData() as InvoiceTotalsData, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const body = built.nodes[0] as DocumentNode & { children: DocumentNode[] };
  const section = body.children[0] as DocumentNode & { derivers?: unknown; children: DocumentNode[] };
  const json = JSON.stringify(section.children);

  assertEquals(section.derivers, [{
    name: "invoiceTotals",
    output: "invoiceTotals",
    inputs: [{ type: "ref", ref: { scope: "data", path: "invoice" } }],
  }]);
  assertStringIncludes(json, '"scope":"derived","path":"invoiceTotals.empty"');
  assertStringIncludes(json, '"scope":"derived","path":"invoiceTotals.multipleBands"');
  assertStringIncludes(json, '"kind":"repeat","source":{"scope":"derived","path":"invoiceTotals.lines"}');
});
