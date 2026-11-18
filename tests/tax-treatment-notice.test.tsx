import { test } from "node:test";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import ts from "typescript";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { buildDocument, createDeriverRegistry, InMemoryDataProvider, resolveDocument } from "docxcelerate";
import type { DocumentNode, ParagraphNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import { transformDocumentSource } from "docxcelerate/transform";
import { registryRoot } from "docxcelerate/registry/install";

/**
 * The tax treatment notice, run the way a project runs it: compiled by the
 * document transform, then built against real data and against publish
 * stand-ins. The node lives outside `tests/`, so it is compiled into
 * `dist-tests/` here, where `docxcelerate` still resolves.
 */
const dir = join(registryRoot(), "components", "tax-treatment-notice");
const outDir = join(process.cwd(), "dist-tests", "tax-treatment-notice-build");

async function compile(file: string, transform: boolean) {
  const source = await readFile(join(dir, file), "utf8");
  const input = transform ? transformDocumentSource(source, { fileName: file }) : { code: source, branches: 0 };
  const js = ts.transpileModule(input.code, {
    compilerOptions: {
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.ESNext,
      jsx: ts.JsxEmit.ReactJSX,
      jsxImportSource: "docxcelerate/template",
    },
  }).outputText.replaceAll(".deriver.ts", ".deriver.js");
  const out = join(outDir, file.replace(/\.tsx?$/, ".js"));

  await mkdir(outDir, { recursive: true });
  await writeFile(out, js);

  return { url: pathToFileURL(out).href, branches: input.branches, source };
}

const nodeBuild = await compile("tax-treatment-notice.node.tsx", true);
const derivBuild = await compile("tax-treatment-notice.deriver.ts", false);
const { TaxTreatmentNotice } = await import(nodeBuild.url) as { TaxTreatmentNotice: () => never };
const taxTreatmentNoticeVat = (await import(derivBuild.url)).default as {
  run: (inputs: unknown[], state?: RuntimeState) => string;
};

type Data = Record<string, unknown>;

const full = {
  invoice: {
    taxTreatment: "standard",
    jurisdiction: "FR",
    exportEvidence: "26DE470123456789A3",
    notionalVat: 184250.75,
  },
  seller: { vatNumber: "FR12345678901" },
  customer: {
    vatNumber: "NL123456789B01",
    resaleCertificate: { number: "SR EAA 102-345678", state: "California" },
  },
};

function templateOf() {
  return template<Data>(
    <Document id="invoice" title="Invoice">
      <Section id="body" title="Body">
        <TaxTreatmentNotice />
      </Section>
    </Document>,
  );
}

/** Builds with real data, the way a local build does, and returns the paragraphs. */
async function run(
  invoice: Data,
  parties: { seller?: Data; customer?: Data } = {},
): Promise<ParagraphNode[]> {
  const data = {
    invoice: { ...full.invoice, ...invoice },
    seller: parties.seller ?? full.seller,
    customer: parties.customer ?? full.customer,
  };
  const built = await buildDocument(templateOf(), data, {
    branchMode: "decide",
    dynamicMode: "placeholder",
    derivers: createDeriverRegistry([taxTreatmentNoticeVat as never]),
  } as never);
  const body = built.nodes[0];

  if (body?.kind !== "section") throw new Error("expected the body section");

  return body.children as ParagraphNode[];
}

async function only(invoice: Data, id: string, parties = {}): Promise<string> {
  const nodes = await run(invoice, parties);

  assertEquals(nodes.map((n) => n.id), [id]);

  return nodes[0].text ?? "";
}

const treat = (taxTreatment: string, rest: Data = {}) => ({ taxTreatment, ...rest });

test("standard prints nothing", async () => {
  assertEquals(await run(treat("standard")), []);
});

test("zero-rated export prints its evidence when given, and the mention alone when not", async () => {
  assertEquals(
    await only(treat("zero-rated-export"), "tax-notice-export-evidenced"),
    "Zero-rated export of goods. Export evidence: 26DE470123456789A3.",
  );
  assertEquals(
    await only(treat("zero-rated-export", { exportEvidence: undefined }), "tax-notice-export"),
    "Zero-rated export of goods.",
  );
  await only(treat("zero-rated-export", { exportEvidence: "" }), "tax-notice-export");
});

test("intra-Community supply needs the customer's VAT number", async () => {
  assertStringIncludes(
    await only(treat("intra-community-supply"), "tax-notice-intra-community"),
    "Customer VAT number: NL123456789B01.",
  );
  await only(treat("intra-community-supply"), "tax-notice-intra-community-incomplete", {
    customer: {},
  });
  await only(treat("intra-community-supply"), "tax-notice-intra-community-incomplete", {
    customer: { vatNumber: "" },
  });
});

test("reverse charge needs both VAT numbers and never says 0%", async () => {
  const text = await only(treat("reverse-charge-services"), "tax-notice-reverse-charge");

  assertEquals(
    text,
    "Reverse charge (Art 196 Directive 2006/112/EC). Supplier VAT number: FR12345678901. Customer VAT number: NL123456789B01.",
  );
  assertEquals(text.includes("0%"), false);

  for (const parties of [{ seller: {} }, { customer: {} }, { seller: {}, customer: {} }]) {
    const incomplete = await only(treat("reverse-charge-services"), "tax-notice-reverse-charge-incomplete", parties);

    assertEquals(incomplete.includes("0%"), false);
  }
});

test("UK domestic reverse charge states the VAT in sterling, signed", async () => {
  assertStringIncludes(
    await only(treat("uk-domestic-reverse-charge"), "tax-notice-uk-reverse-charge"),
    "VAT that would have been due: £184,250.75.",
  );
  assertStringIncludes(
    await only(treat("uk-domestic-reverse-charge", { notionalVat: -2450.5 }), "tax-notice-uk-reverse-charge"),
    "-£2,450.50",
  );
  await only(treat("uk-domestic-reverse-charge", { notionalVat: 0 }), "tax-notice-uk-reverse-charge-incomplete");
  await only(
    treat("uk-domestic-reverse-charge", { notionalVat: undefined }),
    "tax-notice-uk-reverse-charge-incomplete",
  );
});

test("small-business wording is exact, and only for FR and DE", async () => {
  assertEquals(
    await only(treat("small-business-exempt", { jurisdiction: "FR" }), "tax-notice-small-business-fr"),
    "TVA non applicable, art. 293 B du CGI",
  );
  assertEquals(
    await only(treat("small-business-exempt", { jurisdiction: "DE" }), "tax-notice-small-business-de"),
    "Gemäß § 19 UStG wird keine Umsatzsteuer berechnet",
  );

  for (const jurisdiction of ["fr", "BE", undefined]) {
    await only(treat("small-business-exempt", { jurisdiction }), "tax-notice-small-business-incomplete");
  }
});

test("a resale sale needs the certificate number and the state", async () => {
  assertEquals(
    await only(treat("us-resale-exempt"), "tax-notice-resale"),
    "Exempt: sale for resale. Resale certificate SR EAA 102-345678, issued in California.",
  );

  for (const resaleCertificate of [undefined, { number: "", state: "CA" }, { number: "X1", state: "" }]) {
    await only(treat("us-resale-exempt"), "tax-notice-resale-incomplete", {
      customer: { vatNumber: "NL123456789B01", resaleCertificate },
    });
  }
});

test("the margin scheme prints no VAT figure", async () => {
  const text = await only(treat("margin-scheme", { notionalVat: 500 }), "tax-notice-margin-scheme");

  assertEquals(text, "Margin scheme — Second-hand goods");
  assertEquals(text.includes("£"), false);
});

test("a treatment outside the union warns rather than printing nothing", async () => {
  const text = await only(treat("reverse_charge"), "tax-notice-unrecognised");

  assertStringIncludes(text, "not one this notice recognises");
});

// ---------------------------------------------------------------------------
// The deriver
// ---------------------------------------------------------------------------

test("the deriver formats sterling, rounding half away from zero", () => {
  const run = (x: unknown) => taxTreatmentNoticeVat.run([x]);

  assertEquals(run(184250.75), "£184,250.75");
  assertEquals(run(-2450.5), "-£2,450.50");
  assertEquals(run(0.005), "£0.01");
  assertEquals(run("1200"), "£1,200.00");

  assertEquals(run(undefined), "");
  assertEquals(run(null), "");

  for (const bad of [Number.NaN, "abc"]) {
    try {
      run(bad);
    } catch (error) {
      assertStringIncludes((error as Error).message, "taxTreatmentNoticeVat");
      continue;
    }
    throw new Error(`expected ${String(bad)} to throw`);
  }
});

// ---------------------------------------------------------------------------
// Publishing
// ---------------------------------------------------------------------------

test("a publish build carries all fifteen arms, each with a condition", async () => {
  assertEquals(nodeBuild.branches, 15);

  const published = await buildDocument(templateOf(), createPublishData() as Data, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const found = new Map<string, DocumentNode>();
  const walk = (nodes: DocumentNode[]) => {
    for (const node of nodes) {
      found.set(node.id, node);
      const children = (node as { children?: DocumentNode[] }).children;
      if (children) walk(children);
    }
  };

  walk(published.nodes);

  const ids = [
    "export-evidenced", "export", "intra-community", "intra-community-incomplete", "reverse-charge",
    "reverse-charge-incomplete", "uk-reverse-charge", "uk-reverse-charge-incomplete", "small-business-fr",
    "small-business-de", "small-business-incomplete", "resale", "resale-incomplete", "margin-scheme",
    "unrecognised",
  ].map((id) => `tax-notice-${id}`);

  for (const id of ids) {
    const node = found.get(id) as (DocumentNode & { when?: unknown }) | undefined;

    if (!node) throw new Error(`missing ${id}`);
    if (node.when === undefined) throw new Error(`${id} has no condition`);
  }

  assertEquals(JSON.stringify(found.get("tax-notice-uk-reverse-charge")).includes("taxTreatmentNoticeVat"), true);

  // Resolved against real data, an engine settles to the one arm that applies.
  const data = { ...full, invoice: { ...full.invoice, taxTreatment: "uk-domestic-reverse-charge" } };
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data as Record<string, unknown>),
    aiClient: { generateParagraph: () => "" },
  };
  const resolved = await resolveDocument(published, state, {
    derivers: createDeriverRegistry([taxTreatmentNoticeVat as never]),
  });
  const body = resolved.nodes[0] as { children: ParagraphNode[] };

  assertEquals(body.children.map((n) => n.id), ["tax-notice-uk-reverse-charge"]);
  assertStringIncludes(body.children[0].text ?? "", "£184,250.75.");

  // No notionalVat: the deriver still runs, and must not throw.
  const resolveWith = async (taxTreatment: string) => {
    const sparse = { ...full, invoice: { taxTreatment } };
    const r = await resolveDocument(published, {
      ctx: { ...sparse },
      derived: {},
      dataProvider: new InMemoryDataProvider(sparse as Record<string, unknown>),
      aiClient: { generateParagraph: () => "" },
    }, { derivers: createDeriverRegistry([taxTreatmentNoticeVat as never]) });

    return (r.nodes[0] as { children: ParagraphNode[] }).children.map((n) => n.id);
  };

  assertEquals(await resolveWith("standard"), []);
  assertEquals(await resolveWith("uk-domestic-reverse-charge"), ["tax-notice-uk-reverse-charge-incomplete"]);

  await assertRejects(
    () => resolveDocument(published, state, { derivers: createDeriverRegistry() }),
    Error,
    "taxTreatmentNoticeVat",
  );
});
