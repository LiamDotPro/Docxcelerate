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
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/structured-payment-reference.deriver.ts")
>("components/derivers/structured-payment-reference.deriver.ts");
const { StructuredPaymentReference } = await loadRegistryModule<
  typeof import("../registry/components/structured-payment-reference/structured-payment-reference.node.tsx")
>("components/structured-payment-reference/structured-payment-reference.node.tsx");
const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ structuredPaymentReference: mod.run }),
);

const FALLBACK = "Please quote invoice number INV-1 with your payment.";
const INV = { number: "INV-1" };
const issued = (x: string) => ({ issued: x, invalid: null, unissuable: null });
const INVALID = { issued: null, invalid: FALLBACK, unissuable: null };
const UNISSUABLE = { issued: null, invalid: null, unissuable: FALLBACK };
// deno-lint-ignore no-explicit-any
const run = (payment: any, invoice: any = INV) =>
  // Async so a throw is a rejection, as it is through the engine.
  Promise.resolve().then(() => (mod.run as unknown as (inputs: unknown[]) => unknown)([payment, invoice])) as Promise<
    { issued: string | null; invalid: string | null; unissuable: string | null }
  >;
const gen = (scheme: string, base: string) => run({ scheme, base });
const val = (scheme: string, reference: string) => run({ scheme, reference });

test("RF", async () => {
  assertEquals(await gen("rf", "539007547034"), issued("Creditor reference (RF): RF18 5390 0754 7034"));
  assertEquals(await gen("rf", "inv 2026 4817"), issued("Creditor reference (RF): RF61 INV2 0264 817"));
  assertEquals(
    await gen("rf", "ABCDEFGHIJ0123456789Z"),
    issued("Creditor reference (RF): RF48 ABCD EFGH IJ01 2345 6789 Z"),
  );
  assertEquals(await gen("rf", "ABCDEFGHIJ0123456789ZZ"), INVALID);
  assertEquals(await gen("rf", "0"), issued("Creditor reference (RF): RF04 0"));
  assertEquals(await gen("rf", "INV-4817"), INVALID);
  assertEquals(await val("rf", "rf18 5390 0754 7034"), issued("Creditor reference (RF): RF18 5390 0754 7034"));
  assertEquals(await val("rf", "RF19539007547034"), INVALID);
});

test("Norwegian KID", async () => {
  assertEquals(await gen("no-kid-mod10", "1234567"), issued("KID: 12345674"));
  assertEquals(await gen("no-kid-mod11", "1234567"), issued("KID: 12345674"));
  assertEquals(await gen("no-kid-mod11", "12345678"), issued("KID: 123456785"));
  assertEquals(await gen("no-kid-mod11", "0000001"), issued("KID: 00000019"));
  assertEquals(await gen("no-kid-mod11", "0014"), issued("KID: 00140"));
  assertEquals(await gen("no-kid-mod11", "0006"), UNISSUABLE);
  assertEquals(await val("no-kid-mod11", "00006-"), UNISSUABLE);
  assertEquals(await val("no-kid-mod11", "00061"), INVALID);
  assertEquals(await val("no-kid-mod11", "1234567-"), INVALID);
  assertEquals(await val("no-kid-mod10", "12345674"), issued("KID: 12345674"));
  assertEquals(await val("no-kid-mod10", "12345675"), INVALID);
  assertEquals((await gen("no-kid-mod10", "1".repeat(24))).issued?.length, "KID: ".length + 25);
  assertEquals(await gen("no-kid-mod10", "1".repeat(25)), INVALID);
});

test("Finnish", async () => {
  const L = "Reference number (viitenumero): ";
  assertEquals(await gen("fi", "123"), issued(L + "1232"));
  assertEquals(await gen("fi", "0000123"), issued(L + "1232"));
  assertEquals(await gen("fi", "1234561"), issued(L + "123 45614"));
  assertEquals(await gen("fi", "1234567890123456789"), issued(L + "12345 67890 12345 67894"));
  assertEquals(await gen("fi", "12345678901234567890"), INVALID);
  assertEquals(await gen("fi", "12"), INVALID);
  assertEquals(await gen("fi", "000"), INVALID);
  assertEquals(await val("fi", "123 45614"), issued(L + "123 45614"));
  assertEquals(await val("fi", "123 45615"), INVALID);
  assertEquals(await gen("fi-rf", "123"), issued("Creditor reference (RF): RF11 1232"));
  assertEquals(await gen("fi-rf", "1234567"), issued("Creditor reference (RF): RF85 1234 5672"));
  assertEquals(await val("fi-rf", "RF85 1234 5672"), issued("Creditor reference (RF): RF85 1234 5672"));
  const rf = (await gen("rf", "12345673")).issued!.replace("Creditor reference (RF): ", "");
  assertEquals((await val("rf", rf)).issued !== null, true);
  assertEquals(await val("fi-rf", rf), INVALID);
});

test("Belgian OGM", async () => {
  const L = "Structured communication (OGM/VCS): ";
  assertEquals(await gen("be-ogm", "1234567890"), issued(L + "+++123/4567/89002+++"));
  assertEquals(await gen("be-ogm", "0909337554"), issued(L + "+++090/9337/55493+++"));
  assertEquals(await gen("be-ogm", "0000000097"), issued(L + "+++000/0000/09797+++"));
  assertEquals(await gen("be-ogm", "1"), issued(L + "+++000/0000/00101+++"));
  assertEquals(await gen("be-ogm", "12345678901"), INVALID);
  assertEquals(await val("be-ogm", "+++090/9337/55493+++"), issued(L + "+++090/9337/55493+++"));
  assertEquals(await val("be-ogm", "+++090/9337/55494+++"), INVALID);
});

test("Swiss QRR", async () => {
  assertEquals(await gen("ch-qrr", "21000000000313947143000901"), issued("QR reference: 21 00000 00003 13947 14300 09017"));
  assertEquals(await gen("ch-qrr", "1234567"), issued("QR reference: 00 00000 00000 00000 00123 45676"));
  assertEquals(await gen("ch-qrr", "12345678901234567890123456"), issued("QR reference: 12 34567 89012 34567 89012 34567"));
  assertEquals(await gen("ch-qrr", "1".repeat(27)), INVALID);
  assertEquals(await val("ch-qrr", "21 00000 00003 13947 14300 09017"), issued("QR reference: 21 00000 00003 13947 14300 09017"));
  assertEquals(await val("ch-qrr", "2100000000031394714300090"), INVALID);
  assertEquals(await val("ch-qrr", "21000000000313947143000901"), INVALID);
  assertEquals(await val("ch-qrr", "210000000003139471430009018"), INVALID);
});

test("Swedish OCR", async () => {
  assertEquals(await gen("se-ocr", "1234567"), issued("OCR number: 12345674"));
  assertEquals(await gen("se-ocr", "123"), issued("OCR number: 1230"));
  assertEquals(await gen("se-ocr-length", "1234567"), issued("OCR number: 123456790"));
  assertEquals(await gen("se-ocr-length", "123"), issued("OCR number: 12351"));
  assertEquals(await gen("se-ocr-length", "12345678"), issued("OCR number: 1234567806"));
  assertEquals(await val("se-ocr-length", "123456790"), issued("OCR number: 123456790"));
  assertEquals(await val("se-ocr-length", "123456780"), INVALID);
  assertEquals(await val("se-ocr-length", "123456782"), INVALID);
  assertEquals(await val("se-ocr", "12345674"), issued("OCR number: 12345674"));
  assertEquals(await val("se-ocr", "12345675"), INVALID);
});

test("missing data, precedence and throws", async () => {
  assertEquals(await run(undefined), INVALID);
  assertEquals(await run(null), INVALID);
  assertEquals(await run({ scheme: "rf" }), INVALID);
  assertEquals(await gen("rf", "   "), INVALID);
  assertEquals(
    await run({ scheme: "rf", reference: "RF18539007547034", base: "zzz" }),
    issued("Creditor reference (RF): RF18 5390 0754 7034"),
  );
  await assertRejects(() => run({ scheme: "rf", base: 1234567 }), Error, "payment.base must be a string");
  await assertRejects(() => run({ scheme: "rf", reference: 123 }), Error, "payment.reference must be a string");
  await assertRejects(() => run({ scheme: "iban", base: "1" }), Error, "payment.scheme");
  await assertRejects(() => run({ scheme: "rf", base: "1" }, {}), Error, "invoice.number is required");
  await assertRejects(() => run({ scheme: "rf", reference: "RF18539007547034" }, { number: "  " }), Error, "invoice.number is required");
  assertEquals((await run(undefined, { number: "  SI-1  " })).invalid, "Please quote invoice number SI-1 with your payment.");
});

// The node, published and resolved by an engine that only has the bundle.
type Data = { payment: Record<string, string>; invoice: { number: string } };
const documentOf = () =>
  template(
    <Document title="Payment">
      <StructuredPaymentReference />
    </Document>,
  );

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; text?: string; children?: DocumentNode[] };
    return [...(own.text ? [`${own.id}: ${own.text}`] : []), ...texts(own.children ?? [])];
  });
}

const cases: [string, Data, string][] = [
  ["issued", { payment: { scheme: "ch-qrr", base: "21000000000313947143000901" }, invoice: INV }, "structured-reference-issued"],
  ["invalid", { payment: { scheme: "be-ogm", base: "12345678901" }, invoice: INV }, "structured-reference-invalid"],
  ["unissuable", { payment: { scheme: "no-kid-mod11", base: "0006" }, invoice: INV }, "structured-reference-unissuable"],
];

for (const [name, data, id] of cases) {
  test(`${name}: publishing then resolving matches building directly`, async () => {
    const published = await buildDocument(documentOf(), createPublishData() as Data, {
      branchMode: "publish",
      deriverMode: "preserve",
      dynamicMode: "resolve",
      aiClient: { generateParagraph: () => "" },
    });
    const state: RuntimeState = {
      ctx: { ...data },
      derived: {},
      dataProvider: new InMemoryDataProvider(data),
      aiClient: { generateParagraph: () => "" },
    };
    const resolved = await resolveDocument(published, state, { derivers: registry });
    const direct = await buildDocument(documentOf(), data, { dynamicMode: "placeholder" });
    assertEquals(texts(resolved.nodes), texts(direct.nodes));
    assertEquals(texts(direct.nodes).length, 1);
    assertEquals(texts(direct.nodes)[0].startsWith(id + ":"), true);
  });
}
