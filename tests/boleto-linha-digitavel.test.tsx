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
  typeof import("../registry/components/derivers/boleto-linha-digitavel.deriver.ts")
>("components/derivers/boleto-linha-digitavel.deriver.ts");
const { BoletoLinhaDigitavel } = await loadRegistryModule<
  typeof import("../registry/components/boleto-linha-digitavel/boleto-linha-digitavel.node.tsx")
>("components/boleto-linha-digitavel/boleto-linha-digitavel.node.tsx");

type Obj = Record<string, unknown>;
const run = (boleto: Obj, invoice: Obj) => mod.run([boleto, invoice], {} as RuntimeState) as Obj;
const rejects = (boleto: Obj, invoice: Obj, message: string) =>
  assertRejects(() => Promise.resolve().then(() => run(boleto, invoice)), Error, message);

const NB = "R$ ";
const bb = { bankCode: "001", freeField: "0000002656973019362470618" };
const bbInvoice = { total: 2434.79, dueDate: "2016-01-20", currency: "BRL" };
const bbLine = "00190.00009 02656.973019 93624.706185 1 66790000243479";
const rollover = { bankCode: "001", freeField: "0000002345678000012345617" };
const at = (dueDate: string | undefined) => run(rollover, { total: 250, dueDate, currency: "BRL" });
const withheldNull = { line: null, barcode: null, dueDate: null, amount: null };

test("1. builds from the preview data", () => {
  assertEquals(
    run(
      { bankCode: "341", freeField: "1090000012345678901234561" },
      { total: 8750432.19, dueDate: "2026-12-15", currency: "BRL" },
    ),
    {
      line: "34191.09008 00012.345674 89012.345610 4 16610875043219",
      barcode: "34194166108750432191090000012345678901234561",
      dueDate: "15/12/2026",
      amount: NB + "8.750.432,19",
      withheld: null,
    },
  );
});

test("2. the published vector, first factor cycle, with its line given", () => {
  const built = run(bb, bbInvoice);
  assertEquals(built.barcode, "00191667900002434790000002656973019362470618");
  assertEquals(built.line, bbLine);
  assertEquals(run({ ...bb, line: bbLine }, bbInvoice), built);
  assertEquals(run({ ...bb, line: "00190000090265697301993624706185166790000243479" }, bbInvoice), built);
});

test("3. the factor rolls over on 2025-02-22", () => {
  assertEquals(at("2025-02-21").line, "00190.00009 02345.678003 00123.456170 3 99990000025000");
  assertEquals(at("2025-02-22").line, "00190.00009 02345.678003 00123.456170 9 10000000025000");
  const factor = (d: string) => (at(d).line as string).slice(-14, -10);
  assertEquals(factor("2025-02-23"), "1001");
  assertEquals(factor("2000-07-03"), "1000");
  assertEquals(factor("2049-10-13"), "9999");
  assertEquals(factor("2049-10-14"), "1000");
});

test("4. no due date", () => {
  for (const dueDate of [undefined, ""]) {
    const r = run(
      { bankCode: "237", freeField: "3381286008227762063000063" },
      { total: 100, dueDate, currency: "BRL" },
    );
    assertEquals(r.line, "23793.38128 86008.227760 20630.000634 6 00000000010000");
    assertEquals(r.dueDate, "Contra apresentação");
  }
});

test("5. rounding is half up on the written value", () => {
  const r = run(rollover, { total: 1.005, dueDate: "2026-11-30", currency: "BRL" });
  assertEquals(r.line, "00190.00009 02345.678003 00123.456170 4 16460000000101");
  assertEquals(r.amount, NB + "1,01");
});

test("6. maximum", () => {
  const r = run(rollover, { total: 99999999.99, currency: "BRL" });
  assertEquals((r.line as string).slice(-10), "9999999999");
  assertEquals(r.amount, NB + "99.999.999,99");
  assertEquals(String(run(rollover, { total: 100000000, currency: "BRL" }).withheld).includes("99.999.999,99"), true);
});

test("7. zero, negative, tiny, and not a number", () => {
  for (const total of [0, -50, 0.004]) {
    const r = run(rollover, { total, currency: "BRL" });
    assertEquals(String(r.withheld).includes("não há valor a pagar"), true);
  }
  for (const total of [undefined, NaN]) {
    const r = run(rollover, { total, currency: "BRL" });
    assertEquals(String(r.withheld).includes("o total da fatura não é um valor"), true);
  }
});

test("8. wrong lengths are withheld with their reason", () => {
  const inv = { total: 10, currency: "BRL" };
  for (
    const [boleto, reason] of [
      [{ ...rollover, freeField: "000000234567800001234561" }, "o campo livre não tem 25 dígitos"],
      [{ ...rollover, freeField: "00000023456.8000012345617" }, "o campo livre não tem 25 dígitos"],
      [{ ...rollover, bankCode: "34" }, "o código do banco não tem 3 dígitos"],
    ] as const
  ) {
    const r = run(boleto, inv);
    assertEquals(r.withheld, "Linha digitável indisponível: " + reason + ". Solicite o boleto ao beneficiário.");
    assertEquals({ ...r, withheld: null }, { ...withheldNull, withheld: null });
  }
});

test("9. the currency is checked first", () => {
  for (const currency of ["USD", undefined]) {
    const r = run({ ...rollover, freeField: "1" }, { total: 10, currency });
    assertEquals(String(r.withheld).includes("a fatura não está em reais (BRL)"), true);
  }
});

test("10. the due-date window", () => {
  const r = at("2000-07-02");
  assertEquals(String(r.withheld).includes("anterior a 03/07/2000"), true);
  assertEquals(r.line === null, true);
  assertEquals(at("2000-07-03").dueDate, "03/07/2000");
});

test("11. malformed dates throw, even when something else would withhold", async () => {
  for (const dueDate of ["2026-02-30", "15/12/2026"]) {
    await rejects(rollover, { total: 10, dueDate, currency: "BRL" }, "invoice.dueDate must be an ISO 8601 date");
    await rejects(rollover, { total: 0, dueDate, currency: "USD" }, "invoice.dueDate must be an ISO 8601 date");
  }
});

test("12. a malformed line throws", async () => {
  await rejects({ ...bb, line: "0".repeat(46) }, bbInvoice, "boleto.line must be 47 digits");
});

test("13. mismatches throw", async () => {
  const line = bbLine;
  await rejects({ ...bb, line }, { ...bbInvoice, total: 2434.8 }, "R$ 2.434,79");
  await rejects({ ...bb, line }, { ...bbInvoice, dueDate: "2016-01-21" }, "20/01/2016");
  await rejects({ ...bb, line }, { ...bbInvoice, dueDate: undefined }, "no due date");
  await rejects(
    { ...bb, line: "00180.00009 02656.973019 93624.706185 1 66790000243479" },
    bbInvoice,
    "fails the check digit of field 1",
  );
  await rejects({ ...bb, bankCode: "237", line }, bbInvoice, "bank 001");
  // A free field one digit off, with the line's own check digits left valid.
  await rejects({ ...bb, freeField: "0000002656973019362470619", line }, bbInvoice, "free field differs");
});

test("14. a line that cannot be built is withheld, not compared", () => {
  const r = run({ ...bb, freeField: "123", line: bbLine }, bbInvoice);
  assertEquals(r.line, null);
  assertEquals(String(r.withheld).includes("campo livre"), true);
});

// 15. Publishing the node, then resolving it with a deriver from a bundle.
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ boletoLinhaDigitavel: mod.run }));
const previewData = {
  boleto: {
    bankCode: "341",
    freeField: "1090000012345678901234561",
    line: "34191.09008 00012.345674 89012.345610 4 16610875043219",
  },
  invoice: { total: 8750432.19, dueDate: "2026-12-15", currency: "BRL" },
};

function ids(nodes: DocumentNode[]): Record<string, string> {
  const found: Record<string, string> = {};
  const walk = (list: DocumentNode[]) => {
    for (const node of list) {
      const own = node as { id?: string; text?: string; children?: DocumentNode[] };
      if (own.id) found[own.id] = own.text ?? "";
      walk(own.children ?? []);
    }
  };
  walk(nodes);
  return found;
}

for (const arrangement of ["stacked", "table"] as const) {
  const doc = () =>
    template(
      <Document title="Boleto">
        <BoletoLinhaDigitavel arrangement={arrangement} />
      </Document>,
    );
  const resolve = async (data: typeof previewData) => {
    const published = await buildDocument(doc(), createPublishData() as typeof data, {
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
    return ids((await resolveDocument(published, state, { derivers: registry })).nodes);
  };

  test(`${arrangement}: published and resolved, the line is printed`, async () => {
    const found = await resolve(previewData);
    assertEquals(found["boleto-line"], "34191.09008 00012.345674 89012.345610 4 16610875043219");
    assertEquals("boleto-due-date" in found, true);
    assertEquals("boleto-amount" in found, true);
    assertEquals("boleto-withheld" in found, false);
  });

  test(`${arrangement}: published and resolved, nothing to pay is withheld`, async () => {
    const found = await resolve({ ...previewData, invoice: { ...previewData.invoice, total: 0 } });
    assertEquals(Object.keys(found).filter((id) => id.startsWith("boleto-")), ["boleto-withheld"]);
  });
}
