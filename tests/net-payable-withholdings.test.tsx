import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  dataRef,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/net-payable-withholdings.deriver.ts")
>("components/derivers/net-payable-withholdings.deriver.ts");
const { NetPayableWithholdings } = await loadRegistryModule<
  typeof import("../registry/components/net-payable-withholdings/net-payable-withholdings.node.tsx")
>("components/net-payable-withholdings/net-payable-withholdings.node.tsx");
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ netPayableWithholdings: mod.run }));

type Obj = Record<string, unknown>;
type Result = ReturnType<NonNullable<typeof mod.__types>>;
interface Input {
  invoice: Obj;
  supplier: Obj;
  customer: Obj;
}

const es = (invoice: Obj = {}, supplier: Obj = {}, customer: Obj = {}): Input => ({
  invoice: { country: "ES", base: "1000.00", vatPercent: "21", ...invoice },
  supplier: { professional: true, ...supplier },
  customer: { type: "business", ...customer },
});
const it = (invoice: Obj = {}, supplier: Obj = {}, customer: Obj = {}): Input => ({
  invoice: { country: "IT", base: "1000.00", vatPercent: "22", ...invoice },
  supplier: { professional: true, ...supplier },
  customer: { type: "business", withholdingAgent: true, ...customer },
});
const forfettario = (invoice: Obj = {}, supplier: Obj = {}) =>
  it({ vatPercent: undefined, ...invoice }, { forfettario: true, rivalsaInps: true, ...supplier });

async function run({ invoice, supplier, customer }: Input): Promise<Result> {
  const data = { invoice, supplier, customer };
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
  return await registry.run(
    {
      name: "netPayableWithholdings",
      output: "p",
      inputs: [dataRef("invoice"), dataRef("supplier"), dataRef("customer")],
    } as never,
    state,
  ) as Result;
}

/** The same invariant on every figure: the rows add up to the total, and the total less deductions to the net. */
function checkInvariant(r: Result) {
  const c = r.cents;
  assertEquals(c.base + c.rivalsa + c.cassa + c.vat + c.re + c.exempt + c.stamp, c.total);
  assertEquals(c.total - c.withholding - c.splitPayment, c.net);
}

async function figures(input: Input) {
  const r = await run(input);
  checkInvariant(r);
  return { total: r.cents.total, withholding: r.cents.withholding, net: r.cents.net, r };
}

// ICU puts a no-break space before the euro sign.
const eur = (s: string) => s.replace(" €", " €");

test("Spain: IRPF comes off the base, by customer and professional status", async () => {
  for (
    const [input, withholding, net] of [
      [es(), 15000, 106000],
      [es({}, { newProfessional: true }), 7000, 114000],
      [es({}, {}, { type: "consumer" }), 0, 121000],
      [es({}, {}, { type: "public-administration" }), 15000, 106000],
      [es({ exemptAmount: "50.00" }), 15000, 111000],
      [es({}, { professional: undefined }), 0, 121000],
    ] as const
  ) {
    const f = await figures(input);
    assertEquals([f.withholding, f.net], [withholding, net]);
  }
  const r = await run(es());
  assertEquals(
    [r.baseLabel, r.vatLabel, r.vat, r.total, r.withholdingLabel, r.withholding, r.net, r.title],
    [
      "Base imponible",
      "IVA 21% (art. 90 Ley 37/1992)",
      eur("210,00 €"),
      eur("1210,00 €"),
      eur("Retención IRPF 15% sobre 1000,00 € (art. 101.5 Ley 35/2006)"),
      eur("-150,00 €"),
      eur("1060,00 €"),
      "Total a pagar",
    ],
  );
});

test("Spain: recargo de equivalencia follows the IVA rate", async () => {
  for (
    const [vatPercent, re, net, label] of [
      ["21", 5200, 126200, "5,2%"],
      ["10", 1400, 111400, "1,4%"],
      ["4", 500, 104500, "0,5%"],
      ["0", 0, 100000, null],
    ] as const
  ) {
    const f = await figures(es({ vatPercent }, { professional: false }, { reRegime: true }));
    assertEquals([f.r.cents.re, f.net, f.total], [re, net, net]);
    assertEquals(f.r.reLabel === null ? null : f.r.reLabel.includes(label!), label === null ? null : true);
    assertEquals(f.r.vatLabel === null, vatPercent === "0");
  }
});

test("rounding is half up, in integer cents", async () => {
  const r = await figures(es({ base: "0.10" }));
  assertEquals([r.r.cents.vat, r.total, r.withholding, r.net], [2, 12, 2, 10]);
  assertEquals((await run(es({ base: "0.50" }))).cents.vat, 11);
  assertEquals((await run(es({ base: 1000.5 }))).cents.base, 100050);
});

test("Italy: the ritenuta base is the fee plus rivalsa, never the cassa", async () => {
  const rivalsa = await figures(it({}, { rivalsaInps: true }));
  assertEquals([rivalsa.total, rivalsa.withholding, rivalsa.net], [126880, 20800, 106080]);
  assertEquals(rivalsa.r.vatLabel, eur("IVA 22% su 1040,00 € (art. 16 DPR 633/1972)"));

  const cassa = await figures(it({}, { cassaPercent: "4", cassaName: "Cassa Forense" }));
  assertEquals([cassa.total, cassa.withholding, cassa.net], [126880, 20000, 106880]);
  assertEquals(cassa.r.cassaLabel, "Contributo integrativo Cassa Forense 4%");
  assertEquals(cassa.r.withholdingLabel, eur("Ritenuta d'acconto 20% su 1000,00 € (art. 25 DPR 600/1973)"));

  const notAgent = await figures(it({}, {}, { withholdingAgent: false }));
  assertEquals([notAgent.withholding, notAgent.net, notAgent.total], [0, 122000, 122000]);
  assertEquals(notAgent.r.withholdingLabel, null);

  const consumer = await figures(it({}, {}, { type: "consumer", withholdingAgent: undefined }));
  assertEquals(consumer.withholding, 0);
});

test("Italy forfettario: no IVA, bollo as a note or re-charged", async () => {
  const note = await figures(forfettario());
  assertEquals([note.total, note.net, note.r.vatLabel], [104000, 104000, null]);
  assertEquals(note.r.stampNote, eur("Imposta di bollo di 2,00 € assolta dall'emittente (DPR 642/1972)."));
  assertEquals(note.r.withholdingNote?.startsWith("Compenso non soggetto"), true);
  assertEquals(note.r.stampLabel, null);

  const charged = await figures(forfettario({ stampRecharged: true }));
  assertEquals([charged.total, charged.net, charged.r.stampNote], [104200, 104200, null]);
  assertEquals(charged.r.stamp, eur("2,00 €"));
});

test("bollo is strictly over 77.47 carrying no IVA", async () => {
  const due = async (input: Input) => (await run(input)).stampNote !== null;
  assertEquals(await due(forfettario({ base: "77.47" }, { rivalsaInps: false })), false);
  assertEquals(await due(forfettario({ base: "77.48" }, { rivalsaInps: false })), true);
  assertEquals(await due(it({ exemptAmount: "77.47" })), false);
  assertEquals(await due(it({ exemptAmount: "77.48" })), true);
  const spain = await run(es({ exemptAmount: "500.00" }));
  assertEquals([spain.stampLabel, spain.stampNote], [null, null]);
});

test("split payment shows the IVA, then takes it off", async () => {
  const f = await figures(it({ splitPayment: true }, { professional: false }, { type: "public-administration", withholdingAgent: undefined }));
  assertEquals([f.r.cents.vat, f.total, f.r.cents.splitPayment, f.net], [22000, 122000, 22000, 100000]);
  assertEquals(f.r.splitPayment, eur("-220,00 €"));
  assertEquals(f.r.splitPaymentLabel?.includes("scissione dei pagamenti"), true);
});

test("a zero base prints zero rows and does not throw", async () => {
  const r = await figures(es({ base: "0" }));
  assertEquals([r.r.base, r.r.vat, r.r.total, r.r.withholding, r.r.net], ["0,00 €", "0,00 €", "0,00 €", "-0,00 €", "0,00 €"].map(eur));
});

test("unclear or contradictory data throws, naming the field", async () => {
  const bad = (input: Input, field: string) => assertRejects(() => run(input), Error, "netPayableWithholdings: " + field);
  await bad(es({ country: "FR" }), "invoice.country");
  for (const base of ["-100", "10.005", undefined, 0.1 + 0.2]) await bad(es({ base }), "invoice.base");
  await bad(es({ vatPercent: "22" }), "invoice.vatPercent");
  await bad(es({ vatPercent: undefined }), "invoice.vatPercent");
  await bad(forfettario({ vatPercent: "22" }), "invoice.vatPercent");
  await bad(es({}, {}, { type: undefined }), "customer.type");
  await bad(it({}, { professional: false, newProfessional: true }), "supplier.newProfessional");
  await bad(es({}, { professional: false, newProfessional: true }), "supplier.newProfessional");
  await bad(it({}, { rivalsaInps: true, cassaPercent: "4" }), "supplier.rivalsaInps");
  for (const cassaPercent of ["0", "12"]) await bad(it({}, { cassaPercent }), "supplier.cassaPercent");
  await bad(it({}, {}, { withholdingAgent: undefined }), "customer.withholdingAgent");
  await bad(it({}, {}, { type: "consumer", withholdingAgent: true }), "customer.withholdingAgent");
  await bad(es({ splitPayment: true }), "invoice.splitPayment");
  const split = { splitPayment: true };
  await bad(it(split, { professional: false }, { type: "consumer", withholdingAgent: undefined }), "invoice.splitPayment");
  await bad(forfettario(split), "invoice.splitPayment");
  await bad(it({ ...split, vatPercent: "0" }, { professional: false }), "invoice.splitPayment");
  await bad(it(split), "invoice.splitPayment");
  await bad(it({}, {}, { reRegime: true }), "customer.reRegime");
  await bad(es({}, { rivalsaInps: true }), "supplier.rivalsaInps");
  await bad(es({ stampRecharged: true }), "invoice.stampRecharged");
  // The other country's field set to false is not a contradiction.
  await run(es({ splitPayment: false, stampRecharged: false }, { rivalsaInps: false, forfettario: false }));
  await run(it({}, { newProfessional: false }, { reRegime: false }));
});

// The node, published and resolved by an engine that only has the bundle.
const ARRANGEMENTS = ["table", "stacked"] as const;
const documentOf = (arrangement?: (typeof ARRANGEMENTS)[number]) =>
  template(
    <Document title="Net payable">
      <NetPayableWithholdings arrangement={arrangement} />
    </Document>,
  );

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { title?: string; text?: string; children?: DocumentNode[] };
    return [...(node.kind === "section" && own.title ? [own.title] : []), ...(own.text ? [own.text] : []), ...texts(own.children ?? [])];
  });
}
const conditionals = (nodes: DocumentNode[]): number =>
  nodes.reduce(
    (n, node) => n + ("when" in node && node.when ? 1 : 0) + conditionals((node as { children?: DocumentNode[] }).children ?? []),
    0,
  );

const cases: Record<string, Input> = {
  "case 1": es(),
  "case 8": it({}, { cassaPercent: "4" }),
  "case 11": forfettario(),
  "case 15": it({ splitPayment: true }, { professional: false }, { type: "public-administration", withholdingAgent: undefined }),
};

const publish = (arrangement: (typeof ARRANGEMENTS)[number]) =>
  buildDocument(documentOf(arrangement), createPublishData() as unknown as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });

test("published, the node carries 10 conditionals and one deriver call", async () => {
  const built = await publish("table");
  assertEquals(conditionals(built.nodes), 10);
  assertEquals(JSON.stringify(built).split('"netPayableWithholdings"').length - 1 >= 1, true);
  assertEquals((JSON.stringify(built).match(/"name":"netPayableWithholdings"/g) ?? []).length, 1);
});

for (const arrangement of ARRANGEMENTS) {
  for (const [name, input] of Object.entries(cases)) {
    test(`${arrangement}: publishing then resolving matches building directly (${name})`, async () => {
      const state: RuntimeState = {
        ctx: { ...input },
        derived: {},
        dataProvider: new InMemoryDataProvider({ ...input }),
        aiClient: { generateParagraph: () => "" },
      };
      const resolved = await resolveDocument(await publish(arrangement), state, { derivers: registry });
      const direct = await buildDocument(documentOf(arrangement), input as unknown as Obj, { dynamicMode: "placeholder" });
      assertEquals(texts(resolved.nodes), texts(direct.nodes));
    });
  }
}

test("the preview data prints the eight rows, in both arrangements", async () => {
  const input: Input = {
    invoice: { country: "IT", base: "48650.00", vatPercent: "22", exemptAmount: "1285.40", stampRecharged: true },
    supplier: { professional: true, cassaPercent: "4", cassaName: "Cassa Nazionale di Previdenza e Assistenza Forense" },
    customer: { type: "business", withholdingAgent: true },
  };
  const r = await figures(input);
  assertEquals(
    [r.r.base, r.r.cassa, r.r.vat, r.r.exempt, r.r.stamp, r.r.total, r.r.withholding, r.r.net],
    ["48.650,00 €", "1946,00 €", "11.131,12 €", "1285,40 €", "2,00 €", "63.014,52 €", "-9730,00 €", "53.284,52 €"].map(eur),
  );
  const lines = (a: (typeof ARRANGEMENTS)[number]) =>
    buildDocument(documentOf(a), input as unknown as Obj, { dynamicMode: "placeholder" }).then((d) => texts(d.nodes));
  const table = await lines("table");
  const stacked = await lines("stacked");
  assertEquals(table.length, 1 + 16);
  assertEquals(stacked.length, 1 + 8);
  assertEquals(stacked[1], eur("Compenso: 48.650,00 €"));
  assertEquals(table.includes(eur("53.284,52 €")), true);
});
