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
import { assertEquals } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/finnish-virtual-barcode.deriver.ts")
>("components/derivers/finnish-virtual-barcode.deriver.ts");
const { FinnishVirtualBarcode } = await loadRegistryModule<
  typeof import("../registry/components/finnish-virtual-barcode/finnish-virtual-barcode.node.tsx")
>("components/finnish-virtual-barcode/finnish-virtual-barcode.node.tsx");

type Result = { outcome: string; reason: string; iban: string; reference: string; code: string };
// The engine calls a deriver with an array and a state; the run is synchronous.
const run = (payee: unknown, invoice: unknown, payment: unknown, virtualBarcode: unknown) =>
  mod.run([payee, invoice, payment, virtualBarcode], {} as RuntimeState) as Result;

const I = { iban: "FI79 4405 2020 0360 82" };
const EUR = { currency: "EUR" };
const V4_1 = "479440520200360820048831500000000868516259619897100612";
const inv = (total: unknown, dueDate?: string, more: object = {}) => ({ ...EUR, total, dueDate, ...more });
const code = (r: Result) => [r.outcome, r.code];

test("guide vectors", () => {
  const cases: [unknown, unknown, string, string, string][] = [
    [I, 4883.15, "86851 62596 19897", "2010-06-12", V4_1],
    [{ iban: "FI58 1017 1000 0001 22" }, 482.99, "55958 22432 94671", "2012-01-31", "458101710000001220004829900000000559582243294671120131"],
    [I, 4883.15, "RF09 8685 1625 9619 897", "2010-06-12", "579440520200360820048831509000000868516259619897100612"],
    [{ iban: "FI58 1017 1000 0001 22" }, 482.99, "RF06 5595 8224 3294 671", "2012-01-31", "558101710000001220004829906000000559582243294671120131"],
  ];
  for (const [p, total, ref, due, want] of cases) {
    assertEquals(code(run(p, inv(total, due), { reference: ref }, undefined)), ["code", want]);
  }
});

test("preview data, longest RF body", () => {
  const r = run(I, inv(987654.32, "2026-12-31"), { reference: "RF26 2026 1100 4817 3500 1234 5" }, "579440520200360829876543226202611004817350012345261231");
  assertEquals(r, {
    outcome: "code",
    reason: "",
    iban: "FI79 4405 2020 0360 82",
    reference: "RF26 2026 1100 4817 3500 1234 5",
    code: "579440520200360829876543226202611004817350012345261231",
  });
});

test("20-digit viitenumero", () => {
  const r = run(I, inv(987654.32, "2026-12-31"), { reference: "20261100481735001230" }, undefined);
  assertEquals(code(r), ["code", "479440520200360829876543200020261100481735001230261231"]);
  assertEquals(r.reference, "20261 10048 17350 01230");
});

test("short RF pads after the check digits", () => {
  assertEquals(
    run(I, inv(10, "2026-01-05"), { reference: "RF7812345" }, undefined).code,
    "579440520200360820000100078000000000000000012345260105",
  );
});

test("amount field", () => {
  const ref = { reference: "1232" };
  assertEquals(run(I, inv(999999.99, "2026-12-31"), ref, undefined).code, "479440520200360829999999900000000000000000001232261231");
  assertEquals(run(I, inv(1000000, "2026-12-31"), ref, undefined).code, "479440520200360820000000000000000000000000001232261231");
  const r = run(I, inv(0.29), { reference: "0001232" }, undefined);
  assertEquals(r.code, "479440520200360820000002900000000000000000001232000000");
  assertEquals(r.reference, "1232");
  const open = run(I, { ...EUR, openAmount: true }, ref, undefined);
  assertEquals([open.outcome, open.code.slice(17, 25)], ["code", "00000000"]);
});

test("due date field", () => {
  const r = run(I, inv(999999.99, "2026-12-31", { onReceipt: true }), { reference: "1232" }, undefined);
  assertEquals(r.code.slice(-6), "000000");
});

test("lower-case input is accepted and printed upper case", () => {
  const r = run({ iban: "fi79 4405 2020 0360 82" }, inv(4883.15, "2010-06-12"), { reference: "rf09 8685 1625 9619 897" }, undefined);
  assertEquals(code(r), ["code", "579440520200360820048831509000000868516259619897100612"]);
  assertEquals([r.iban, r.reference], ["FI79 4405 2020 0360 82", "RF09 8685 1625 9619 897"]);
});

test("empty and timestamped due dates", () => {
  const ref = { reference: "1232" };
  assertEquals(run(I, inv(10, ""), ref, undefined).code.slice(-6), "000000");
  assertEquals(run(I, inv(10, "2026-12-31T23:59:00Z"), ref, undefined).code.slice(-6), "261231");
});

test("nothing owed prints nothing", () => {
  const none = { outcome: "none", reason: "", iban: "", reference: "", code: "" };
  assertEquals(run(I, inv(0), { reference: "1232" }, undefined), none);
  assertEquals(run({ iban: "bad" }, inv(-120.5), { reference: "1232" }, undefined), none);
});

test("withheld cases still print what they can", () => {
  const ref = { reference: "1232" };
  const w = (r: Result, reason: string) => assertEquals([r.outcome, r.reason, r.code], ["withheld", reason, ""]);
  let r = run(I, inv(12.345, "2026-12-31"), ref, undefined);
  w(r, "amount");
  assertEquals([r.iban, r.reference], ["FI79 4405 2020 0360 82", "1232"]);
  w(run(I, inv(undefined), ref, undefined), "amount");
  w(run(I, inv("4883.15"), ref, undefined), "amount");

  r = run({ iban: " FI79 4405 2020 0360 83 " }, inv(10), ref, undefined);
  w(r, "iban");
  assertEquals([r.iban, r.reference], ["FI79 4405 2020 0360 83", "1232"]);

  r = run({ iban: "SE4550000000058398257466" }, inv(10), ref, undefined);
  w(r, "not-finnish");
  assertEquals(r.iban, "SE45 5000 0000 0583 9825 7466");

  w(run(I, { ...inv(10), currency: "SEK" }, ref, undefined), "currency");
  r = run(I, inv(10), { reference: "1233" }, undefined);
  w(r, "reference");
  assertEquals(r.reference, "1233");
  w(run(I, inv(10), { reference: "RF1912345" }, undefined), "reference");
  w(run(I, inv(10), { reference: "0000" }, undefined), "reference");
  r = run(I, inv(10), { reference: "RF96INV2026X" }, undefined);
  w(r, "rf-alphanumeric");
  assertEquals(r.reference, "RF96 INV2 026X");
  w(run(I, inv(10, "2026-02-30"), ref, undefined), "due-date");
  w(run(I, inv(10, "1999-12-31"), ref, undefined), "due-date");

  r = run({}, inv(10), {}, undefined);
  w(r, "iban");
  assertEquals([r.iban, r.reference], ["", ""]);
  w(run(undefined, undefined, undefined, undefined), "iban");
});

test("reference is grouped even when the IBAN fails first", () => {
  assertEquals(run({ iban: "x" }, inv(10), { reference: "0086851625961 9897" }, undefined).reference, "86851 62596 19897");
});

test("a reference is grouped in fives from the right", () => {
  assertEquals(run(I, inv(10, "2026-12-31"), { reference: "1234561" }, undefined).reference, "12 34561");
});

test("ERP cross-check", () => {
  const p = { reference: "86851 62596 19897" };
  const spaced = "4794 4052 0200 3608 2004 8831 5000 0000 0868 5162 5961 9897 1006 12";
  assertEquals(code(run(I, inv(4883.15, "2010-06-12"), p, spaced)), ["code", V4_1]);
  for (const erp of ["579440520200360820048831509000000868516259619897100612", "2" + V4_1.slice(1)]) {
    const r = run(I, inv(4883.15, "2010-06-12"), p, erp);
    assertEquals([r.outcome, r.reason, r.code], ["withheld", "mismatch", ""]);
  }
});

// ---- rendering and publishing ----

const data = {
  payee: I,
  invoice: inv(4883.15, "2010-06-12"),
  payment: { reference: "86851 62596 19897" },
};
const ARRANGEMENTS = ["stacked", "table"] as const;
const documentOf = (arrangement?: (typeof ARRANGEMENTS)[number]) =>
  template(<Document title="Invoice"><FinnishVirtualBarcode arrangement={arrangement} /></Document>);

type Found = { id?: string; text?: string };
function walk(nodes: DocumentNode[]): Found[] {
  return nodes.flatMap((n) => {
    const own = n as { id?: string; text?: string; children?: DocumentNode[]; rows?: DocumentNode[]; cells?: DocumentNode[] };
    return [{ id: own.id, text: own.text }, ...walk([...(own.children ?? []), ...(own.rows ?? []), ...(own.cells ?? [])])];
  });
}
const ids = (nodes: DocumentNode[]) => walk(nodes).map((f) => f.id).filter(Boolean);

for (const arrangement of ARRANGEMENTS) {
  test(`${arrangement}: the code is one paragraph of exactly 54 digits`, async () => {
    const { nodes } = await buildDocument(documentOf(arrangement), data, { dynamicMode: "placeholder" });
    const found = walk(nodes);
    for (const id of ["virtual-barcode-iban", "virtual-barcode-reference", "virtual-barcode-label", "virtual-barcode-code"]) {
      assertEquals(ids(nodes).includes(id), true);
    }
    assertEquals(found.find((f) => f.id === "virtual-barcode-code")?.text, V4_1);
    assertEquals(ids(nodes).includes("virtual-barcode-details"), arrangement === "table");
  });

  test(`${arrangement}: a withheld code keeps the IBAN and reference`, async () => {
    const bad = { ...data, payment: { reference: "1233" } };
    const { nodes } = await buildDocument(documentOf(arrangement), bad, { dynamicMode: "placeholder" });
    const got = ids(nodes);
    assertEquals(got.includes("virtual-barcode-label") || got.includes("virtual-barcode-code"), false);
    assertEquals(got.includes("virtual-barcode-iban") && got.includes("virtual-barcode-reference"), true);
  });

  test(`${arrangement}: nothing owed prints no node at all`, async () => {
    const { nodes } = await buildDocument(documentOf(arrangement), { ...data, invoice: inv(0) }, { dynamicMode: "placeholder" });
    assertEquals(ids(nodes).filter((id) => String(id).startsWith("virtual-barcode")), []);
  });
}

test("no arrangement is the stacked one", async () => {
  const build = (a?: "stacked") => buildDocument(documentOf(a), data, { dynamicMode: "placeholder" });
  assertEquals((await build()).nodes, (await build("stacked")).nodes);
});

for (const arrangement of ARRANGEMENTS) {
  test(`${arrangement}: published, bundled and resolved, the code is printed, and not when nothing is owed`, async () => {
    const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof data, {
      branchMode: "publish",
      deriverMode: "preserve",
      dynamicMode: "resolve",
      aiClient: { generateParagraph: () => "" },
    });
    const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ finnishVirtualBarcode: mod.run }));
    const resolve = (d: typeof data) => {
      const state: RuntimeState = {
        ctx: { ...d },
        derived: {},
        dataProvider: new InMemoryDataProvider(d),
        aiClient: { generateParagraph: () => "" },
      };
      return resolveDocument(published, state, { derivers });
    };
    const printed = await resolve(data);
    assertEquals(walk(printed.nodes).find((f) => f.id === "virtual-barcode-code")?.text, V4_1);
    const none = await resolve({ ...data, invoice: inv(0) });
    assertEquals(ids(none.nodes).filter((id) => String(id).startsWith("virtual-barcode")), []);
  });
}
