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
  typeof import("../registry/components/derivers/japan-qualified-invoice.deriver.ts")
>("components/derivers/japan-qualified-invoice.deriver.ts");
const { JapanQualifiedInvoice } = await loadRegistryModule<
  typeof import("../registry/components/japan-qualified-invoice/japan-qualified-invoice.node.tsx")
>("components/japan-qualified-invoice/japan-qualified-invoice.node.tsx");

type Obj = Record<string, unknown>;
type Line = { description: string; amountYen: unknown; rate: unknown };
const T = "T7000012050002";
const line = (amountYen: unknown, rate: unknown = "10", description = "品"): Line => ({ description, amountYen, rate });
const base = (over: Obj = {}): Obj => ({
  registrationNumber: T,
  rounding: "floor",
  basis: "exclusive",
  transactionDate: "2026-11-30",
  buyerName: "株式会社テスト 御中",
  lines: [line(1000)],
  ...over,
});
const unregistered = (over: Obj = {}) => base({ registrationNumber: undefined, rounding: undefined, basis: undefined, ...over });
// deno-lint-ignore no-explicit-any
const run = (inv: unknown): any => mod.run([inv], {} as RuntimeState);
const throws = (inv: unknown, ...words: string[]) => {
  try {
    run(inv);
  } catch (e) {
    const m = (e as Error).message;
    if (!m.startsWith("japanQualifiedInvoice: ")) throw new Error("bad prefix: " + m);
    for (const w of words) if (!m.includes(w)) throw new Error(`"${m}" lacks "${w}"`);
    return;
  }
  throw new Error("did not throw");
};
const tax = (rounding: string, basis: string, rate: string, ...amounts: number[]) =>
  run(base({ rounding, basis, lines: amounts.map((a) => line(a, rate)) }))[rate === "10" ? "standard" : "reduced"].tax;

test("tax is worked per band, not per line", () => {
  assertEquals(tax("floor", "exclusive", "10", 105, 105, 105), "31円");
  assertEquals(tax("half-up", "exclusive", "10", 105, 105, 105), "32円");
  assertEquals(tax("ceil", "exclusive", "10", 105, 105, 105), "32円");
});

test("rounding modes", () => {
  assertEquals(["floor", "ceil", "half-up"].map((r) => tax(r, "exclusive", "10", 1005)), ["100円", "101円", "101円"]);
  assertEquals(["floor", "ceil", "half-up"].map((r) => tax(r, "exclusive", "8", 1234)), ["98円", "99円", "99円"]);
  assertEquals(["floor", "ceil", "half-up"].map((r) => tax(r, "inclusive", "10", 1000)), ["90円", "91円", "91円"]);
  assertEquals(tax("floor", "inclusive", "8", 1080), "80円");
  assertEquals(["floor", "ceil", "half-up"].map((r) => tax(r, "inclusive", "8", 1000)), ["74円", "75円", "74円"]);
});

test("inclusive basis labels and total", () => {
  const r = run(base({ basis: "inclusive" }));
  assertEquals(r.standard.label, "10%対象（税込）");
  assertEquals(r.standard.taxLabel, "内消費税（10%）");
  assertEquals(r.amountHeader, "金額（税込）");
  assertEquals(r.total, "1,000円");
  const x = run(base());
  assertEquals([x.standard.label, x.standard.taxLabel, x.amountHeader, x.total], ["10%対象（税抜）", "消費税（10%）", "金額（税抜）", "1,100円"]);
});

test("mixed bands: 10% first, 8% marked, legend", () => {
  const r = run(base({ lines: [line(1, "8", "a"), line(2, "10", "b"), line(3, "8", "c"), line(4, "10", "d")] }));
  assertEquals(r.items.slice(0, 4).map((i: { description: string }) => i.description), ["b", "d", "a ※", "c ※"]);
  assertEquals(r.legend, "※は軽減税率対象");
});

test("only 10% lines have no reduced band or legend", () => {
  const r = run(base());
  assertEquals([r.reduced.total, r.reduced.tax, r.legend], [null, null, null]);
});

test("zero, negative lines and negative band", () => {
  const z = run(base({ lines: [line(0)] }));
  assertEquals([z.standard.total, z.standard.tax, z.total], ["0円", "0円", "0円"]);
  const n = run(base({ lines: [line(10000), line(-2000)] }));
  assertEquals([n.standard.total, n.standard.tax, n.items[1].amount], ["8,000円", "800円", "-2,000円"]);
  throws(base({ lines: [line(-500, "8")] }), "適格返還請求書");
});

test("slots", () => {
  const one = run(base());
  assertEquals(one.items[0].description, "品");
  for (const i of one.items.slice(1)) assertEquals(i, { description: "", amount: "" });
  assertEquals(one.items.length, 10);
  const ten = run(base({ lines: Array.from({ length: 10 }, () => line(1)) }));
  assertEquals(ten.items.every((i: { description: string }) => i.description === "品"), true);
});

test("invalid lines throw", () => {
  throws(base({ lines: Array.from({ length: 11 }, () => line(1)) }), "10");
  throws(base({ lines: [] }), "lines");
  throws(base({ lines: undefined }), "lines");
  throws(base({ lines: [line(1, "0")] }), "rate");
  throws(base({ lines: [line(1, 10)] }), "rate");
  throws(base({ lines: [line(100.5)] }), "amountYen");
  throws(base({ lines: [line(1e13)] }), "amountYen");
  throws(base({ lines: [line(1, "10", "  ")] }), "description");
  throws(undefined, "jpInvoice");
});

test("registration number", () => {
  assertEquals(run(base()).registration, T);
  assertEquals(run(base({ registrationNumber: " t7-0000-1205-0002 " })).registration, T);
  assertEquals(run(base({ registrationNumber: "T9234567890123" })).registration, "T9234567890123");
  throws(base({ registrationNumber: "T8000012050002" }), "registrationNumber", "check digit");
  throws(base({ registrationNumber: "T700001205000" }), "registrationNumber");
  throws(base({ registrationNumber: "7000012050002" }), "registrationNumber");
});

test("kinds", () => {
  assertEquals(run(base()).heading, "適格請求書");
  throws(base({ buyerName: undefined }), "buyerName");
  throws(base({ buyerName: "  " }), "buyerName");
  for (const k of ["rounding", "basis"]) {
    throws(base({ [k]: undefined }), k);
    throws(base({ [k]: "nope" }), k);
  }
  const s = run(base({ simplified: true, buyerName: undefined, lines: [line(1000), line(500, "8")] }));
  assertEquals([s.heading, s.registration, s.legend, s.standard.tax, s.reduced.tax], ["適格簡易請求書", T, "※は軽減税率対象", "100円", "40円"]);
  throws(base({ simplified: true, registrationNumber: undefined }), "simplified");
});

test("unregistered", () => {
  for (const v of [undefined, "", "  "]) {
    const r = run(unregistered({ registrationNumber: v, lines: [line(1000), line(500, "8")] }));
    assertEquals(
      [r.heading, r.registration, r.amountHeader, r.standard.label, r.reduced.label, r.standard.tax, r.reduced.tax, r.totalLabel, r.total],
      ["請求書", null, "金額", "10%対象", "8%対象", null, null, "合計", "1,500円"],
    );
    assertEquals(JSON.stringify(r).includes("適格"), false);
  }
});

test("transitional note by date", () => {
  const note = (d: string) => run(unregistered({ transactionDate: d })).note as string | null;
  for (const d of ["2023-10-01", "2026-09-30"]) {
    const n = note(d)!;
    for (const w of ["80%", "10億円", "2026年9月30日までの取引"]) assertEquals(n.includes(w), true);
    assertEquals(n.includes("適格"), false);
  }
  for (const d of ["2026-10-01", "2028-09-30"]) {
    assertEquals(note(d)!.includes("70%") && note(d)!.includes("1億円") && !note(d)!.includes("10億円"), true);
  }
  assertEquals(note("2028-10-01")!.includes("50%"), true);
  for (const d of ["2030-10-01", "2031-09-30"]) assertEquals(note(d)!.includes("30%"), true);
  assertEquals(note("2031-10-01"), null);
  assertEquals(run(base()).note, null);
});

test("dates", () => {
  throws(base({ transactionDate: "2023-09-30" }), "transactionDate");
  for (const d of ["2026-02-30", "30/11/2026", undefined]) throws(base({ transactionDate: d }), "transactionDate");
  assertEquals(run(base({ transactionDate: "2026-11-30T23:59:00+09:00" })).date, "2026年11月30日");
});

test("the preview data adds up", () => {
  const r = run(base({
    registrationNumber: "T9234567890123",
    lines: [
      line(46980000, "8"), line(128450000), line(27315750, "8"), line(3862500),
      line(18642300, "8"), line(1275000), line(4215615, "8"),
    ],
  }));
  assertEquals(
    [r.standard.total, r.standard.tax, r.reduced.total, r.reduced.tax, r.total],
    ["133,587,500円", "13,358,750円", "97,153,665円", "7,772,293円", "251,872,208円"],
  );
});

// The node, published and resolved by an engine that only has the bundle.
const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ japanQualifiedInvoice: mod.run }));
const documentOf = (arrangement?: "table" | "stacked") =>
  template(
    <Document title="Invoice">
      <JapanQualifiedInvoice arrangement={arrangement} />
    </Document>,
  );

function textOf(node: DocumentNode): string[] {
  const own = node as { text?: string; children?: DocumentNode[] };
  return own.text ? [own.text] : (own.children ?? []).flatMap(textOf);
}

/** Every node id in the document, with the text under it. */
function collect(nodes: DocumentNode[], out: Map<string, string[]> = new Map()): Map<string, string[]> {
  for (const node of nodes) {
    const own = node as { id?: string; children?: DocumentNode[] };
    if (own.id) out.set(own.id, textOf(node));
    collect(own.children ?? [], out);
  }
  return out;
}

async function published(inv: Obj, arrangement?: "table" | "stacked") {
  const data = { jpInvoice: inv };
  const doc = await buildDocument(documentOf(arrangement), createPublishData() as typeof data, {
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
  return collect((await resolveDocument(doc, state, { derivers: registry })).nodes);
}

for (const arrangement of ["table", "stacked"] as const) {
  test(`${arrangement}: a registered invoice publishes and resolves`, async () => {
    const ids = await published(base({ lines: [line(1000), line(500, "8", "米")] }), arrangement);
    assertEquals(ids.has("jp-invoice-line-1") && ids.has("jp-invoice-line-2"), true);
    assertEquals(ids.has("jp-invoice-line-3"), false);
    for (const id of ["buyer", "registration", "standard-total", "standard-tax", "reduced-total", "reduced-tax", "legend", "total"]) {
      assertEquals(ids.has("jp-invoice-" + id), true);
    }
    assertEquals(ids.has("jp-invoice-note"), false);
    assertEquals(ids.get("jp-invoice-legend")!.join(""), "※は軽減税率対象");
    assertEquals(ids.get("jp-invoice-total")!.join("").includes("1,640円"), true);
  });

  test(`${arrangement}: an unregistered invoice publishes and resolves`, async () => {
    const ids = await published(unregistered({ buyerName: undefined }), arrangement);
    for (const id of ["registration", "buyer", "standard-tax", "reduced-total", "legend"]) {
      assertEquals(ids.has("jp-invoice-" + id), false);
    }
    assertEquals(ids.has("jp-invoice-note") && ids.has("jp-invoice-standard-total"), true);
    assertEquals(ids.get("jp-invoice-total")!.join("").includes("1,000円"), true);
  });
}

test("both arrangements print the same totals", async () => {
  const inv = base({ lines: [line(1000), line(500, "8")] });
  const total = async (a: "table" | "stacked") => (await published(inv, a)).get("jp-invoice-total")!.join("").replace(/\s|　/g, "");
  assertEquals(await total("table"), await total("stacked"));
});
