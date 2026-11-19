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

const nodeSource = "components/exemption-certificate-check/exemption-certificate-check.node.tsx";
const { ExemptionCertificateCheck } = await loadRegistryModule<
  typeof import("../registry/components/exemption-certificate-check/exemption-certificate-check.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/exemption-certificate-check.deriver.ts")
>("components/derivers/exemption-certificate-check.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ exemptionCertificateCheck: mod.run }));
// The section's arm is in the node; the sixteen row arms are in an arrangement.
const countBranches = async (source: string) =>
  transformDocumentSource(await readFile(join(registryRoot(), source), "utf8"), { fileName: source }).branches;
const branchesOf = async (arrangement: "table" | "stacked") =>
  (await countBranches(nodeSource)) +
  (await countBranches(`components/arrangements/exemption-certificate-check/${arrangement}.arrangement.tsx`));

// deno-lint-ignore no-explicit-any
type Obj = Record<string, any>;
const run = (e: Obj) => mod.run([e], {} as RuntimeState) as Obj;

const SST = ["NC", "GA", "OH", "NJ"];
const base = (over: Obj = {}): Obj => ({
  invoiceDate: "2026-11-30",
  shipToState: "NC",
  sstMembers: { asOf: "2026-10-01", states: SST },
  ...over,
});
const line = (category?: string, amountMinor = 10000): Obj => ({ description: "Widget", amountMinor, category });
const one = (over: Obj, l: Obj = line("resale")) => run(base({ lines: [l], ...over })).lines[0] as Obj;
const sstG = (o: Obj = {}) => ({ form: "SST", states: ["NC"], reason: "G", ...o });

test("no lines prints nothing", async () => {
  for (const lines of [undefined, []]) {
    const o = run(base({ lines }));
    assertEquals([o.hasLines, o.summary, o.basis, o.exemptMinor, o.taxableMinor, o.rows.length], [false, "", "", 0, 0, 16]);
    assertEquals(o.rows.every((r: Obj) => r.description === "" && r.status === ""), true);
    const doc = await buildDocument(templateOf(), { exemption: base({ lines }) }, opts());
    assertEquals(ids(doc.nodes).includes("exemption-check"), false);
  }
});

test("no certificates", () => {
  const o = run(base({ lines: [line("resale", 12345)] }));
  assertEquals(o.lines[0].failedRule, "No exemption certificate on file");
  assertEquals(o.summary, "Exempt: 0 of 1 lines, $0.00. Taxable: 1 of 1 lines, $123.45.");
});

test("the ship-to state, not the billing state", () => {
  const l = one({ certificates: [{ form: "state", state: "VA", reason: "G" }] });
  assertEquals([l.taxable, l.failedRule], [true, "Certificate covers VA, not the ship-to state NC"]);
});

test("SST form outside the member list", () => {
  for (const st of ["CA", "TX", "FL", "NY"]) {
    const l = one({ shipToState: st, certificates: [sstG({ states: [st] })] });
    assertEquals(l.failedRule, st + " is not a Streamlined Sales Tax full member as of October 1, 2026, so Form F0003 is not accepted there");
  }
});

test("MTC acceptance", () => {
  const mtc = (o: Obj = {}) => ({ form: "MTC", states: ["MD"], reason: "G", ...o });
  const md = { shipToState: "MD" };
  assertEquals(one({ ...md, certificates: [mtc()] }).failedRule, "No list of states accepting the MTC uniform certificate was supplied");
  const list = (resaleOnly?: boolean) => ({ mtcAcceptingStates: { asOf: "2026-07-01", states: [{ code: " md ", resaleOnly }] } });
  assertEquals(one({ ...md, ...list(true), certificates: [mtc()] }).taxable, false);
  assertEquals(
    one({ ...md, ...list(true), certificates: [mtc({ reason: "E" })] }, line("own-use")).failedRule,
    "MD accepts the MTC uniform certificate for resale only, and this one claims Charitable organization (E)",
  );
  assertEquals(one({ ...md, ...list(undefined as never), certificates: [mtc({ reason: "E" })] }).taxable, true);
  assertEquals(one({ ...md, ...list(false), certificates: [mtc({ reason: "E" })] }).taxable, false);
  assertEquals(one({ ...md, ...list(false), certificates: [mtc({ states: ["VA"] })] }).failedRule, "Certificate covers VA, not the ship-to state MD");
  assertEquals(one({ shipToState: "GA", ...list(false), certificates: [mtc({ states: ["GA"] })] }).failedRule, "GA does not accept the MTC uniform certificate");
});

test("reasons", () => {
  const c = (reason: string) => ({ certificates: [sstG({ reason })] });
  assertEquals(
    one(c("G"), line("own-use")).failedRule,
    "Resale (G) covers only lines categorised resale; this line is own-use",
  );
  assertEquals(one(c("G"), line()).failedRule, "Resale (G) covers only lines categorised resale; this line has no category");
  assertEquals(one(c("H"), line("resale")).failedRule, "Agricultural production (H) covers only lines categorised agricultural; this line is resale");
  for (const r of ["E", "A", "M"]) {
    for (const cat of ["own-use", "resale", undefined]) assertEquals(one(c(r), line(cat)).taxable, false);
  }
  assertEquals(one(c("J")).exemptReason, "Direct pay permit (J), Streamlined certificate (F0003); the buyer pays the tax to the state directly");
  assertEquals(one(c("L")).failedRule, "Other (L) needs a manual review and is never accepted automatically");
  assertEquals(
    one({ certificates: [sstG({ number: "N-1" })] }).exemptReason,
    "Resale (G), Streamlined certificate (F0003) N-1",
  );
  assertEquals(
    one({ shipToState: "VA", certificates: [{ form: "state", state: "VA", reason: "I" }] }, line("industrial")).exemptReason,
    "Industrial production or manufacturing (I), VA state certificate",
  );
});

test("a certificate that names no state", () => {
  assertEquals(one({ certificates: [{ form: "SST", reason: "E" }] }).failedRule, "Certificate names no state");
  assertEquals(one({ certificates: [{ form: "SST", state: "nc", reason: "E" }] }).taxable, false);
});

test("expiry is inclusive", () => {
  const e = (expires?: string) => one({ certificates: [sstG({ expires })] });
  assertEquals(e("2026-11-30").taxable, false);
  assertEquals(e(undefined).taxable, false);
  assertEquals(e("2026-11-29").failedRule, "Certificate expired on November 29, 2026");
});

test("a pending certificate in an SST state", () => {
  const p = (checkedOn?: string) => one({ checkedOn, certificates: [sstG({ pending: true, expires: "2020-01-01" })] });
  assertEquals(p().exemptReason, "Exemption certificate pending, required by February 28, 2027");
  assertEquals(p("2027-02-28").taxable, false);
  assertEquals(
    p("2027-03-01").failedRule,
    "Exemption certificate pending since November 30, 2026 was not received by February 28, 2027",
  );
});

test("a pending certificate outside an SST state", () => {
  const l = one({ shipToState: "CA", certificates: [{ form: "state", state: "CA", reason: "G", pending: true }] });
  assertEquals(l.failedRule, "A pending certificate is honoured only in Streamlined member states, and CA is not one");
});

test("receivedOn", () => {
  const r = (receivedOn: string, shipToState = "NC") =>
    one({ shipToState, certificates: [shipToState === "NC" ? sstG({ receivedOn }) : { form: "state", state: shipToState, reason: "G", receivedOn }] });
  assertEquals(r("2026-11-30").taxable, false);
  assertEquals(r("2027-02-28").taxable, false);
  assertEquals(r("2027-03-01").failedRule, "Certificate received on March 1, 2027, after the 90-day deadline of February 28, 2027");
  assertEquals(r("2026-12-01", "CA").failedRule, "Certificate received on December 1, 2026, after the sale; CA gives no 90-day window here");
});

test("precedence", () => {
  const full = one({ certificates: [sstG({ pending: true }), sstG({ number: "FULL" })] });
  assertEquals(full.exemptReason, "Resale (G), Streamlined certificate (F0003) FULL");
  const furthest = one({
    certificates: [{ form: "state", state: "NC", reason: "H" }, { form: "state", state: "VA", reason: "G" }],
  });
  assertEquals(furthest.failedRule, "Certificate covers VA, not the ship-to state NC");
  const tie = one({ certificates: [sstG({ reason: "H" }), sstG({ reason: "I" })] });
  assertEquals(tie.failedRule, "Agricultural production (H) covers only lines categorised agricultural; this line is resale");
});

test("amounts and currency", () => {
  const o = run(base({
    certificates: [sstG()],
    lines: [line("resale", 0), line("own-use", -12500), line("resale", 5000)],
  }));
  assertEquals(o.rows.slice(0, 3).map((r: Obj) => r.amount), ["$0.00", "-$125.00", "$50.00"]);
  assertEquals([o.exemptMinor, o.taxableMinor], [5000, -12500]);
  assertEquals(o.summary, "Exempt: 2 of 3 lines, $50.00. Taxable: 1 of 3 lines, -$125.00.");
  assertEquals(run(base({ currency: "JPY", lines: [line("resale", 1500)] })).rows[0].amount, "¥1,500");
  assertEquals(o.lines[1].category, "own-use");
  assertEquals("category" in run(base({ lines: [{ description: "x", amountMinor: 1 }] })).lines[0], false);
});

test("basis", () => {
  assertEquals(
    run(base({ lines: [line()] })).basis,
    "Checked on November 30, 2026 for delivery to NC, against Streamlined Sales Tax full members as of October 1, 2026.",
  );
  const withMtc = run(base({ checkedOn: "2026-12-14", mtcAcceptingStates: { asOf: "2026-07-01", states: [] }, lines: [line()] }));
  assertEquals(
    withMtc.basis,
    "Checked on December 14, 2026 for delivery to NC, against Streamlined Sales Tax full members as of October 1, 2026 and states accepting the MTC uniform certificate as of July 1, 2026.",
  );
});

test("lowercase and padded codes match", () => {
  const l = one({ shipToState: " nc ", sstMembers: { asOf: "2026-10-01", states: [" nc "] }, certificates: [sstG({ states: [" Nc"] })] });
  assertEquals(l.taxable, false);
});

test("16 lines fill every slot and 17 throw", () => {
  const many = (n: number) => Array.from({ length: n }, (_, i) => ({ description: "L" + i, amountMinor: 100 }));
  assertEquals(run(base({ lines: many(16) })).rows.every((r: Obj) => r.description !== ""), true);
  assertThrows(() => run(base({ lines: many(17) })), "17 lines; this component prints at most 16");
});

function assertThrows(fn: () => unknown, includes: string) {
  let message = "";
  try {
    fn();
  } catch (e) {
    message = (e as Error).message;
  }
  assertStringIncludes(message, includes);
}

test("validation names the field", () => {
  const bad = (over: Obj, path: string) => assertThrows(() => run(base({ lines: [line()], ...over })), path);
  bad({ invoiceDate: undefined }, "exemption.invoiceDate");
  bad({ invoiceDate: "30/11/2026" }, "exemption.invoiceDate");
  bad({ checkedOn: "soon" }, "exemption.checkedOn");
  bad({ checkedOn: "2026-11-29" }, "exemption.checkedOn");
  bad({ shipToState: "N" }, "exemption.shipToState");
  bad({ shipToState: undefined }, "exemption.shipToState");
  bad({ currency: "XX1" }, "exemption.currency");
  bad({ sstMembers: undefined }, "exemption.sstMembers.asOf");
  bad({ sstMembers: { asOf: "nope", states: [] } }, "exemption.sstMembers.asOf");
  bad({ sstMembers: { asOf: "2026-10-01", states: "NC" } }, "exemption.sstMembers.states");
  bad({ mtcAcceptingStates: { asOf: "nope", states: [] } }, "exemption.mtcAcceptingStates.asOf");
  bad({ mtcAcceptingStates: { asOf: "2026-07-01", states: {} } }, "exemption.mtcAcceptingStates.states");
  bad({ certificates: [sstG({ form: "XYZ" })] }, "exemption.certificates[0].form");
  bad({ certificates: [sstG({ reason: "Z" })] }, "exemption.certificates[0].reason");
  bad({ certificates: [sstG(), sstG({ expires: "never" })] }, "exemption.certificates[1].expires");
  bad({ certificates: [sstG({ receivedOn: "later" })] }, "exemption.certificates[0].receivedOn");
  bad({ lines: [line(), { description: " ", amountMinor: 1 }] }, "exemption.lines[1].description");
  bad({ lines: [line(), line(), { description: "x", amountMinor: 1.5 }] }, "exemption.lines[2].amountMinor");
});

test("the catalog's preview data", () => {
  const entry = COMPONENTS.find((c) => c.id === "exemption-certificate-check")!;
  const o = run((entry.previewData as { exemption: Obj }).exemption);
  assertEquals(o.rows.slice(0, 4).map((r: Obj) => r.reason), [
    "Resale (G), Streamlined certificate (F0003) NC-SST-500194-7731-REV3",
    "Certificate covers VA, not the ship-to state NC",
    "Exemption certificate pending, required by February 28, 2027",
    "Resale (G) covers only lines categorised resale; this line is own-use",
  ]);
  assertEquals(
    o.summary,
    "Exempt: 2 of 4 lines, $1,970,849.75 (1 pending a certificate). Taxable: 2 of 4 lines, $1,013,125.50.",
  );
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = (arrangement?: "table" | "stacked") =>
  template<Obj>(
    <Document id="d" title="Invoice">
      <Section id="body" title="Body">
        <ExemptionCertificateCheck arrangement={arrangement} />
      </Section>
    </Document>,
  );
const opts = () => ({ branchMode: "decide", dynamicMode: "placeholder", derivers }) as never;

function collect(nodes: DocumentNode[], into: DocumentNode[] = []): DocumentNode[] {
  for (const n of nodes) {
    into.push(n);
    const children = (n as { children?: DocumentNode[] }).children;
    if (children) collect(children, into);
  }
  return into;
}
const ids = (nodes: DocumentNode[]) => collect(nodes).map((n) => n.id);
// A resolved tree numbers its cells differently, so compare what a reader sees.
const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);

function state(data: Obj): RuntimeState {
  return { ctx: { ...data }, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } };
}

const preview = () => (COMPONENTS.find((c) => c.id === "exemption-certificate-check")!.previewData as { exemption: Obj }).exemption;

for (const arrangement of ["table", "stacked"] as const) {
test(`a publish build carries 17 arms and one deriver call (${arrangement})`, async () => {
  assertEquals(await branchesOf(arrangement), 17);
  const published = await buildDocument(templateOf(arrangement), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 17);
  const section = all.find((n) => n.id === "exemption-check") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  const d = JSON.parse(JSON.stringify(section.derivers[0]));
  assertEquals([d.name, d.inputs], ["exemptionCertificateCheck", [{ type: "ref", ref: { scope: "data", path: "exemption" } }]]);

  const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ exemptionCertificateCheck: mod.run }));
  for (const exemption of [preview(), base({ lines: [] })]) {
    const direct = await buildDocument(templateOf(arrangement), { exemption }, opts());
    const resolved = await resolveDocument(published, state({ exemption }), { derivers: registry });
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
  }
});
}

test("both arrangements print the same rows and texts", async () => {
  const get = async (a: "table" | "stacked") => {
    const doc = await buildDocument(templateOf(a), { exemption: preview() }, opts());
    return collect(doc.nodes).filter((n) => n.id.startsWith("exemption-check-line-"));
  };
  const [t, s] = [await get("table"), await get("stacked")];
  assertEquals(t.map((n) => n.id), s.map((n) => n.id));
  assertEquals(t.length, 4);
  for (const [i, n] of s.entries()) assertStringIncludes(JSON.stringify(n), run(preview()).rows[i].reason);
});
