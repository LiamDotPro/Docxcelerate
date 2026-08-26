/**
 * The invoice verification runner: invariants -> build -> probes ->
 * objectives -> report. See scripts/lib/VERIFY-CONTRACT.md — that file is the
 * spec this one implements, and where the two disagree the contract wins.
 *
 * The runner never judges the document itself. Every threshold lives in
 * ./lib/objectives.mjs (hashed per run, so a moved goalpost is visible), every
 * measurement comes from a probe, and this file only wires them together and
 * writes the board. A first board that is nearly all FAIL is the expected
 * output, not a malfunction.
 *
 * Probes and objectives are imported dynamically and every failure is caught:
 * the harness is authored by several hands at once, and a half-authored
 * sibling must degrade to `probes[X].ok = false` on the board rather than
 * crash the run. The probe JSON files named in the contract are the real
 * interface; the runner reads them when a probe returns nothing.
 *
 * The measurements bag handed to every objective/gate/precondition evaluator:
 *
 *   {
 *     a, b, c, v,        // parsed measure-a/b/c/v.json, or null if the probe failed
 *     model,             // the built DocumentModel (null if the build failed)
 *     publish,           // { ok, error?, documentPath? } from the publish-path build
 *     buildVariants,     // --full only: { lines3: aFacts|null, lines30: aFacts|null }
 *     repoTests,         // --full only: { ok, exitCode } from repo-root `npm test` (G5)
 *     full,              // whether --full was passed
 *     build,             // { ok, invariants: { I1..I4 } } — G6 grades these
 *   }
 *
 * `a.variants` is attached to probe A's facts when variant builds ran, since
 * the objectives that compare builds read them from there.
 */
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { PAGE_ONLY_STYLE } from "./lib/docx-page.mjs";
import { buildBorderFixture, buildVerificationArtifacts } from "./lib/verify-build.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const WEBSITE = resolve(HERE, "..");
const REPO = resolve(WEBSITE, "..");
const LIB = resolve(HERE, "lib");
const VERIFY = resolve(WEBSITE, ".verify");
const PREVIEW_DATA = resolve(WEBSITE, "src/demo/documents/invoice/preview-data.ts");

/** Where each probe's facts land — the contract's stable interface. */
const MEASURE_FILES = {
  A: resolve(VERIFY, "measure-a.json"),
  B: resolve(VERIFY, "measure-b.json"),
  C: resolve(VERIFY, "measure-c.json"),
  V: resolve(VERIFY, "measure-v.json"),
};

const PROBE_MODULES = {
  A: "probe-ooxml.mjs",
  B: "probe-preview.mjs",
  C: "probe-word.mjs",
  V: "probe-visual.mjs",
};

/** Generous ceilings: Word COM and PDF rasterising are slow, hangs are worse. */
const PROBE_TIMEOUT_MS = { A: 60_000, B: 240_000, C: 300_000, V: 300_000 };

const sha256 = (bytes) => createHash("sha256").update(bytes).digest("hex");

/** A short printable form for the board's `(measured vs expected)` column. */
function shown(value) {
  if (value === undefined || value === null) return "null";
  const text = typeof value === "string" ? value : JSON.stringify(value);
  return text.length > 72 ? `${text.slice(0, 69)}...` : text;
}

/** A promise with a ceiling, so one hung probe cannot hang the whole board. */
function withTimeout(promise, ms, label) {
  let timer;
  const ceiling = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms / 1000}s`)), ms);
  });
  return Promise.race([promise, ceiling]).finally(() => clearTimeout(timer));
}

/** Reads a JSON file, or null when it is missing or unparseable. */
async function readJson(path) {
  try {
    return JSON.parse(await readFile(path, "utf8"));
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Invariants
// ---------------------------------------------------------------------------

/**
 * I1 — no CSS patching. The live PAGE_ONLY_STYLE must still be byte-identical
 * to the frozen fixture; a drifted stylesheet means the preview measurements
 * could be flattering the packer, so every B/V objective becomes SUSPECT.
 */
async function checkI1() {
  let fixture;
  try {
    fixture = await readFile(resolve(LIB, "page-only-style.fixture"));
  } catch {
    return { value: "missing", note: "page-only-style.fixture not found" };
  }
  const live = sha256(Buffer.from(PAGE_ONLY_STYLE, "utf8"));
  const frozen = sha256(fixture);
  return live === frozen
    ? { value: "clean" }
    : { value: "dirty", note: `PAGE_ONLY_STYLE sha ${live.slice(0, 12)} != fixture ${frozen.slice(0, 12)}` };
}

/**
 * I4 — no shrinking the data. The numbers on the board are only meaningful if
 * the invoice still carries its longest strings and its real total; parse the
 * preview data at source and recompute the arithmetic the fixture shows.
 */
async function checkI4() {
  let source;
  try {
    source = await readFile(PREVIEW_DATA, "utf8");
  } catch {
    return { value: "dirty", note: "preview-data.ts unreadable" };
  }

  const linesBlock = source.match(/lines:\s*\[([\s\S]*?)\n\s*\]/)?.[1] ?? "";
  // Line entries are flat object literals, so a no-nesting match walks them.
  const entries = [...linesBlock.matchAll(/\{[^{}]*\}/g)].map(([entry]) => ({
    desc: entry.match(/desc:\s*"([^"]*)"/)?.[1] ?? "",
    meta: entry.match(/meta:\s*"([^"]*)"/)?.[1] ?? "",
    qty: Number(entry.match(/qty:\s*([\d.]+)/)?.[1] ?? NaN),
    rate: Number(entry.match(/rate:\s*([\d.]+)/)?.[1] ?? NaN),
  }));

  const vatRate = Number(source.match(/vatRate:\s*([\d.]+)/)?.[1] ?? NaN);
  const subtotal = entries.reduce((sum, line) => sum + line.qty * line.rate, 0);
  const total = Math.round(subtotal * (1 + vatRate));
  const maxDesc = Math.max(0, ...entries.map((line) => line.desc.length));
  const maxMeta = Math.max(0, ...entries.map((line) => line.meta.length));

  const failures = [];
  if (entries.length < 7) failures.push(`lines ${entries.length} < 7`);
  if (maxDesc < 30) failures.push(`max desc ${maxDesc} < 30`);
  if (maxMeta < 50) failures.push(`max meta ${maxMeta} < 50`);
  if (subtotal !== 18650) failures.push(`subtotal ${subtotal} != 18650`);
  if (total !== 22380) failures.push(`total ${total} != 22380`);

  return failures.length === 0
    ? { value: "clean", detail: { lines: entries.length, maxDesc, maxMeta, subtotal, total } }
    : { value: "dirty", note: failures.join("; ") };
}

// ---------------------------------------------------------------------------
// Probes
// ---------------------------------------------------------------------------

/**
 * Finds a probe module's entry point without pinning its author to one export
 * name: `default`, a conventional name, or the module's only function.
 */
function probeEntry(mod) {
  const exact = [mod.default, mod.run, mod.probe, mod.measure, mod.runProbe];
  const named = exact.find((candidate) => typeof candidate === "function");
  if (named) return named;
  // A probe module ships helpers alongside its entry (rasterisers, PNG
  // readers), so "the only function" is not enough: prefer a name that says
  // probe, then one that says run.
  const fns = Object.entries(mod).filter(([, value]) => typeof value === "function");
  const byName = (pattern) => fns.find(([key]) => pattern.test(key))?.[1] ?? null;
  return byName(/probe/i) ?? byName(/^(run|main|collect)/i) ?? (fns.length === 1 ? fns[0][1] : null);
}

/**
 * Runs one probe against one build. Returns `{ ok, facts, note }`; facts come
 * from the probe's return value, falling back to the measure file the
 * contract says every probe writes.
 */
async function runProbe(letter, context) {
  const modulePath = resolve(LIB, PROBE_MODULES[letter]);
  let mod;
  try {
    mod = await import(pathToFileURL(modulePath).href);
  } catch (error) {
    const missingSelf =
      error?.code === "ERR_MODULE_NOT_FOUND" && String(error.message).includes(PROBE_MODULES[letter]);
    return { ok: false, facts: null, note: missingSelf ? "not yet authored" : `import failed: ${error.message}` };
  }

  const entry = probeEntry(mod);
  if (entry === null) {
    return { ok: false, facts: null, note: "no callable export found" };
  }

  try {
    const returned = await withTimeout(
      Promise.resolve(entry(context)),
      PROBE_TIMEOUT_MS[letter],
      `probe ${letter}`,
    );
    const facts = returned && typeof returned === "object"
      ? returned
      : await readJson(MEASURE_FILES[letter]);
    if (facts === null) {
      return { ok: false, facts: null, note: "probe returned nothing and wrote no measure file" };
    }
    // A probe that reports its own failure is a failed probe: its objectives
    // must see null, not half-measured facts.
    if (facts.ok === false) {
      return { ok: false, facts: null, note: String(facts.reason ?? facts.error ?? "probe reported ok:false") };
    }
    return { ok: true, facts };
  } catch (error) {
    if (letter === "C") stopStrayWord();
    return { ok: false, facts: null, note: String(error?.message ?? error) };
  }
}

/**
 * The contract's sanctioned last resort for a failed Word probe: a WINWORD
 * left behind by a crashed COM session would wedge every later run.
 */
function stopStrayWord() {
  try {
    spawnSync("taskkill", ["/F", "/IM", "WINWORD.EXE"], { stdio: "ignore", timeout: 10_000 });
  } catch {
    // Nothing to kill, or nothing we can do — either way the report stands.
  }
}

// ---------------------------------------------------------------------------
// Objectives
// ---------------------------------------------------------------------------

/** Accepts arrays or keyed objects, so objectives.mjs can use either shape. */
function asDefinitionList(collection) {
  if (Array.isArray(collection)) return collection.filter((def) => def && typeof def === "object");
  if (collection && typeof collection === "object") {
    return Object.entries(collection).map(([id, def]) =>
      typeof def === "function" ? { id, evaluate: def } : { id, ...def },
    );
  }
  return [];
}

/** The evaluator of a definition, wherever its author put it. */
function evaluatorOf(def) {
  const candidates = [def.evaluate, def.eval, def.check, def.fn];
  return candidates.find((candidate) => typeof candidate === "function") ?? null;
}

const VALID_STATUSES = new Set(["PASS", "FAIL", "BLOCKED", "SUSPECT"]);

/**
 * Runs one definition against the measurements and pins the result to the
 * four contract statuses. A thrown evaluator is a FAIL with the message as
 * evidence — an objective that cannot run has not passed.
 */
function evaluate(def, m) {
  const id = def.id ?? "?";
  const base = {
    id,
    gap: def.gap ?? id.split(".")[0],
    description: def.description ?? def.desc ?? def.title ?? "",
    status: "FAIL",
    measured: null,
    expected: def.expected ?? null,
    note: "",
  };

  const evaluator = evaluatorOf(def);
  if (evaluator === null) return { ...base, note: "no evaluator" };

  try {
    const result = evaluator(m);
    if (typeof result === "boolean") return { ...base, status: result ? "PASS" : "FAIL" };
    if (result && typeof result === "object") {
      const status = VALID_STATUSES.has(result.status)
        ? result.status
        : (result.pass ?? result.ok) ? "PASS" : "FAIL";
      return {
        ...base,
        ...result,
        status,
        measured: result.measured ?? null,
        expected: result.expected ?? base.expected,
        note: result.note ?? "",
      };
    }
    return { ...base, note: `evaluator returned ${typeof result}` };
  } catch (error) {
    return { ...base, note: `evaluator threw: ${error?.message ?? error}` };
  }
}

/** The probe letter an objective id names: F1.B2 -> "B". */
const probeLetterOf = (id) => id.split(".")[1]?.[0] ?? "";

// ---------------------------------------------------------------------------
// The run
// ---------------------------------------------------------------------------

async function main() {
  const full = process.argv.includes("--full");
  const startedAt = new Date().toISOString();
  const runId = `vr-${Date.now()}`;

  // -- Invariants first: a tainted run should know it is tainted before it
  //    spends minutes measuring.
  const i1 = await checkI1();
  const i4 = await checkI4();

  let objectivesSha256 = null;
  try {
    objectivesSha256 = sha256(await readFile(resolve(LIB, "objectives.mjs")));
  } catch {
    // Missing objectives.mjs: the board still prints, with zero objectives.
  }

  const history = (await readJson(resolve(VERIFY, "history.json"))) ?? { runs: [], attempts: [] };
  history.runs ??= [];
  const previousRun = history.runs.at(-1) ?? null;
  const i2 =
    objectivesSha256 === null
      ? { value: "missing", note: "objectives.mjs not yet authored" }
      : previousRun === null || previousRun.objectivesSha256 === objectivesSha256
        ? { value: "clean" }
        : { value: "changed", note: "objectives.mjs hash differs from previous run" };

  // -- Build. A failed build still reports: every probe is then a failure with
  //    the build error as its note.
  let build = null;
  let buildError = null;
  try {
    build = await buildVerificationArtifacts({ publish: true });
  } catch (error) {
    buildError = String(error?.message ?? error);
  }
  const publish = build?.publish ?? { ok: false, error: buildError ?? "build did not run" };

  // -- --full extras: variant builds probed by A (F8.A3), and the repo tests
  //    (G5). Variants run before the main probes so measure-a.json ends the
  //    run describing the main build, not a variant.
  let buildVariants = null;
  let repoTests = null;
  if (full) {
    buildVariants = {};
    for (const [key, lines] of [["lines3", 3], ["lines30", 30]]) {
      try {
        const variant = await buildVerificationArtifacts({ lines });
        const probed = await runProbe("A", {
          verifyDir: VERIFY,
          docxPath: variant.docxPath,
          htmlPath: variant.htmlPath,
          modelPath: variant.modelPath,
          model: variant.model,
        });
        buildVariants[key] = probed.ok ? probed.facts : null;
      } catch (error) {
        buildVariants[key] = null;
        console.error(`variant build --lines=${lines} failed: ${error?.message ?? error}`);
      }
    }

    // One command string: an args array plus `shell` trips DEP0190 on Windows.
    const tests = spawnSync("npm test", {
      cwd: REPO,
      shell: true,
      encoding: "utf8",
      timeout: 900_000,
    });
    // `documents:check` is the website's own check that every demo document
    // still builds; `astro check` is what the repo calls it.
    const documents = spawnSync("npm run check", {
      cwd: WEBSITE,
      shell: true,
      encoding: "utf8",
      timeout: 900_000,
    });
    repoTests = {
      ok: tests.status === 0,
      exitCode: tests.status,
      documentsOk: documents.status === 0,
      documentsExitCode: documents.status,
    };
  }

  // -- Probes, C before V: V rasterises the PDF that C exports.
  const context = {
    verifyDir: VERIFY,
    docxPath: build?.docxPath ?? resolve(VERIFY, "invoice.docx"),
    htmlPath: build?.htmlPath ?? resolve(VERIFY, "invoice.html"),
    modelPath: build?.modelPath ?? resolve(VERIFY, "model.json"),
    model: build?.model ?? null,
    pdfPath: resolve(VERIFY, "invoice.pdf"),
  };

  const probes = {};
  for (const letter of ["A", "B", "C", "V"]) {
    probes[letter] = buildError === null
      ? await runProbe(letter, context)
      : { ok: false, facts: null, note: `build failed: ${buildError}` };
  }

  const invariants = { I1: i1.value, I2: i2.value, I3: "clean", I4: i4.value };

  // -- Objectives. The measurements bag is the only thing evaluators see;
  //    a failed probe hands them null, and they answer FAIL, not throw.
  const m = {
    a: probes.A.facts,
    b: probes.B.facts,
    c: probes.C.facts,
    v: probes.V.facts,
    model: build?.model ?? null,
    publish,
    buildVariants,
    repoTests,
    full,
    // G6 reads the invariants back as a measurement, so the run can fail on a
    // tainted input rather than only annotating one. The publish result rides
    // here too — the D13/D14 objectives read the published document from it.
    build: {
      ok: buildError === null,
      invariants,
      publish,
      // G5's shape: the two green checks a release depends on, named as the
      // objective names them.
      fullChecks: repoTests === null ? undefined : {
        testsOk: repoTests.ok,
        documentsCheckOk: repoTests.documentsOk,
      },
    },
  };

  // The published `document.json` itself, so an objective can ask what
  // travelled rather than only whether the build survived.
  if (publish.ok === true && typeof publish.documentPath === "string") {
    try {
      m.build.publish = { ...publish, documentJson: await readFile(publish.documentPath, "utf8") };
    } catch (error) {
      console.error(`published document unreadable: ${error?.message ?? error}`);
    }
  }

  // A variant build is probe A's facts about another build of the same
  // document. The objectives that compare builds read them off the main facts,
  // so they are attached where those objectives look.
  if (m.a !== null && buildVariants !== null) {
    m.a.variants = buildVariants;
  }

  // The publish-path docx, measured the same way the ordinary one is. A
  // decision baked at build time survives the first build and dies here, so
  // this is where "it still works when published" is actually asked.
  if (m.a !== null && typeof publish.docxPath === "string") {
    const probed = await runProbe("A", { ...context, docxPath: publish.docxPath });
    m.a.variants = { ...(m.a.variants ?? {}), publish: probed.ok ? probed.facts : null };
  }

  // F9's border fixture is a packer question, not an invoice one — its own
  // tiny document, packed and read back, so no theme or row separator can be
  // mistaken for the block's own edges.
  if (m.a !== null) {
    try {
      m.a.borderFixture = await buildBorderFixture();
    } catch (error) {
      m.a.borderFixture = null;
      console.error(`border fixture build failed: ${error?.message ?? error}`);
    }
  }

  let definitions = { objectives: [], gates: [], preconditions: [], note: null };
  try {
    const mod = await import(pathToFileURL(resolve(LIB, "objectives.mjs")).href);
    definitions = {
      objectives: asDefinitionList(mod.OBJECTIVES ?? mod.objectives),
      gates: asDefinitionList(mod.GATES ?? mod.gates),
      preconditions: asDefinitionList(mod.PRECONDITIONS ?? mod.preconditions),
      note: null,
    };
  } catch (error) {
    definitions.note =
      error?.code === "ERR_MODULE_NOT_FOUND" && String(error.message).includes("objectives.mjs")
        ? "objectives.mjs not yet authored"
        : `objectives.mjs failed to load: ${error?.message ?? error}`;
  }

  // Preconditions first — P0's verdict gates every visual objective.
  const preconditions = {};
  for (const def of definitions.preconditions) {
    const result = evaluate(def, m);
    preconditions[result.id] = { ...result, id: undefined, gap: undefined };
  }
  const p0Pass = preconditions.P0?.status === "PASS";

  const objectives = definitions.objectives.map((def) => evaluate(def, m));
  const gates = definitions.gates.map((def) => evaluate(def, m));

  // -- blockedBy: a dependency that has not happened makes the objective
  //    unmeasurable, which is BLOCKED, never PASS and never FAIL.
  for (const entry of [...objectives, ...gates]) {
    const def = [...definitions.objectives, ...definitions.gates].find((d) => d.id === entry.id);
    const blockedBy = def?.blockedBy;
    if (blockedBy === "publish-build" && publish.ok !== true) {
      entry.status = "BLOCKED";
      entry.note = `publish build failed: ${publish.error ?? "unknown"}`;
    } else if (blockedBy === "P0" && !p0Pass) {
      entry.status = "BLOCKED";
      entry.note = "P0 not PASS";
    }
  }

  // -- SUSPECT: a green that an invariant casts doubt on is not a green.
  //    Only PASS is downgraded — a FAIL under a tainted input is still not a
  //    pass, and BLOCKED already says the measurement never happened.
  if (i1.value !== "clean") {
    for (const entry of objectives) {
      if (entry.status === "PASS" && ["B", "V"].includes(probeLetterOf(entry.id))) {
        entry.status = "SUSPECT";
        entry.note = `I1 ${i1.value}: ${i1.note ?? "stylesheet drifted"}`;
      }
    }
  }
  if (i2.value === "changed" && previousRun !== null) {
    for (const entry of [...objectives, ...gates]) {
      if (entry.status === "PASS" && previousRun.statuses?.[entry.id] === "FAIL") {
        entry.status = "SUSPECT";
        entry.note = "threshold changed in same run";
      }
    }
  }
  if (i4.value !== "clean") {
    for (const entry of [...objectives, ...gates]) {
      const suspect = entry.gap === "F7" || entry.id === "G1" || entry.id === "G2";
      if (suspect && entry.status === "PASS") {
        entry.status = "SUSPECT";
        entry.note = `I4 dirty: ${i4.note ?? "preview data shrunk"}`;
      }
    }
  }

  // -- The board.
  const everyEntry = [
    ...objectives,
    ...gates,
    ...Object.values(preconditions).map((entry) => ({ status: entry.status })),
  ];
  const summary = {
    pass: everyEntry.filter((entry) => entry.status === "PASS").length,
    fail: everyEntry.filter((entry) => entry.status === "FAIL").length,
    blocked: everyEntry.filter((entry) => entry.status === "BLOCKED").length,
    suspect: everyEntry.filter((entry) => entry.status === "SUSPECT").length,
  };

  const report = {
    runId,
    startedAt,
    finishedAt: new Date().toISOString(),
    objectivesSha256,
    preconditions,
    invariants,
    invariantNotes: {
      ...(i1.note ? { I1: i1.note } : {}),
      ...(i2.note ? { I2: i2.note } : {}),
      ...(i4.note ? { I4: i4.note } : {}),
      ...(definitions.note ? { objectives: definitions.note } : {}),
      ...(buildError ? { build: buildError } : {}),
    },
    probes: Object.fromEntries(
      Object.entries(probes).map(([letter, probe]) => [
        letter,
        { ok: probe.ok, ...(probe.note ? { note: probe.note } : {}) },
      ]),
    ),
    build: {
      ok: buildError === null,
      publish,
      ...(full
        ? {
            repoTests,
            variantsProbed: {
              lines3: buildVariants?.lines3 !== null,
              lines30: buildVariants?.lines30 !== null,
            },
          }
        : {}),
    },
    objectives: objectives.map(({ id, gap, status, measured, expected, note }) =>
      ({ id, gap, status, measured, expected, note })),
    gates: Object.fromEntries(
      gates.map(({ id, status, measured, expected, note }) =>
        [id, { status, measured, expected, ...(note ? { note } : {}) }]),
    ),
    summary,
  };

  await writeFile(resolve(VERIFY, "report.json"), JSON.stringify(report, null, 2), "utf8");

  // -- History: this run's statuses, so the next run can see FAIL->PASS flips
  //    (I2) without ever merging boards (I3).
  const statuses = {};
  for (const entry of [...objectives, ...gates]) statuses[entry.id] = entry.status;
  for (const [id, entry] of Object.entries(preconditions)) statuses[id] = entry.status;
  history.runs.push({ runId, objectivesSha256, statuses });
  history.attempts ??= [];
  await writeFile(resolve(VERIFY, "history.json"), JSON.stringify(history, null, 2), "utf8");

  // -- Console board.
  console.log(`run ${runId}${full ? " (--full)" : ""}`);
  console.log(
    `invariants: I1 ${invariants.I1} / I2 ${invariants.I2} / I3 clean / I4 ${invariants.I4}`,
  );
  console.log(`build: ${buildError === null ? "ok" : `FAILED — ${buildError}`}; publish: ${
    publish.ok ? "ok" : `FAILED — ${shown(publish.error)}`}`);
  for (const [letter, probe] of Object.entries(probes)) {
    console.log(`probe ${letter}: ${probe.ok ? "ok" : `FAILED — ${probe.note}`}`);
  }
  if (definitions.note) console.log(`objectives: ${definitions.note}`);
  for (const entry of Object.entries(preconditions)) {
    console.log(`${entry[1].status.padEnd(7)} ${entry[0].padEnd(9)} precondition`);
  }
  for (const entry of objectives) {
    console.log(`${entry.status.padEnd(7)} ${entry.id.padEnd(9)} ${entry.description ?? ""}  (${
      shown(entry.measured)} vs ${shown(entry.expected)})${entry.note ? `  [${entry.note}]` : ""}`);
  }
  for (const entry of gates) {
    console.log(`${entry.status.padEnd(7)} ${entry.id.padEnd(9)} ${entry.description ?? ""}  (${
      shown(entry.measured)} vs ${shown(entry.expected)})${entry.note ? `  [${entry.note}]` : ""}`);
  }
  console.log(
    `BOARD: ${summary.pass} pass / ${summary.fail} fail / ${summary.blocked} blocked / ${summary.suspect} suspect`,
  );

  // Exit 0 only for a wholly green single run with clean invariants — the
  // plan's definition of done, applied mechanically.
  const allPass =
    everyEntry.length > 0 &&
    everyEntry.every((entry) => entry.status === "PASS") &&
    Object.values(invariants).every((value) => value === "clean") &&
    Object.values(probes).every((probe) => probe.ok) &&
    definitions.note === null;
  process.exitCode = allPass ? 0 : 1;
}

main().catch((error) => {
  // A runner crash still owes the loop a diagnosis on stderr and a red exit.
  console.error("verify-invoice crashed:", error);
  process.exitCode = 1;
});
