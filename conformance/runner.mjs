/**
 * The conformance runner: build every case, probe it, score it, print the board.
 *
 * Exit code is the whole contract. Zero means every case that claims to work
 * does, and every case that claims not to still does not. A gap we have written
 * down does not fail the build — otherwise the roadmap is a broken build — but
 * a gap that quietly closed is reported STALE, because a suite you can green by
 * deletion is not a suite.
 *
 * Usage:
 *   node runner.mjs                     every case, every tier that can run here
 *   node runner.mjs --case text/align   one case
 *   node runner.mjs --area text         one area
 *   node runner.mjs --tier a            the file only: no Chrome, no Word
 *   node runner.mjs --word              include the Word tiers
 *   node runner.mjs --json              the board as JSON, for the report
 *
 * @module
 */

import { readdir, writeFile } from "node:fs/promises";
import { relative, resolve, sep } from "node:path";
import { pathToFileURL } from "node:url";
import { buildCase, CASES_DIR, loadCase, OUT_DIR } from "./lib/build.mjs";
import { measureOoxml, ooxmlView } from "./lib/probe-ooxml.mjs";
import { chromePath, previewView, runPreviewProbe } from "./lib/probe-preview.mjs";
import { runWordProbe, wordAvailable, wordView } from "./lib/probe-word.mjs";
import { parityTable, parityView } from "./lib/parity.mjs";
import { rasterisePdf } from "./lib/raster.mjs";
import { recorder } from "./lib/case.mjs";

/** The tiers, in the order a case is scored in, with the unit each measures in. */
const TIERS = [
  { key: "ooxml", label: "A", unit: "raw", needs: "none" },
  { key: "preview", label: "B", unit: "px", needs: "chrome" },
  { key: "word", label: "C", unit: "pt", needs: "word" },
  { key: "parity", label: "X", unit: "mm", needs: "word" },
];

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const files = await findCases(options);

  if (files.length === 0) {
    console.error("no cases matched");
    process.exitCode = 1;
    return;
  }

  const chrome = chromePath();
  const canPreview = options.tiers.includes("preview") && chrome !== null;
  const canWord = options.word && wordAvailable();

  if (options.tiers.includes("preview") && chrome === null) {
    console.log("! Chrome not found — the preview tier is SKIPPED (set CHROME_PATH)\n");
  }
  if (options.word && !wordAvailable()) {
    console.log("! Word not available here — the Word and parity tiers are SKIPPED\n");
  }

  const board = [];

  for (const file of files) {
    board.push(await runOne(file, { canPreview, canWord, options }));
  }

  await writeFile(
    resolve(OUT_DIR, "report.json"),
    JSON.stringify({ generatedAt: new Date().toISOString(), cases: board }, null, 2),
    "utf8",
  );

  if (options.json) {
    console.log(JSON.stringify(board, null, 2));
  }

  printBoard(board);

  const bad = board.filter((entry) => entry.status === "FAIL" || entry.status === "STALE");
  process.exitCode = bad.length === 0 ? 0 : 1;
}

/**
 * One case, from source file to a scored entry.
 *
 * A tier that cannot run here is SKIPPED, never silently green — the whole
 * value of the board is that a green line means something was checked.
 */
async function runOne(file, { canPreview, canWord, options }) {
  const id = relative(CASES_DIR, file).replaceAll(sep, "/").replace(/\.case\.tsx$/, "");
  const entry = { id, file: relative(process.cwd(), file), tiers: {}, assertions: [] };

  let spec;
  try {
    spec = await loadCase(file);
  } catch (error) {
    entry.status = "FAIL";
    entry.error = `case would not load: ${error?.message ?? error}`;
    return entry;
  }

  entry.title = spec.title;
  entry.feature = spec.feature ?? null;
  entry.word = spec.word ?? null;
  entry.claim = spec.claim;
  entry.knownRed = spec.knownRed ?? [];

  let artefacts;
  try {
    artefacts = await buildCase(spec);
  } catch (error) {
    entry.status = "FAIL";
    entry.error = `case would not build: ${error?.message ?? error}`;
    return entry;
  }

  entry.out = relative(process.cwd(), artefacts.outDir);

  // --- tier A: the file ----------------------------------------------------

  if (options.tiers.includes("ooxml")) {
    await score(entry, spec, "ooxml", "raw", async () => {
      const measure = measureOoxml(artefacts.bytes);
      await writeFile(
        resolve(artefacts.outDir, "measure-a.json"),
        JSON.stringify(
          { ...measure, documentXml: undefined, stylesXml: undefined, settingsXml: undefined },
          null,
          2,
        ),
        "utf8",
      );
      return [ooxmlView(measure)];
    });
  }

  // --- tier B: the preview -------------------------------------------------

  let preview = null;

  if (options.tiers.includes("preview")) {
    if (!canPreview) {
      entry.tiers.preview = { status: "SKIPPED", note: "Chrome not available" };
    } else {
      await score(entry, spec, "preview", "px", async () => {
        const measure = await runPreviewProbe({
          outDir: artefacts.outDir,
          htmlPath: artefacts.htmlPath,
          bodyFont: artefacts.model.style?.typography?.bodyFont,
          screenshots: options.screenshots,
        });
        if (measure.error !== null) {
          throw new Error(measure.error);
        }
        preview = previewView(measure);
        return [preview];
      });
    }
  }

  // --- tiers C and X: Word, and the two engines against each other ----------

  let word = null;

  if (options.tiers.includes("word")) {
    if (!canWord) {
      const note = options.word ? "Word is not available here" : "Word tiers not requested (--word)";
      entry.tiers.word = { status: "SKIPPED", note };
    } else {
      await score(entry, spec, "word", "pt", async () => {
        const measure = await runWordProbe({ outDir: artefacts.outDir, docxPath: artefacts.docxPath });
        if (measure.ok !== true) {
          throw new Error(measure.error ?? "the Word probe reported failure");
        }
        word = wordView(measure);
        return [word];
      });

      // Word's PDF export is what the board's second screenshot is made of.
      // A missing picture costs the board a column rather than a run, so this
      // never throws — but it does say why, because a column that goes quietly
      // missing is one nobody investigates.
      if (options.screenshots) {
        try {
          entry.wordScreenshots = await rasterisePdf(artefacts.outDir);
        } catch (error) {
          entry.wordScreenshots = [];
          entry.rasterError = String(error?.message ?? error);
        }
      }
    }
  }

  if (options.tiers.includes("parity")) {
    if (preview === null || word === null) {
      entry.tiers.parity = {
        status: "SKIPPED",
        note: preview === null ? "no preview measurement to compare" : "no Word measurement to compare",
      };
    } else {
      const view = parityView(preview, word);
      await score(entry, spec, "parity", "mm", async () => [view]);
      // The survey, not the assertions: every paragraph both engines found,
      // and how far apart they put it.
      entry.parityTable = parityTable(preview, word);
      await writeFile(
        resolve(artefacts.outDir, "measure-x.json"),
        JSON.stringify({ probe: "X", rows: entry.parityTable }, null, 2),
        "utf8",
      );
    }
  }

  entry.status = statusOf(entry, spec);

  await writeFile(
    resolve(artefacts.outDir, "result.json"),
    JSON.stringify(entry, null, 2),
    "utf8",
  );

  return entry;
}

/** Runs one tier's assertions and records what they found. */
async function score(entry, spec, tier, unit, measure) {
  const assert = spec.expect?.[tier];

  if (assert === undefined) {
    entry.tiers[tier] = { status: "SKIPPED", note: "the case makes no claim about this tier" };
    return;
  }

  const is = recorder(tier, unit);

  try {
    const views = await measure();
    await assert(...views, is);
  } catch (error) {
    entry.tiers[tier] = { status: "FAIL", note: `probe failed: ${error?.message ?? error}` };
    entry.assertions.push({
      id: `${tier}.0`,
      tier,
      ok: false,
      message: "the probe ran",
      measured: String(error?.message ?? error),
      expected: "a measurement",
    });
    return;
  }

  entry.assertions.push(...is.results);

  const failed = is.results.filter((result) => !result.ok);
  entry.tiers[tier] = {
    status: failed.length === 0 ? "PASS" : "FAIL",
    passed: is.results.length - failed.length,
    total: is.results.length,
  };
}

/**
 * What the case's claim makes of its tiers.
 *
 * This is the mechanism that turns a test suite into a roadmap. An unsupported
 * case that fails is the roadmap; one that passes is a promotion waiting to be
 * written down.
 */
function statusOf(entry, spec) {
  const ran = Object.entries(entry.tiers).filter(([, tier]) => tier.status !== "SKIPPED");
  if (ran.length === 0) return "SKIPPED";

  const failed = ran.filter(([, tier]) => tier.status === "FAIL");

  if (spec.claim === "supported") {
    return failed.length === 0 ? "PASS" : "FAIL";
  }

  if (spec.claim === "unsupported") {
    return failed.length === 0 ? "STALE" : "KNOWN";
  }

  // Partial: the tiers the case named may be red, and nothing else may be.
  const unexpected = failed.filter(([name]) => !spec.knownRed.includes(name));
  const expectedRed = spec.knownRed.filter((name) =>
    ran.some(([tier, value]) => tier === name && value.status === "FAIL")
  );

  if (unexpected.length > 0) return "FAIL";
  return expectedRed.length === spec.knownRed.length ? "KNOWN" : "STALE";
}

// ---------------------------------------------------------------------------
// The board
// ---------------------------------------------------------------------------

const MARK = {
  PASS: "PASS ",
  FAIL: "FAIL ",
  KNOWN: "KNOWN",
  STALE: "STALE",
  SKIPPED: "SKIP ",
};

function printBoard(board) {
  for (const entry of board) {
    console.log(`\n${MARK[entry.status] ?? entry.status}  ${entry.id}  —  ${entry.title ?? ""}`);

    if (entry.error !== undefined) {
      console.log(`       ${entry.error}`);
      continue;
    }

    for (const { key, label } of TIERS) {
      const tier = entry.tiers[key];
      if (tier === undefined) continue;
      const detail = tier.total === undefined
        ? tier.note ?? ""
        : `${tier.passed}/${tier.total}`;
      console.log(`       ${label} ${key.padEnd(8)} ${tier.status.padEnd(8)} ${detail}`);
    }

    for (const result of entry.assertions.filter((item) => !item.ok)) {
      console.log(`         · ${result.tier}: ${result.message}`);
      console.log(`             measured ${result.measured}   expected ${result.expected}`);
      if (result.note !== undefined) console.log(`             ${result.note}`);
    }
  }

  const tally = board.reduce((counts, entry) => {
    counts[entry.status] = (counts[entry.status] ?? 0) + 1;
    return counts;
  }, {});

  console.log(
    `\nBOARD: ${tally.PASS ?? 0} pass / ${tally.FAIL ?? 0} fail / ` +
    `${tally.KNOWN ?? 0} known / ${tally.STALE ?? 0} stale / ${tally.SKIPPED ?? 0} skipped`,
  );
}

// ---------------------------------------------------------------------------
// Arguments and discovery
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const options = {
    case: null,
    area: null,
    tiers: TIERS.map((tier) => tier.key),
    word: false,
    json: false,
    screenshots: true,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === "--case") options.case = argv[++index];
    else if (arg === "--area") options.area = argv[++index];
    else if (arg === "--word") options.word = true;
    else if (arg === "--json") options.json = true;
    else if (arg === "--no-screenshots") options.screenshots = false;
    else if (arg === "--tier") {
      const tier = String(argv[++index]).toLowerCase();
      const named = TIERS.find((entry) => entry.label.toLowerCase() === tier || entry.key === tier);
      options.tiers = named === undefined ? options.tiers : [named.key];
    }
  }

  return options;
}

async function findCases(options) {
  const found = [];

  async function walk(dir) {
    for (const item of await readdir(dir, { withFileTypes: true })) {
      const full = resolve(dir, item.name);
      if (item.isDirectory()) {
        if (item.name.startsWith("_")) continue;
        await walk(full);
      } else if (item.name.endsWith(".case.tsx")) {
        found.push(full);
      }
    }
  }

  await walk(CASES_DIR);
  found.sort();

  return found.filter((file) => {
    const id = relative(CASES_DIR, file).replaceAll(sep, "/").replace(/\.case\.tsx$/, "");
    if (options.case !== null && id !== options.case) return false;
    if (options.area !== null && !id.startsWith(`${options.area}/`)) return false;
    return true;
  });
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}
