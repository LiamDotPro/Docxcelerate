/**
 * Probe C — the wrapper that runs Word and lands its JSON.
 *
 * Word is the one probe that can hang rather than fail: a licensing dialog, a
 * corrupt open, a stuck repagination all block COM forever with no exception
 * for the script's own try/catch to see. So the clock lives here — a hard
 * timeout, then kill the PowerShell tree AND any WINWORD left behind, because
 * an orphaned hidden Word poisons every later run on the machine.
 *
 * @module
 */

import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { PT_PER_MM } from "./case.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const PS1 = resolve(HERE, "probe-word.ps1");
const TIMEOUT_MS = 180_000;

/** Runs a command, collecting output, with a hard timeout. */
function run(command, args, timeoutMs) {
  return new Promise((done) => {
    const child = spawn(command, args, { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";
    let timedOut = false;

    const timer = timeoutMs
      ? setTimeout(() => {
        timedOut = true;
        // /T takes the whole tree: PowerShell and the WINWORD it spawned.
        spawn("taskkill", ["/PID", String(child.pid), "/T", "/F"], { windowsHide: true })
          .on("error", () => {});
      }, timeoutMs)
      : null;

    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", (error) => {
      if (timer) clearTimeout(timer);
      done({ code: -1, stdout, stderr: `${stderr}\n${error.message}`, timedOut });
    });
    child.on("close", (code) => {
      if (timer) clearTimeout(timer);
      done({ code, stdout, stderr, timedOut });
    });
  });
}

async function killWord() {
  await run("taskkill", ["/IM", "WINWORD.EXE", "/F", "/T"], 15_000);
}

/**
 * The last stdout line that parses as JSON.
 *
 * Word and PowerShell can precede the payload with noise — ActiveX warnings,
 * progress spew — so the stream is walked from the end rather than trusted
 * whole.
 */
function parseLastJsonLine(stdout) {
  const lines = stdout.split(/\r?\n/).filter((line) => line.trim().length > 0);

  for (let index = lines.length - 1; index >= 0; index -= 1) {
    const line = lines[index].trim();
    if (!line.startsWith("{")) continue;
    try {
      return JSON.parse(line);
    } catch {
      // Not the JSON line; keep walking up.
    }
  }

  return null;
}

/** Whether this machine can run the Word tier at all. */
export function wordAvailable() {
  return process.platform === "win32" && existsSync(PS1);
}

/**
 * Measures one case in Word and writes `measure-c.json` beside it.
 *
 * On failure it still writes the file, with the evidence, so the runner can
 * mark the tier FAIL with a note rather than crashing the board.
 */
export async function runWordProbe({ outDir, docxPath }) {
  const pdfPath = resolve(outDir, "case.pdf");
  const outPath = resolve(outDir, "measure-c.json");

  const result = await run("powershell", [
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", PS1,
    "-DocxPath", docxPath,
    "-PdfPath", pdfPath,
  ], TIMEOUT_MS);

  if (result.timedOut) {
    await killWord();
    const failure = {
      probe: "C",
      ok: false,
      error: `the Word probe timed out after ${TIMEOUT_MS / 1000}s; WINWORD killed`,
      stderr: result.stderr.slice(-2000),
    };
    await writeFile(outPath, JSON.stringify(failure, null, 2), "utf8");
    return failure;
  }

  const parsed = parseLastJsonLine(result.stdout);

  if (parsed === null) {
    const failure = {
      probe: "C",
      ok: false,
      error: "the Word probe emitted no JSON",
      stdout: result.stdout.slice(-2000),
      stderr: result.stderr.slice(-2000),
    };
    await writeFile(outPath, JSON.stringify(failure, null, 2), "utf8");
    return failure;
  }

  await writeFile(outPath, JSON.stringify(parsed, null, 2), "utf8");

  return parsed;
}

/**
 * The view a case is handed. Everything is in points, page-relative reduced to
 * content-relative — the same frame the preview reports in, so the two can be
 * compared without either side doing arithmetic in an assertion.
 */
export function wordView(measure) {
  const setup = measure.pageSetup ?? { leftMargin: 0, topMargin: 0 };

  // Word counts a paragraph inside a table too; the body slice does not.
  const body = (measure.paragraphs ?? []).filter((p) => p.inTable !== true);

  // Case-insensitive, because `w:caps` means the file's case and the printed
  // case differ: a label written "Invoice reference" comes back from Word as
  // "INVOICE REFERENCE". An anchor names what a reader sees, and a reader sees
  // both spellings as the same words.
  const find = (anchor) => {
    const wanted = anchor.toLowerCase();
    return body.find((p) => (p.text ?? "").toLowerCase().includes(wanted)) ??
      missingParagraph(anchor);
  };

  const relative = (para) => ({
    ...para,
    // Page-relative points, less the margin, is content-relative points.
    x: para.x === null || para.x === undefined ? null : round(para.x - setup.leftMargin),
    y: para.y === null || para.y === undefined ? null : round(para.y - setup.topMargin),
    xEnd: para.xEnd === null || para.xEnd === undefined ? null : round(para.xEnd - setup.leftMargin),
  });

  return {
    ...measure,
    paragraphs: body,

    /** Millimetres, in the points everything here is measured in. */
    mm: (value) => value * PT_PER_MM,

    para: (anchor) => relative(find(anchor)),

    /** The stops the document itself declared, without Word's default grid. */
    customTabStops: (anchor) => (find(anchor).tabStops ?? []).filter((stop) => stop.custom),
    paraAt: (index) => relative(body[index] ?? missingParagraph(`#${index}`)),

    pageCount: () => measure.pages ?? null,
  };
}

function round(value) {
  return Math.round(value * 100) / 100;
}

function missingParagraph(anchor) {
  return {
    missing: true,
    anchor,
    index: -1,
    text: "",
    inTable: null,
    alignment: null,
    leftIndent: null,
    rightIndent: null,
    firstLineIndent: null,
    spaceBefore: null,
    spaceAfter: null,
    lineSpacing: null,
    lineSpacingRule: null,
    keepWithNext: null,
    keepTogether: null,
    widowControl: null,
    pageBreakBefore: null,
    tabStops: [],
    tabStopCount: null,
    customTabStopCount: null,
    shading: null,
    borders: null,
    fontName: null,
    fontSize: null,
    bold: null,
    italic: null,
    allCaps: null,
    characterSpacing: null,
    fontColor: null,
    page: null,
    x: null,
    y: null,
    xEnd: null,
    yEnd: null,
    pageEnd: null,
    linesSplitAcrossPages: null,
    lineCount: null,
  };
}
