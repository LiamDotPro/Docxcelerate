/**
 * Probe C wrapper — runs probe-word.ps1 (Word over COM) and lands its JSON as
 * `.verify/measure-c.json`.
 *
 * Word is the one probe that can hang rather than fail: a licensing dialog, a
 * corrupt open, a stuck repagination all block COM forever with no exception
 * for the ps1's try/catch to see. So the wrapper owns the clock — 180s, then
 * kill the PowerShell tree AND any WINWORD.EXE left behind, because an
 * orphaned hidden Word poisons every later run on this machine.
 *
 * The ps1 emits one line of compact JSON on stdout, but Word/PowerShell can
 * precede it with noise (ActiveX warnings, progress spew), so we parse the
 * LAST line that parses rather than trusting the whole stream.
 *
 * See VERIFY-CONTRACT.md ("measure-c.json") for the output schema.
 */
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..", "..");
const VERIFY_DIR = resolve(ROOT, ".verify");
const PS1_PATH = resolve(HERE, "probe-word.ps1");
const TIMEOUT_MS = 180_000;

/** Runs a command, collecting stdout/stderr, with a hard timeout. */
function run(command, args, timeoutMs) {
  return new Promise((resolvePromise) => {
    const child = spawn(command, args, { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";
    let timedOut = false;

    const timer = timeoutMs
      ? setTimeout(() => {
          timedOut = true;
          // /T takes the whole tree: powershell AND the WINWORD it spawned.
          spawn("taskkill", ["/PID", String(child.pid), "/T", "/F"], { windowsHide: true })
            .on("error", () => {});
        }, timeoutMs)
      : null;

    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", (error) => {
      if (timer) clearTimeout(timer);
      resolvePromise({ code: -1, stdout, stderr: `${stderr}\n${error.message}`, timedOut });
    });
    child.on("close", (code) => {
      if (timer) clearTimeout(timer);
      resolvePromise({ code, stdout, stderr, timedOut });
    });
  });
}

/**
 * Kills any WINWORD.EXE by name. Last resort per the contract: the ps1's
 * finally block is the polite path, this is for when it never ran. taskkill
 * exit 128 (no such process) is the normal, ignorable case.
 */
async function killWord() {
  await run("taskkill", ["/IM", "WINWORD.EXE", "/F", "/T"], 15_000);
}

/** The last stdout line that parses as a JSON object — see module comment. */
function parseLastJsonLine(stdout) {
  const lines = stdout.split(/\r?\n/).filter((line) => line.trim().length > 0);
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i].trim();
    if (!line.startsWith("{")) continue;
    try {
      return JSON.parse(line);
    } catch {
      // Not the JSON line; keep walking up.
    }
  }
  return null;
}

/**
 * Runs probe C and writes `.verify/measure-c.json`.
 *
 * On failure it still writes the file (with `ok: false` and the evidence) so
 * the runner can mark every C objective FAIL with a note instead of crashing,
 * then throws for callers that treat a dead probe as fatal.
 *
 * @param {{ docxPath?: string, pdfPath?: string, outPath?: string }} options
 * @returns The parsed measurement object.
 */
export async function measure(options = {}) {
  const docxPath = resolve(options.docxPath ?? resolve(VERIFY_DIR, "invoice.docx"));
  const pdfPath = resolve(options.pdfPath ?? resolve(VERIFY_DIR, "invoice.pdf"));
  const outPath = resolve(options.outPath ?? resolve(VERIFY_DIR, "measure-c.json"));

  if (!existsSync(docxPath)) {
    throw new Error(`probe-word: ${docxPath} does not exist — run verify-build.mjs first`);
  }
  await mkdir(dirname(outPath), { recursive: true });

  const result = await run("powershell", [
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", PS1_PATH,
    "-DocxPath", docxPath,
    "-PdfPath", pdfPath,
  ], TIMEOUT_MS);

  if (result.timedOut) {
    await killWord();
    const failure = {
      probe: "C",
      ok: false,
      error: `probe-word.ps1 timed out after ${TIMEOUT_MS / 1000}s; WINWORD killed`,
      stderr: result.stderr.slice(-2000),
    };
    await writeFile(outPath, JSON.stringify(failure, null, 2), "utf8");
    throw new Error(failure.error);
  }

  const measurement = parseLastJsonLine(result.stdout);
  if (!measurement) {
    // The ps1 died before its JSON line — its finally may not have run, so
    // sweep up any hidden Word before reporting.
    await killWord();
    const failure = {
      probe: "C",
      ok: false,
      error: `probe-word.ps1 exited ${result.code} without emitting JSON`,
      stderr: result.stderr.slice(-2000),
      stdout: result.stdout.slice(-2000),
    };
    await writeFile(outPath, JSON.stringify(failure, null, 2), "utf8");
    throw new Error(failure.error);
  }

  await writeFile(outPath, JSON.stringify(measurement, null, 2), "utf8");
  return measurement;
}

// CLI: node scripts/lib/probe-word.mjs [--docx=path] [--pdf=path]
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const options = {};
  for (const arg of process.argv.slice(2)) {
    if (arg.startsWith("--docx=")) options.docxPath = arg.slice(7);
    if (arg.startsWith("--pdf=")) options.pdfPath = arg.slice(6);
  }
  measure(options).then((measurement) => {
    console.log(JSON.stringify(measurement, null, 2));
  }, (error) => {
    console.error("probe-word failed:", error.message ?? error);
    process.exitCode = 1;
  });
}
