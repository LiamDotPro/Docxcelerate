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

  // A cell's geometry, reduced to the same content-relative points a
  // paragraph's is. Word reports both against the sheet; every other tier
  // reports against the text column, and a frame mismatch is how a harness
  // invents a divergence that is really its own arithmetic.
  const relativeCell = (cell, table) => ({
    ...cell,
    table: table.path ?? table.index,
    x: cell.x === null || cell.x === undefined ? null : round(cell.x - setup.leftMargin),
    y: cell.y === null || cell.y === undefined ? null : round(cell.y - setup.topMargin),
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

    /**
     * One running strip, by the kind of page Word draws it on.
     *
     * `kind` is `primary`, `firstPage` or `evenPages`; `part` is `header` or
     * `footer`. Word keeps a slot for every combination whether the document
     * uses it or not, so `exists` is what separates a strip that is there from
     * a slot that could hold one.
     */
    furniture(kind, part) {
      const strip = measure.furniture?.[`${kind}.${part}`];
      return strip ?? { exists: null, text: null, lines: null, y: null, x: null };
    },

    /** How far the header sits from the top of the sheet, in points. */
    headerDistance: () => measure.pageSetup?.headerDistance ?? null,
    /** How far the footer sits from the bottom of the sheet, in points. */
    footerDistance: () => measure.pageSetup?.footerDistance ?? null,
    /** Whether Word was told the first page has furniture of its own. */
    differentFirstPage: () => measure.pageSetup?.differentFirstPageHeaderFooter ?? null,
    /** Whether Word was told left and right pages differ. */
    differentOddAndEven: () => measure.pageSetup?.oddAndEvenPagesHeaderFooter ?? null,
    // --- tables -------------------------------------------------------------
    //
    // Word counts a cell's paragraph among the document's paragraphs, and the
    // body slice above deliberately leaves it out. These are the same
    // questions asked of a cell, with the geometry reduced to the same
    // content-relative frame every other measurement here is in.

    /** Every table Word found in the body, in order. */
    tables: listOf(measure.tables),

    /** One table, counting from zero. */
    table(index = 0) {
      return tableView(listOf(measure.tables)[index] ?? missingTable(`#${index}`));
    },

    /**
     * The cell whose text contains this anchor, wherever it is.
     *
     * Nested tables included, and case-insensitively — a heading cell set in
     * `w:caps` comes back from Word in capitals whatever the file says.
     */
    cell(anchor) {
      const wanted = anchor.toLowerCase();

      for (const table of flatTables(listOf(measure.tables))) {
        for (const cell of listOf(table.cells)) {
          if ((cell.text ?? "").toLowerCase().includes(wanted)) {
            return relativeCell(cell, table);
          }
        }
      }

      return missingCell(anchor);
    },

    /** Every cell whose text contains this anchor, nested tables included. */
    cells(anchor) {
      const wanted = anchor.toLowerCase();

      return flatTables(listOf(measure.tables)).flatMap((table) =>
        listOf(table.cells)
          .filter((cell) => (cell.text ?? "").toLowerCase().includes(wanted))
          .map((cell) => relativeCell(cell, table))
      );
    },

    /** One cell by where it sits: table, then row and column of the grid. */
    cellAt(table, row, column) {
      const found = flatTables(listOf(measure.tables))[table];
      if (found === undefined) return missingCell(`#${table}.${row}.${column}`);

      const cell = listOf(found.cells)
        .find((entry) => entry.row === row && entry.column === column);

      return cell === undefined
        ? missingCell(`#${table}.${row}.${column}`)
        : relativeCell(cell, found);
    },

  };
}


/**
 * A list, whatever the probe made of it.
 *
 * A field the Word probe could not fill is absent rather than empty — a run
 * that failed before it reached the tables, an older `measure-c.json` read
 * back from disk — and a case asking about a table in one of those should get
 * "not found" rather than a crash, which is the rule everything else in this
 * harness follows. It also absorbs the one shape `ConvertTo-Json` can make of
 * a single-item collection, which the probe's own `@(...)` already guards
 * against; two cheap guards against a silent shape change are worth one bug
 * that only appears on a table with one row.
 */
function listOf(value) {
  if (Array.isArray(value)) return value;
  return value === null || value === undefined ? [] : [value];
}

/**
 * Every table in a list, the nested ones included, innermost first.
 *
 * Word reports a cell's range as everything inside it, a nested table's words
 * included, so an outer cell matches every anchor its inner table matches.
 * The inner cell is the more specific of the two and is what an anchor names,
 * so the deepest tables are searched first. Probe A orders its own the same
 * way, which is what keeps one anchor meaning one cell on both sides.
 */
function flatTables(tables) {
  return tables.flatMap((table) => [...flatTables(listOf(table.nested)), table]);
}

/** A table with its rows and cells one step away, in the grid's own terms. */
function tableView(table) {
  return {
    ...table,
    cells: listOf(table.cells),
    rows: listOf(table.rows),

    /** One row's properties — what it repeats as, and how tall it is. */
    row(index = 0) {
      return listOf(table.rows)[index] ?? { index: -1, missing: true, headingFormat: null };
    },

    /** One cell, by the row and column of the grid it occupies. */
    cell(row = 0, column = 0) {
      const found = listOf(table.cells)
        .find((entry) => entry.row === row && entry.column === column);

      return found === undefined ? missingCell(`#${row}.${column}`) : found;
    },

    /** Which pages the table's cells landed on, in order, without repeats. */
    pages() {
      return [...new Set(listOf(table.cells).map((cell) => cell.page).filter((page) => page))];
    },
  };
}

/**
 * The stand-ins for a table and a cell Word did not find.
 *
 * Null rather than undefined throughout, so an assertion about a table the
 * packer never wrote reports the property it wanted rather than dying on it.
 */
function missingTable(anchor) {
  return {
    missing: true,
    anchor,
    index: -1,
    path: null,
    depth: null,
    rowCount: null,
    columnCount: null,
    uniform: null,
    style: null,
    preferredWidth: null,
    preferredWidthType: null,
    alignment: null,
    leftIndent: null,
    wrapAroundText: null,
    allowAutoFit: null,
    page: null,
    x: null,
    y: null,
    rows: [],
    cells: [],
    nested: [],
  };
}

function missingCell(anchor) {
  return {
    missing: true,
    anchor,
    row: null,
    column: null,
    text: "",
    width: null,
    preferredWidth: null,
    height: null,
    heightRule: null,
    vAlign: null,
    shading: null,
    padding: null,
    borders: null,
    alignment: null,
    page: null,
    x: null,
    y: null,
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
