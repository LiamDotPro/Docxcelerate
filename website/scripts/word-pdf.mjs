/**
 * Has Microsoft Word for Mac open a .docx and export it as a PDF, so the file
 * itself — not a preview of it — can be scored against a design.
 *
 *   node scripts/word-pdf.mjs .parity/docx/portfolio-review.engine.docx
 *
 * Writes `<dir>/word/<name>.pdf` beside the .docx and prints Word's own page
 * count. Word is sandboxed and asks before writing anywhere new, so the PDF is
 * saved into Word's container and copied out. The document is addressed by
 * name and closed unsaved afterwards; no other open document is touched.
 *
 * @module
 */

import { execFile } from "node:child_process";
import { copyFile, mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { promisify } from "node:util";

const run = promisify(execFile);
const source = resolve(process.argv[2] ?? "");
const name = basename(source);
const stem = name.replace(/\.docx$/, "");
const container = join(homedir(), "Library/Containers/com.microsoft.Word/Data/Documents");
const inside = join(container, `${stem}.pdf`);

const script = `
with timeout of 240 seconds
  tell application "Microsoft Word"
    if (exists document ${JSON.stringify(name)}) then close document ${JSON.stringify(name)} saving no
    open (POSIX file ${JSON.stringify(source)})
    delay 3
    set doc to document ${JSON.stringify(name)}
    set pages to (compute statistics doc statistic statistic pages)
    save as doc file name ${JSON.stringify(inside)} file format format PDF
    close doc saving no
    return pages
  end tell
end timeout`;

const { stdout } = await run("osascript", ["-e", script], { timeout: 300_000 });
const out = join(dirname(source), "word", `${stem}.pdf`);
await mkdir(dirname(out), { recursive: true });
await copyFile(inside, out);
console.log(`${name}: ${stdout.trim()} pages in Word -> ${out}`);
