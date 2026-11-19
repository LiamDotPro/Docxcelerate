/**
 * Where a page's vertical drift is introduced, rather than where it shows.
 *
 *   node scripts/parity-steps.mjs portfolio-review 3 [--engine]
 *
 * Reads the text drift `design-parity.mjs` wrote and prints, in the canvas's
 * reading order, each run's drift and the change from the run above it. A
 * spacing mistake shows as one step; everything under it then carries it.
 */
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

const [id, page, flag] = process.argv.slice(2);
const dir = resolve(import.meta.dirname, "..", ".parity", `${id}${flag === "--engine" ? ".engine" : ""}`);
const drift = JSON.parse(await readFile(resolve(dir, "text-drift.json"), "utf8"))[Number(page) - 1];
let last = 0;
for (const m of drift.moved) {
  const step = +(m.dy - last).toFixed(1);
  last = m.dy;
  const flagged = Math.abs(step) > 1 ? " ◀" : "";
  console.log(`${String(m.dy).padStart(7)} ${String(step).padStart(7)}${flagged.padEnd(3)} ${m.text.padEnd(44)} ${m.style.join("; ")}`);
}
console.log(`missing: ${drift.missing.join(" | ")}`);
console.log(`extra:   ${drift.extra.join(" | ")}`);
