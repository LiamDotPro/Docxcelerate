/**
 * The board: one page showing, per case, the code, what the preview drew, and
 * what Word made of it.
 *
 * The three columns are the whole point. A number that says the preview and
 * Word are 1.6mm apart is a fact; the two pictures side by side are what makes
 * somebody look at the fact. Everything is inlined — the page is one file, and
 * a report that needs a server beside it is a report nobody opens twice.
 *
 * @module
 */

import { mkdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { OUT_DIR, ROOT } from "./build.mjs";

/** A PNG as a data URI, or null when it was never taken. */
async function inlineImage(path) {
  if (!existsSync(path)) return null;
  const bytes = await readFile(path);
  return `data:image/png;base64,${bytes.toString("base64")}`;
}

/**
 * Writes `conformance/report/index.html` from the last run's report.
 *
 * @param {string} outPath Where to write. Defaults to the report directory.
 */
export async function writeBoard(outPath = resolve(ROOT, "report", "index.html")) {
  await mkdir(dirname(outPath), { recursive: true });

  const report = JSON.parse(await readFile(resolve(OUT_DIR, "report.json"), "utf8"));
  const sections = [];

  for (const entry of report.cases) {
    // The runner records paths relative to wherever it was invoked, which is
    // the conformance directory itself.
    const outDir = resolve(ROOT, entry.out ?? "");
    const source = entry.file === undefined
      ? ""
      : await readFile(resolve(ROOT, entry.file), "utf8").catch(() => "");

    const previewShots = [];
    const wordShots = [];

    for (let page = 1; page <= 4; page += 1) {
      const preview = await inlineImage(resolve(outDir, `preview-p${page}.png`));
      const word = await inlineImage(resolve(outDir, `word-p${page}.png`));
      if (preview !== null) previewShots.push(preview);
      if (word !== null) wordShots.push(word);
    }

    sections.push(renderCase(entry, source, previewShots, wordShots));
  }

  const tally = report.cases.reduce((counts, entry) => {
    counts[entry.status] = (counts[entry.status] ?? 0) + 1;
    return counts;
  }, {});

  const html = page(sections.join("\n"), tally, report.generatedAt);
  await writeFile(outPath, html, "utf8");

  return outPath;
}

/** Just the document, with the case's framing stripped — what a reader wants. */
function documentOf(source) {
  const match = /document: template\(\n([\s\S]*?)\n  \),/.exec(source);
  const body = match === null ? source : match[1];
  const styleMatch = /style: (withBlocks\(\{[\s\S]*?\n  \}\)),/.exec(source);

  return [
    styleMatch === null ? null : `// the theme's side of it\nstyle: ${styleMatch[1]},`,
    body.replace(/^ {4}/gm, ""),
  ].filter(Boolean).join("\n\n");
}

function renderCase(entry, source, previewShots, wordShots) {
  const tiers = ["ooxml", "preview", "word", "parity"]
    .filter((key) => entry.tiers?.[key] !== undefined)
    .map((key) => {
      const tier = entry.tiers[key];
      const detail = tier.total === undefined ? "" : `${tier.passed}/${tier.total}`;
      return `<span class="tier ${tier.status.toLowerCase()}"><b>${label(key)}</b> ${key} ${detail || tier.status}</span>`;
    })
    .join("");

  const failures = (entry.assertions ?? []).filter((item) => !item.ok);
  const failureRows = failures.length === 0
    ? ""
    : `<table class="fails"><thead><tr><th>tier</th><th>what it asked</th><th>measured</th><th>expected</th></tr></thead><tbody>${
      failures.map((item) => `<tr><td class="mono">${esc(item.tier)}</td><td>${esc(item.message)}</td><td class="mono bad">${esc(item.measured)}</td><td class="mono">${esc(item.expected)}</td></tr>`).join("")
    }</tbody></table>`;

  const parityRows = (entry.parityTable ?? []).length === 0
    ? ""
    : `<details class="parity"><summary>Every paragraph and cell, preview against Word</summary>
       <table><thead><tr><th>what</th><th>page</th><th>Δx</th><th>Δy</th></tr></thead><tbody>${
      entry.parityTable.map((row) => {
        const drift = Math.max(Math.abs(row.dxMm ?? 0), Math.abs(row.dyMm ?? 0));
        const cls = drift > 1 ? "bad" : drift > 0.5 ? "warn" : "";
        // A cell and a paragraph are measured the same way and mean different
        // things, so the survey says which each row is.
        const kind = row.cell === true ? "cell" : "para";
        return `<tr><td><span class="kind">${kind}</span> ${esc(row.text)}</td><td class="mono">${row.previewPage ?? "—"} / ${row.wordPage ?? "—"}</td><td class="mono ${cls}">${fmt(row.dxMm)}</td><td class="mono ${cls}">${fmt(row.dyMm)}</td></tr>`;
      }).join("")
    }</tbody></table></details>`;

  const shot = (uri, caption, missing) =>
    uri === null
      ? `<div class="shot empty">${caption}<br><small>${missing}</small></div>`
      : `<figure class="shot"><img src="${uri}" alt="${esc(caption)}"><figcaption>${caption}</figcaption></figure>`;

  // A page one engine has and the other does not is a finding, not a gap in
  // the report — docx-preview breaks only where the file says to, so a
  // document Word runs onto three sheets is drawn as one long one. Saying "no
  // such page" beats "not taken", which reads as the harness having failed.
  const pages = Math.max(previewShots.length, wordShots.length, 1);
  const stages = Array.from({ length: pages }, (_, index) => `
    <div class="stage">
      ${shot(previewShots[index] ?? null, `Preview · page ${index + 1}`,
    index < previewShots.length ? "not captured" : "the preview has no such page")}
      ${shot(wordShots[index] ?? null, `Word · page ${index + 1}`,
    index < wordShots.length ? "not captured" : "Word has no such page")}
    </div>`).join("");

  return `
<section class="case" id="${esc(entry.id)}">
  <header>
    <div class="hrow">
      <span class="status ${String(entry.status).toLowerCase()}">${esc(entry.status)}</span>
      <h2>${esc(entry.title ?? entry.id)}</h2>
    </div>
    <p class="meta"><code>${esc(entry.id)}</code> · claims <b>${esc(entry.claim ?? "?")}</b>${entry.word ? ` · ${esc(entry.word)}` : ""}</p>
    <div class="tiers">${tiers}</div>
  </header>

  <div class="triptych">
    <div class="col code">
      <p class="collabel">The code</p>
      <pre>${esc(documentOf(source))}</pre>
    </div>
    <div class="col shots">
      <p class="collabel">What the preview shows, and what Word makes of it</p>
      ${stages}
    </div>
  </div>

  ${failureRows}
  ${parityRows}
</section>`;
}

function label(key) {
  return { ooxml: "A", preview: "B", word: "C", parity: "X" }[key] ?? "?";
}

function fmt(value) {
  return value === null || value === undefined ? "—" : `${value > 0 ? "+" : ""}${value.toFixed(2)}mm`;
}

function esc(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function page(body, tally, generatedAt) {
  return `<title>Paragraph Conformance</title>
<style>
  @import url("https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;700&family=Source+Serif+4:opsz,wght@8..60,400;8..60,600&family=IBM+Plex+Mono:wght@400;500&display=swap");

  :root {
    --paper: #f4f6f9; --surface: #fff; --sunk: #eceff5;
    --ink: #14181f; --soft: #3d4557; --muted: #666f84;
    --rule: #dce0ea; --firm: #c3cad9; --accent: #3b3fbf;
    --pass: #2b6c51; --pass-bg: #e2f0e9;
    --fail: #a8352c; --fail-bg: #f7e4e2;
    --known: #8a5f0f; --known-bg: #f6ecd6;
    --skip-ink: #6d7488; --skip-bg: #e8eaf0;
    --font-display: "Archivo", system-ui, sans-serif;
    --font-body: "Source Serif 4", Georgia, serif;
    --font-mono: "IBM Plex Mono", Consolas, monospace;
  }
  @media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) {
    --paper: #0e1118; --surface: #161a23; --sunk: #1c212c;
    --ink: #e6e9f0; --soft: #b9c0cf; --muted: #8f97ab;
    --rule: #262c39; --firm: #38404f; --accent: #9195f7;
    --pass: #64c294; --pass-bg: #16302a;
    --fail: #e08379; --fail-bg: #331a17;
    --known: #dcaa52; --known-bg: #33290f;
    --skip-ink: #838b9e; --skip-bg: #1e232e;
  } }
  :root[data-theme="dark"] {
    --paper: #0e1118; --surface: #161a23; --sunk: #1c212c;
    --ink: #e6e9f0; --soft: #b9c0cf; --muted: #8f97ab;
    --rule: #262c39; --firm: #38404f; --accent: #9195f7;
    --pass: #64c294; --pass-bg: #16302a;
    --fail: #e08379; --fail-bg: #331a17;
    --known: #dcaa52; --known-bg: #33290f;
    --skip-ink: #838b9e; --skip-bg: #1e232e;
  }

  * { box-sizing: border-box; }
  body { margin: 0; background: var(--paper); color: var(--ink);
         font-family: var(--font-body); font-size: 16px; line-height: 1.55; }
  .sheet { max-width: 84rem; margin: 0 auto; padding: 3rem 1.5rem 6rem; }

  h1 { font-family: var(--font-display); font-size: clamp(2rem, 5vw, 3rem);
       letter-spacing: -0.025em; margin: 0 0 .6rem; line-height: 1.05; }
  .eyebrow { font-family: var(--font-display); font-size: .7rem; font-weight: 600;
             letter-spacing: .16em; text-transform: uppercase; color: var(--accent); margin: 0 0 .9rem; }
  .lede { max-width: 60ch; color: var(--soft); margin: 0; }
  .tally { display: flex; flex-wrap: wrap; gap: .5rem; margin: 2rem 0 0;
           padding-top: 1.4rem; border-top: 2px solid var(--ink); }

  h2 { font-family: var(--font-display); font-size: 1.3rem; letter-spacing: -.015em;
       margin: 0; line-height: 1.2; text-wrap: balance; }
  .case { background: var(--surface); border: 1px solid var(--rule);
          margin-top: 2rem; padding: 1.5rem; }
  .hrow { display: flex; align-items: baseline; gap: .8rem; flex-wrap: wrap; }
  .meta { font-size: .85rem; color: var(--muted); margin: .5rem 0 0; }
  .tiers { display: flex; flex-wrap: wrap; gap: .4rem; margin-top: .9rem; }

  .status, .tier { font-family: var(--font-display); font-size: .65rem; font-weight: 600;
                   letter-spacing: .09em; text-transform: uppercase; padding: .22rem .55rem; border-radius: 2px; }
  .status.pass, .tier.pass { background: var(--pass-bg); color: var(--pass); }
  .status.fail, .tier.fail { background: var(--fail-bg); color: var(--fail); }
  .status.known, .tier.known { background: var(--known-bg); color: var(--known); }
  .status.stale { background: var(--known-bg); color: var(--known); }
  .status.skipped, .tier.skipped { background: var(--skip-bg); color: var(--skip-ink); }
  .tier b { font-family: var(--font-mono); margin-right: .3em; }

  .triptych { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1.15fr);
              gap: 1.5rem; margin-top: 1.5rem; }
  @media (max-width: 900px) { .triptych { grid-template-columns: 1fr; } }
  .collabel { font-family: var(--font-display); font-size: .66rem; font-weight: 600;
              letter-spacing: .13em; text-transform: uppercase; color: var(--muted); margin: 0 0 .6rem; }
  pre { font-family: var(--font-mono); font-size: .72rem; line-height: 1.65; margin: 0;
        background: var(--sunk); border-left: 3px solid var(--accent);
        padding: 1rem; overflow-x: auto; color: var(--soft); }

  .stage { display: grid; grid-template-columns: 1fr 1fr; gap: .8rem; margin-bottom: 1rem; }
  .shot { margin: 0; }
  .shot img { width: 100%; height: auto; display: block; border: 1px solid var(--firm); background: #fff; }
  .shot figcaption { font-family: var(--font-display); font-size: .64rem; letter-spacing: .08em;
                     text-transform: uppercase; color: var(--muted); margin-top: .4rem; }
  .shot.empty { font-family: var(--font-display); font-size: .68rem; color: var(--muted);
                border: 1px dashed var(--firm); padding: 2rem 1rem; text-align: center;
                letter-spacing: .06em; text-transform: uppercase; line-height: 2; }
  .shot.empty small { text-transform: none; letter-spacing: 0; font-size: .95em; opacity: .8; }
  .kind { font: 600 .72em ui-monospace, monospace; letter-spacing: .08em; text-transform: uppercase;
          color: var(--muted); margin-right: .5em; }

  table { border-collapse: collapse; width: 100%; font-size: .82rem; margin-top: 1.2rem; }
  th { font-family: var(--font-display); font-size: .62rem; font-weight: 600; letter-spacing: .11em;
       text-transform: uppercase; color: var(--muted); text-align: left;
       padding: 0 .8rem .5rem 0; border-bottom: 1px solid var(--firm); }
  td { padding: .45rem .8rem .45rem 0; border-bottom: 1px solid var(--rule); vertical-align: top; }
  .mono { font-family: var(--font-mono); font-size: .76rem; font-variant-numeric: tabular-nums; }
  .bad { color: var(--fail); }
  .warn { color: var(--known); }
  .fails { background: transparent; }
  details.parity { margin-top: 1.2rem; }
  summary { font-family: var(--font-display); font-size: .7rem; letter-spacing: .1em;
            text-transform: uppercase; color: var(--muted); cursor: pointer; }
  summary:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
  .colophon { margin-top: 3rem; padding-top: 1rem; border-top: 2px solid var(--ink);
              font-family: var(--font-display); font-size: .72rem; color: var(--muted); }
</style>

<div class="sheet">
  <p class="eyebrow">Docxcelerate · conformance</p>
  <h1>Paragraph Conformance</h1>
  <p class="lede">One small document per paragraph feature, packed to a <code>.docx</code>, read back
    by docx-preview, and opened in Word 16. Three readings of the same bytes, and the fourth tier
    holds the first two against each other.</p>
  <div class="tally">
    ${Object.entries(tally).map(([status, count]) =>
    `<span class="status ${status.toLowerCase()}">${count} ${status}</span>`
  ).join("")}
  </div>

  ${body}

  <p class="colophon">Generated ${esc(generatedAt)} · Word 16.0 · Chrome headless at 96dpi ·
    geometry in millimetres from the top-left of the text column</p>
</div>
`;
}
