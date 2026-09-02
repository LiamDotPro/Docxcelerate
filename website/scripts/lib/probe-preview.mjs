/**
 * Probe B of the invoice verification harness: what the baked preview
 * actually lays out.
 *
 * The preview page (`.verify/invoice.html`) is docx-preview's reading of the
 * packed file, already baked by verify-build.mjs. This probe opens it in a
 * real layout engine — headless Chrome — because jsdom does not do layout,
 * and every number the objectives care about (region rects, row heights,
 * section overflow) only exists once something has laid the page out.
 *
 * Mechanics, per VERIFY-CONTRACT.md "Environment facts": a tiny same-origin
 * static server (a file:// parent is blocked from reading an http:// iframe,
 * which bit the design-side harness already), a harness page that iframes the
 * preview and prints base64 JSON into a `<pre>`, and `--dump-dom` to read it
 * back. Base64 because `--dump-dom` re-entity-encodes text nodes, and
 * un-escaping HTML by hand is how measurement bugs are born.
 *
 * A region that cannot be located is recorded as `null`, never thrown on:
 * the first board is expected to be mostly FAIL, and "not found" is a
 * measurement, not an error.
 */
import { execFile } from "node:child_process";
import { createServer } from "node:http";
import { copyFile, mkdtemp, readFile, rm, unlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";
import { setTimeout as sleep } from "node:timers/promises";
import pngjs from "pngjs";

const { PNG } = pngjs;
const execFileAsync = promisify(execFile);

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..", "..");
const VERIFY_DIR = resolve(ROOT, ".verify");
const CHROME = "C:/Program Files/Google/Chrome/Application/chrome.exe";
const PORT = 8901;

/**
 * One width for both Chrome runs. The measuring run reads rects inside a
 * 1000px iframe; the screenshot run loads the preview directly in a 1000px
 * window. Same width, same `margin: 0 auto` centring — so the document
 * coordinates measured in the first run are valid crop rects in the second.
 */
const VIEWPORT_WIDTH = 1000;

/**
 * The anchored regions of the contract's region table, with the element kinds
 * worth measuring. "Smallest element containing the anchor" alone would
 * always return a bare span, whose rect is a text extent rather than the
 * region the design fixture measured — so each region names the containers
 * it means, in preference order (first selector with any match wins, then
 * smallest by area). The geometric regions (rule, p2-rule, charges-rule) and
 * the table-derived ones (charges-head/-body, footer, footer-2) are located
 * in the harness itself, not from this list.
 */
const ANCHOR_REGIONS = [
  ["letterhead", ["table"], "Software consultancy · Manchester"],
  ["band", ["table"], "Issue date"],
  ["status-pill", ["span"], "Awaiting payment"],
  ["parties", ["table"], "Billed to"],
  ["summary-label", ["span"], "Engagement summary"],
  ["summary", ["p"], "Sprint 14 closed out"],
  ["totals-panel", ["table"], "Subtotal"],
  ["total-bar", ["tr", "td"], "Total due"],
  ["closer", ["p"], "are on page 2"],
  // "Awaiting payment" is lowercase in the DOM (uppercased by CSS), so the
  // literal "PAYMENT" only matches the page-2 wordmark once D8 lands.
  ["p2-letterhead", ["table", "p"], "PAYMENT"],
  ["bank-grid", ["table"], "Sort code"],
  ["reference-panel", ["table", "td", "p"], "Payment reference"],
  // The design's `terms` region is the whole terms block; a cell (after the
  // D5 two-column layout lands) is the honest container, a paragraph the
  // honest fallback until then.
  ["terms", ["td", "p"], "Payment within 14 days"],
  ["scan-card", ["table", "td", "p"], "Scan to pay"],
];

/** Report order: the contract's region table, top to bottom. */
const REGION_ORDER = [
  "letterhead", "rule", "band", "status-pill", "parties", "summary-label",
  "summary", "charges-head", "charges-body", "charges-rule", "totals-panel",
  "total-bar", "closer", "footer", "p2-letterhead", "p2-rule", "bank-grid",
  "reference-panel", "terms", "scan-card", "footer-2",
];

/**
 * The measuring page. It iframes the preview same-origin, waits for layout,
 * and prints everything Probe B records as one base64 blob. Written in
 * old-fashioned browser JS on purpose: no template-literal characters, so it
 * can live inside this file's template literal without escape gymnastics.
 */
function harnessHtml() {
  return `<!doctype html>
<html>
<head><meta charset="utf-8"><title>probe B</title></head>
<body style="margin:0">
<pre id="out">pending</pre>
<iframe id="f" src="./invoice.html" style="display:block;border:0;width:${VIEWPORT_WIDTH}px;height:4000px"></iframe>
<script>
(function () {
  "use strict";
  var NAVY = "rgb(44, 61, 143)";
  var HAIRLINE = "rgb(217, 221, 235)";
  var ANCHOR_REGIONS = ${JSON.stringify(ANCHOR_REGIONS)};

  function round2(n) { return Math.round(n * 100) / 100; }
  function round1(n) { return Math.round(n * 10) / 10; }

  function emit(obj) {
    var json = JSON.stringify(obj);
    document.getElementById("out").textContent =
      "@@MEASURE-B@@" + btoa(unescape(encodeURIComponent(json))) + "@@END@@";
  }

  // Two timeouts: one for the iframe to finish its first layout, a second
  // after the frame is grown past its content so nothing scrolls or clips.
  // Chrome's --virtual-time-budget fast-forwards both, so the wall-clock
  // cost is nil.
  window.addEventListener("load", function () {
    setTimeout(function () {
      try {
        var frame = document.getElementById("f");
        var doc = frame.contentDocument;
        frame.style.height = (doc.documentElement.scrollHeight + 100) + "px";
        setTimeout(function () {
          try { emit(collect(doc)); }
          catch (error) { emit({ error: String(error && error.stack || error) }); }
        }, 200);
      } catch (error) { emit({ error: String(error) }); }
    }, 300);
  });

  function collect(doc) {
    var view = doc.defaultView;
    var sections = [].slice.call(doc.querySelectorAll("section.docx"));

    // Content box of each page: docx-preview writes the page margins as
    // section padding, so the content box is the printable area — the frame
    // every region rect is recorded against, and the crop rect for the
    // per-page screenshots.
    var boxes = sections.map(function (s) {
      var r = s.getBoundingClientRect();
      var cs = view.getComputedStyle(s);
      var pad = {
        left: parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth),
        top: parseFloat(cs.paddingTop) + parseFloat(cs.borderTopWidth),
        right: parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth),
        bottom: parseFloat(cs.paddingBottom) + parseFloat(cs.borderBottomWidth),
      };
      return {
        el: s, rect: r, pad: pad,
        content: {
          x: r.x + pad.left, y: r.y + pad.top,
          w: r.width - pad.left - pad.right, h: r.height - pad.top - pad.bottom,
        },
      };
    });

    function sectionIndexOf(el) {
      return sections.indexOf(el.closest("section.docx"));
    }

    function rel(el) {
      if (!el) return null;
      var i = sectionIndexOf(el);
      if (i < 0) return null;
      var r = el.getBoundingClientRect();
      var c = boxes[i].content;
      return {
        sectionIndex: i,
        x: round2(r.x - c.x), y: round2(r.y - c.y),
        w: round2(r.width), h: round2(r.height),
      };
    }

    function unionOf(els) {
      if (!els.length) return null;
      var i = sectionIndexOf(els[0]);
      if (i < 0) return null;
      var x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
      els.forEach(function (el) {
        var r = el.getBoundingClientRect();
        x1 = Math.min(x1, r.x); y1 = Math.min(y1, r.y);
        x2 = Math.max(x2, r.x + r.width); y2 = Math.max(y2, r.y + r.height);
      });
      var c = boxes[i].content;
      return {
        sectionIndex: i,
        x: round2(x1 - c.x), y: round2(y1 - c.y),
        w: round2(x2 - x1), h: round2(y2 - y1),
      };
    }

    // How tall a running strip actually draws.
    //
    // Not the container's height: docx-preview gives a header and a footer a
    // fixed reserve (a min-height cancelled by an equal negative margin) taken
    // from the section's header distance, so the box is the same size on a page
    // that shows nothing as on one that shows a bar. What a reader sees is the
    // content inside it, so that is what is measured.
    function drawnHeight(el) {
      if (!el) return null;
      var top = Infinity, bottom = -Infinity;
      for (var i = 0; i < el.children.length; i++) {
        var r = el.children[i].getBoundingClientRect();
        if (r.height <= 0) continue;
        if (r.top < top) top = r.top;
        if (r.bottom > bottom) bottom = r.bottom;
      }
      return bottom > top ? round2(bottom - top) : 0;
    }

    function smallest(els) {
      return els.slice().sort(function (a, b) {
        var ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
        return ra.width * ra.height - rb.width * rb.height;
      })[0] || null;
    }

    // What an element reads as on the page, which is not always what its text
    // node says: a word set in capitals by the theme is still lower case in
    // the DOM, and the anchors name what a reader sees.
    function renderedText(el) {
      var text = el.textContent;
      try {
        var transform = el.ownerDocument.defaultView.getComputedStyle(el).textTransform;
        if (transform === "uppercase") return text.toUpperCase();
        if (transform === "lowercase") return text.toLowerCase();
      } catch (error) { /* no view to compute against; the text stands. */ }
      return text;
    }

    function byText(roots, selectors, text) {
      for (var s = 0; s < selectors.length; s++) {
        var hits = [];
        roots.forEach(function (root) {
          var els = root.querySelectorAll(selectors[s]);
          for (var e = 0; e < els.length; e++) {
            // Either form counts: an anchor may be written as the theme sets
            // it (PAYMENT, in capitals) or as the document spells it
            // ("Awaiting payment"), and both name the same element.
            if (
              els[e].textContent.indexOf(text) !== -1 ||
              renderedText(els[e]).indexOf(text) !== -1
            ) hits.push(els[e]);
          }
        });
        if (hits.length) return smallest(hits);
      }
      return null;
    }

    // Body regions live in the articles; searching the sections whole would
    // also match the running header, which repeats the sender's name.
    var articles = boxes
      .map(function (b) { return b.el.querySelector(":scope > article"); })
      .filter(Boolean);

    var regions = {};
    ANCHOR_REGIONS.forEach(function (entry) {
      regions[entry[0]] = rel(byText(articles, entry[1], entry[2]));
    });

    // The charges table: the one whose header row says "Description". Rows
    // after it are the charge rows — F7 reads their heights one by one.
    var chargesTable = null;
    for (var a = 0; a < articles.length && !chargesTable; a++) {
      var tables = articles[a].querySelectorAll("table");
      for (var t = 0; t < tables.length; t++) {
        var rows = tables[t].rows;
        if (rows.length && rows[0].textContent.indexOf("Description") !== -1) {
          chargesTable = tables[t];
          break;
        }
      }
    }
    var chargeRows = chargesTable ? [].slice.call(chargesTable.rows, 1) : [];
    regions["charges-head"] = chargesTable ? rel(chargesTable.rows[0]) : null;
    regions["charges-body"] = unionOf(chargeRows);
    var chargeRowHeights = chargeRows.map(function (row) {
      return round2(row.getBoundingClientRect().height);
    });
    var chargesBottom = null;
    if (chargeRows.length) {
      var lastRect = chargeRows[chargeRows.length - 1].getBoundingClientRect();
      chargesBottom = lastRect.y + lastRect.height;
    }

    // Colour-strip regions are located geometrically, not by anchor: a rule
    // has no text to anchor on. Full-width, the exact serialised colour
    // (docx-preview always says rgb(...), never hex), and no taller than 5px
    // — today's rules are 17.4pt-tall paragraphs, so they correctly fail
    // this test and come back null.
    function stripsIn(i) {
      var out = { navy: [], hairline: [] };
      var article = boxes[i] && boxes[i].el.querySelector(":scope > article");
      if (!article) return out;
      var all = article.querySelectorAll("*");
      for (var e = 0; e < all.length; e++) {
        var bg = view.getComputedStyle(all[e]).backgroundColor;
        if (bg !== NAVY && bg !== HAIRLINE) continue;
        var r = all[e].getBoundingClientRect();
        if (r.height <= 0 || r.height > 5) continue;
        if (r.width < boxes[i].content.w - 2) continue;
        (bg === NAVY ? out.navy : out.hairline).push(all[e]);
      }
      var byY = function (a, b) {
        return a.getBoundingClientRect().y - b.getBoundingClientRect().y;
      };
      out.navy.sort(byY);
      out.hairline.sort(byY);
      return out;
    }
    var strips0 = stripsIn(0);
    var strips1 = stripsIn(1);
    regions["rule"] = rel(strips0.navy[0] || null);
    regions["p2-rule"] = rel(strips1.navy[0] || null);
    // The charges-rule closes the charges table, so of the hairline strips
    // on page one it is the first at or below the last charge row.
    var chargesRule = null;
    if (strips0.hairline.length) {
      if (chargesBottom === null) {
        chargesRule = strips0.hairline[0];
      } else {
        for (var h = 0; h < strips0.hairline.length; h++) {
          if (strips0.hairline[h].getBoundingClientRect().y >= chargesBottom - 2) {
            chargesRule = strips0.hairline[h];
            break;
          }
        }
      }
    }
    regions["charges-rule"] = rel(chargesRule);

    // The page footers. The region worth comparing against the design's
    // footer bar is the table docx-preview puts inside the footer container
    // — the container itself carries a negative margin and a min-height that
    // are docx-preview furniture, not the bar.
    function footerRegion(i, anchor) {
      var footer = boxes[i] && boxes[i].el.querySelector(":scope > footer");
      if (!footer) return null;
      if (anchor && footer.textContent.indexOf(anchor) === -1) return null;
      return rel(footer.querySelector("table") || footer);
    }
    regions["footer"] = footerRegion(0, "Registered in England");
    regions["footer-2"] = footerRegion(1, null);

    // Per-section furniture: header/footer element heights (F2 compares
    // them across pages) and the page-number text — the element whose text
    // is digits around a slash. docx-preview does not evaluate PAGE fields,
    // so until C6 lands this reads a bare "/", which is itself the finding.
    var sectionsOut = boxes.map(function (b) {
      var header = b.el.querySelector(":scope > header");
      var footer = b.el.querySelector(":scope > footer");
      var pageNumber = null;
      var pageNumberRightPx = null;
      if (footer) {
        var candidates = [];
        var all = footer.querySelectorAll("*");
        for (var e = 0; e < all.length; e++) {
          var text = all[e].textContent.replace(/\\u00a0/g, " ").trim();
          if (text.indexOf("/") === -1) continue;
          if (!/^\\d*\\s*\\/\\s*\\d*$/.test(text)) continue;
          candidates.push(all[e]);
        }
        var el = smallest(candidates);
        if (el) {
          pageNumber = {
            text: el.textContent.replace(/\\u00a0/g, " ").trim(),
            rect: rel(el),
          };
          // F5 compares against a design inset measured from the paper's
          // edge, not from the text column, so this one is page-relative
          // while the rect above stays content-relative like every region.
          pageNumberRightPx = round2(el.getBoundingClientRect().right - b.rect.x);
        }
      }

      // F12: the page should end with its content, not with an empty line
      // left over from however the break was carried.
      var trailingEmptyPara = null;
      var article = b.el.querySelector(":scope > article") || b.el;
      var blocks = article.children;
      for (var k = blocks.length - 1; k >= 0; k--) {
        var last = blocks[k];
        if (last.tagName === "HEADER" || last.tagName === "FOOTER") continue;
        trailingEmptyPara = last.textContent.trim() === "" &&
          !last.querySelector("img, svg, table");
        break;
      }
      return {
        w: round2(b.rect.width),
        h: round2(b.rect.height),
        // The sheet's own rect in the screenshot. The design fixtures and the
        // Word raster are both whole pages, so a preview crop has to be one
        // too — cropped to the content box instead, every page-relative rect
        // laid onto it lands short by the margin.
        box: {
          x: round2(b.rect.x), y: round2(b.rect.y),
          w: round2(b.rect.width), h: round2(b.rect.height),
        },
        contentBox: {
          x: round2(b.content.x), y: round2(b.content.y),
          w: round2(b.content.w), h: round2(b.content.h),
        },
        padding: {
          left: round2(b.pad.left), top: round2(b.pad.top),
          right: round2(b.pad.right), bottom: round2(b.pad.bottom),
        },
        headerHeight: drawnHeight(header),
        footerHeight: drawnHeight(footer),
        // What the strip is made of, when its height and Word's disagree —
        // a bar is one line of small print, so a cell taller than that is the
        // answer to where the extra depth came from.
        footerCells: footer
          ? [].slice.call(footer.querySelectorAll("td")).map(function (td) {
            var cs = view.getComputedStyle(td);
            return {
              tag: td.tagName,
              h: round2(td.getBoundingClientRect().height),
              lineHeight: cs.lineHeight,
              padding: cs.paddingTop,
              text: td.textContent.replace(/\\s+/g, " ").trim().slice(0, 24),
            };
          })
          : null,
        pageNumber: pageNumber,
        pageNumberRightPx: pageNumberRightPx,
        trailingEmptyPara: trailingEmptyPara,
      };
    });

    // The faces the money columns actually resolved to (F6), read from
    // computed style so a face the browser cannot resolve still reports
    // what was asked for.
    var fonts = { amountCell: null, descriptionCell: null };
    if (chargeRows.length && chargeRows[0].cells.length) {
      var cells = chargeRows[0].cells;
      var descSpan = cells[0].querySelector("span");
      var amountSpan = cells[cells.length - 1].querySelector("span");
      if (descSpan) fonts.descriptionCell = view.getComputedStyle(descSpan).fontFamily;
      if (amountSpan) fonts.amountCell = view.getComputedStyle(amountSpan).fontFamily;
    }

    // P0's Chrome half: Aptos resolved means "Handgloves 1234" at 40px
    // measures ~297.2; an unresolved face falls back to the monospace width,
    // ~329.9, and the two being equal is exactly the failure P0 exists for.
    var ctx = document.createElement("canvas").getContext("2d");
    ctx.font = "40px Aptos, monospace";
    var aptosWidth = ctx.measureText("Handgloves 1234").width;
    ctx.font = "40px monospace";
    var monoWidth = ctx.measureText("Handgloves 1234").width;

    return {
      sections: sectionsOut,
      regions: regions,
      chargeRowHeights: chargeRowHeights,
      fonts: fonts,
      fontProbe: { aptosWidth: round1(aptosWidth), monoWidth: round1(monoWidth) },
      documentHeight: Math.ceil(doc.documentElement.scrollHeight),
    };
  }
})();
</script>
</body>
</html>
`;
}

/** Runs Chrome once with the flags verified to work here, and returns stdout. */
async function runChrome(args) {
  const { stdout } = await execFileAsync(CHROME, [
    "--headless=new",
    "--disable-gpu",
    "--hide-scrollbars",
    "--force-device-scale-factor=1",
    "--virtual-time-budget=15000",
    ...args,
  ], { maxBuffer: 64 * 1024 * 1024, timeout: 120000, killSignal: "SIGKILL" });
  return stdout;
}

/**
 * A static server for the harness and the preview it iframes. Same-origin is
 * the whole reason it exists. The port is shared with the other probes by
 * contract, so a busy port gets a few retries before it is a failure.
 */
async function startServer(root) {
  const types = { ".html": "text/html; charset=utf-8", ".png": "image/png" };

  for (let attempt = 1; ; attempt++) {
    try {
      return await new Promise((resolvePromise, rejectPromise) => {
        const server = createServer(async (request, response) => {
          try {
            const pathname = decodeURIComponent(new URL(request.url, "http://localhost").pathname);
            const file = join(root, pathname.replaceAll("..", ""));
            const body = await readFile(file);
            response.writeHead(200, { "Content-Type": types[extname(file)] ?? "application/octet-stream" });
            response.end(body);
          } catch {
            response.writeHead(404);
            response.end("not found");
          }
        });
        server.once("error", rejectPromise);
        server.listen(PORT, "127.0.0.1", () => resolvePromise(server));
      });
    } catch (error) {
      if (error?.code !== "EADDRINUSE" || attempt >= 4) {
        throw new Error(`static server could not bind 127.0.0.1:${PORT}: ${error?.message ?? error}`);
      }
      await sleep(1500);
    }
  }
}

/** The harness's base64 payload, dug out of the dumped DOM. */
function decodeMeasurement(dom) {
  const match = /@@MEASURE-B@@([A-Za-z0-9+/=]+)@@END@@/.exec(dom);
  if (!match) {
    throw new Error("harness produced no measurement (marker not found in dumped DOM)");
  }
  return JSON.parse(Buffer.from(match[1], "base64").toString("utf8"));
}

/**
 * Crops each page out of the full-page screenshot.
 *
 * The whole sheet, not its content box: probe V lays the design fixture's
 * rects onto these pixels, and those rects are page-relative — as is the Word
 * raster it compares them against. A crop that started at the text column
 * would put every one of them a margin out, which is why the cross-engine
 * numbers used to move for reasons that had nothing to do with the document.
 */
async function cropSections(fullPngPath, sections) {
  const full = PNG.sync.read(await readFile(fullPngPath));
  const written = [];

  for (let index = 0; index < sections.length && index < 8; index++) {
    const box = sections[index].box ?? sections[index].contentBox;
    const x = Math.max(0, Math.round(box.x));
    const y = Math.max(0, Math.round(box.y));
    const width = Math.min(Math.round(box.w), full.width - x);
    const height = Math.min(Math.round(box.h), full.height - y);
    if (width <= 0 || height <= 0) continue;

    const crop = new PNG({ width, height });
    PNG.bitblt(full, crop, x, y, width, height, 0, 0);
    const file = `preview-p${index + 1}.png`;
    await writeFile(resolve(VERIFY_DIR, file), PNG.sync.write(crop));
    written.push({ file, w: width, h: height });
  }

  // Stale pages from a run that had more sections would otherwise survive
  // and feed probe V a page that no longer exists.
  for (let index = written.length; index < 8; index++) {
    await unlink(resolve(VERIFY_DIR, `preview-p${index + 1}.png`)).catch(() => {});
  }

  return written;
}

/**
 * Measures the baked preview and writes `.verify/measure-b.json`.
 *
 * @returns The written measurement, for a runner that wants it in-process.
 */
export async function runPreviewProbe() {
  const output = {
    probe: "B",
    generatedAt: new Date().toISOString(),
    source: "invoice.html",
    viewportWidth: VIEWPORT_WIDTH,
    pxPerMm: 3.77953,
    error: null,
    sections: [],
    regions: Object.fromEntries(REGION_ORDER.map((name) => [name, null])),
    chargeRowHeights: [],
    fonts: { amountCell: null, descriptionCell: null },
    fontProbe: null,
    screenshots: [],
  };

  let server;
  let tempDir;

  try {
    tempDir = await mkdtemp(join(tmpdir(), "probe-b-"));
    await copyFile(resolve(VERIFY_DIR, "invoice.html"), join(tempDir, "invoice.html"));
    await writeFile(join(tempDir, "__probe-b.html"), harnessHtml(), "utf8");

    server = await startServer(tempDir);

    const dom = await runChrome([
      `--window-size=${VIEWPORT_WIDTH},900`,
      "--dump-dom",
      `http://127.0.0.1:${PORT}/__probe-b.html`,
    ]);
    const measured = decodeMeasurement(dom);
    if (measured.error) {
      throw new Error(`harness error: ${measured.error}`);
    }

    output.sections = measured.sections;
    // Regions come back in whatever order the harness found them; the report
    // keeps the contract's table order so a human can read it top to bottom.
    for (const name of REGION_ORDER) {
      output.regions[name] = measured.regions[name] ?? null;
    }
    output.chargeRowHeights = measured.chargeRowHeights;
    output.fonts = measured.fonts;
    output.fontProbe = measured.fontProbe;
    output.documentHeight = measured.documentHeight;

    // Second Chrome run: the preview straight (no harness around it), in a
    // window tall enough that "the viewport" and "the page" are the same
    // thing — new headless screenshots clip to the window.
    if (measured.sections.length > 0) {
      const height = Math.max(900, Math.min(measured.documentHeight + 20, 6000));
      const fullPngPath = join(tempDir, "full.png");
      await runChrome([
        `--window-size=${VIEWPORT_WIDTH},${height}`,
        `--screenshot=${fullPngPath}`,
        `http://127.0.0.1:${PORT}/invoice.html`,
      ]);
      output.screenshots = await cropSections(fullPngPath, measured.sections);
    }
  } catch (error) {
    // The probe never throws its measurements away: a crash is recorded in
    // the JSON the objectives read, so "probe missing" is a status on the
    // board rather than a stack trace in a log.
    output.error = String(error?.message ?? error);
  } finally {
    if (server) {
      server.closeAllConnections?.();
      await new Promise((resolvePromise) => server.close(resolvePromise));
    }
    if (tempDir) {
      await rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }

  await writeFile(
    resolve(VERIFY_DIR, "measure-b.json"),
    JSON.stringify(output, null, 2),
    "utf8",
  );

  return output;
}

// CLI: node scripts/lib/probe-preview.mjs — prints a human summary of what
// was (and was not) found, so a red region is one glance away from its rect.
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  runPreviewProbe().then((output) => {
    if (output.error) {
      console.error("probe B failed:", output.error);
      process.exitCode = 1;
    }
    output.sections.forEach((section, index) => {
      console.log(
        `section ${index + 1}: ${section.w}x${section.h}px` +
        ` header=${section.headerHeight}px footer=${section.footerHeight}px` +
        ` pageNumber=${section.pageNumber ? JSON.stringify(section.pageNumber.text) : "null"}`,
      );
    });
    for (const name of REGION_ORDER) {
      const region = output.regions[name];
      console.log(region
        ? `region ${name}: section=${region.sectionIndex} x=${region.x} y=${region.y} w=${region.w} h=${region.h}`
        : `region ${name}: not found`);
    }
    console.log("chargeRowHeights:", JSON.stringify(output.chargeRowHeights));
    console.log("fonts:", JSON.stringify(output.fonts));
    console.log("fontProbe:", JSON.stringify(output.fontProbe));
    console.log("screenshots:", JSON.stringify(output.screenshots));
  }, (error) => {
    console.error("probe B crashed:", error);
    process.exitCode = 1;
  });
}
