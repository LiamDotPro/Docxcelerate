/**
 * Probe B — what the preview actually lays out.
 *
 * The preview page is docx-preview's reading of the packed file, already baked
 * by the build step. This probe opens it in a real layout engine, because
 * jsdom does not lay anything out and every number worth asserting on — where
 * a line starts, how tall a paragraph draws, the gap under it — only exists
 * once something has laid the page out.
 *
 * Two things it does that a naive version would not:
 *
 * It measures **drawn text**, not element boxes. A centred paragraph's `<p>` is
 * the full width of the column whether it is centred or not; only the line
 * rectangles inside it move. Reading the box would report every alignment as
 * identical and every one of those reports would be wrong.
 *
 * It checks the **font resolved** before believing anything else. A browser
 * handed a face it does not have falls back silently, and every measurement
 * taken against the fallback is wrong by a plausible amount. That is a
 * precondition, not an assertion.
 *
 * @module
 */

import { execFile } from "node:child_process";
import { createServer } from "node:http";
import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { setTimeout as sleep } from "node:timers/promises";
import pngjs from "pngjs";
import { PX_PER_MM } from "./case.mjs";

const { PNG } = pngjs;
const execFileAsync = promisify(execFile);

const HERE = dirname(fileURLToPath(import.meta.url));

const PORT = 8902;

/**
 * Where Chrome is. The environment variable is the escape hatch for a machine
 * that keeps it somewhere else; the two defaults are where Windows puts it.
 */
const CHROME_CANDIDATES = [
  process.env.CHROME_PATH,
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
].filter(Boolean);

/** One width for both Chrome runs, so measured coordinates are valid crop rects. */
const VIEWPORT_WIDTH = 1000;

export function chromePath() {
  return CHROME_CANDIDATES.find((candidate) => existsSync(candidate)) ?? null;
}

/**
 * The measuring page.
 *
 * Written in old-fashioned browser JS on purpose: no template literals, so it
 * survives living inside one. It iframes the preview same-origin — a `file://`
 * parent cannot read an `http://` child — waits for layout, and prints its
 * findings as one base64 blob, because `--dump-dom` re-entity-encodes text and
 * unescaping HTML by hand is how measurement bugs are born.
 */
function harnessHtml(bodyFont) {
  return `<!doctype html>
<html>
<head><meta charset="utf-8"><title>probe B</title></head>
<body style="margin:0">
<pre id="out">pending</pre>
<iframe id="f" src="./case.html" style="display:block;border:0;width:${VIEWPORT_WIDTH}px;height:6000px"></iframe>
<script>
(function () {
  "use strict";
  var BODY_FONT = ${JSON.stringify(bodyFont ?? "Aptos")};

  function r2(n) { return Math.round(n * 100) / 100; }

  function emit(obj) {
    var json = JSON.stringify(obj);
    document.getElementById("out").textContent =
      "@@MEASURE-B@@" + btoa(unescape(encodeURIComponent(json))) + "@@END@@";
  }

  window.addEventListener("load", function () {
    setTimeout(function () {
      try {
        var frame = document.getElementById("f");
        var doc = frame.contentDocument;
        frame.style.height = (doc.documentElement.scrollHeight + 200) + "px";
        setTimeout(function () {
          try { emit(collect(doc)); }
          catch (error) { emit({ error: String((error && error.stack) || error) }); }
        }, 200);
      } catch (error) { emit({ error: String(error) }); }
    }, 300);
  });

  // Whether a face actually resolved, by asking a canvas to measure the same
  // string in it and in a face nothing resolves to. Equal widths mean the
  // browser fell back, and every other number on this page is then suspect.
  function fontProbe(name) {
    try {
      var canvas = document.createElement("canvas");
      var ctx = canvas.getContext("2d");
      var sample = "Handgloves 0123456789 mmmiiillW";
      ctx.font = '16px "' + name + '"';
      var named = ctx.measureText(sample).width;
      ctx.font = '16px "__no_such_face_anywhere__"';
      var fallback = ctx.measureText(sample).width;
      return { name: name, named: r2(named), fallback: r2(fallback), resolved: Math.abs(named - fallback) > 0.5 };
    } catch (error) {
      return { name: name, named: null, fallback: null, resolved: null };
    }
  }

  function collect(doc) {
    var view = doc.defaultView;
    var sections = [].slice.call(doc.querySelectorAll("section.docx"));

    // docx-preview writes the page margins as the section's padding, so the
    // content box is the printable area — the frame every rect is recorded
    // against, and the same frame Word's page-relative numbers reduce to.
    var boxes = sections.map(function (s) {
      var rect = s.getBoundingClientRect();
      var cs = view.getComputedStyle(s);
      var padLeft = parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth);
      var padTop = parseFloat(cs.paddingTop) + parseFloat(cs.borderTopWidth);
      var padRight = parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth);
      var padBottom = parseFloat(cs.paddingBottom) + parseFloat(cs.borderBottomWidth);
      return {
        el: s,
        page: { x: rect.x, y: rect.y, w: rect.width, h: rect.height },
        content: {
          x: rect.x + padLeft, y: rect.y + padTop,
          w: rect.width - padLeft - padRight, h: rect.height - padTop - padBottom
        }
      };
    });

    function sectionIndexOf(el) {
      var section = el.closest("section.docx");
      return sections.indexOf(section);
    }

    // Every line rectangle the paragraph's text actually draws in. This is
    // what separates a centred paragraph from a left one: the box is the same,
    // the lines are not.
    function lineRectsOf(el) {
      var range = el.ownerDocument.createRange();
      range.selectNodeContents(el);
      var rects = [].slice.call(range.getClientRects());
      range.detach && range.detach();

      // Client rects arrive fragmented — one per text node per line — so they
      // are merged by line, on the vertical midpoint rather than the top,
      // because runs of different sizes on one line do not share a top edge.
      var lines = [];
      rects.forEach(function (rect) {
        if (rect.width <= 0 || rect.height <= 0) return;
        var mid = rect.y + rect.height / 2;
        var line = null;
        for (var i = 0; i < lines.length; i++) {
          if (Math.abs(lines[i].mid - mid) <= Math.max(2, rect.height / 2)) { line = lines[i]; break; }
        }
        if (line === null) {
          lines.push({ mid: mid, top: rect.y, bottom: rect.y + rect.height, left: rect.x, right: rect.x + rect.width });
        } else {
          line.top = Math.min(line.top, rect.y);
          line.bottom = Math.max(line.bottom, rect.y + rect.height);
          line.left = Math.min(line.left, rect.x);
          line.right = Math.max(line.right, rect.x + rect.width);
        }
      });
      lines.sort(function (a, b) { return a.top - b.top; });
      return lines;
    }

    // How many tables a table is standing inside.
    function depthOf(table) {
      var depth = 0;
      var at = table.parentNode;
      while (at && at.closest) {
        var outer = at.closest("table");
        if (outer === null) break;
        depth += 1;
        at = outer.parentNode;
      }
      return depth;
    }

    // The lines a cell's words are drawn on.
    //
    // Not lineRectsOf(td). A range over a cell's contents holds a block
    // element -- the paragraph docx-preview put there -- and a range holding a
    // block reports that block's box, which is the full width of the column
    // whichever way the text inside it is set. Measured that way a
    // right-ranged cell and a left one start in the same place, and the number
    // is the cell's, not the text's. So the paragraphs are measured, and the
    // cell is only measured directly when it holds none.
    function cellLinesOf(td) {
      var paras = [].slice.call(td.querySelectorAll("p"));
      if (paras.length === 0) return lineRectsOf(td);

      var lines = [];
      paras.forEach(function (p) { lines = lines.concat(lineRectsOf(p)); });
      lines.sort(function (a, b) { return a.top - b.top; });
      return lines;
    }

    var paragraphs = [];
    var articles = boxes.map(function (b) {
      return b.el.querySelector(":scope > article") || b.el;
    });

    articles.forEach(function (article, pageIndex) {
      // Direct children only: a paragraph inside a table cell belongs to the
      // table slice, and counting it here would shift every index.
      var kids = [].slice.call(article.children);
      kids.forEach(function (el) {
        if (el.tagName !== "P") return;

        var rect = el.getBoundingClientRect();
        var cs = view.getComputedStyle(el);
        var content = boxes[pageIndex].content;
        var lines = lineRectsOf(el);

        // Paragraph properties live on the <p>; run properties live on the
        // <span> docx-preview wraps the text in. Reading a face or a size off
        // the paragraph returns whatever the page inherits -- Times New Roman
        // at 16px, black, which is the browser's opinion and not the
        // document's. So the run side is read from the first element that
        // actually holds text.
        var runEl = el.querySelector("span, a, b, i, em, strong") || el;
        var rcs = view.getComputedStyle(runEl);

        paragraphs.push({
          pageIndex: pageIndex,
          text: el.textContent,
          // Page-relative, against the content box, in px at 96dpi.
          x: r2(rect.x - content.x),
          y: r2(rect.y - content.y),
          w: r2(rect.width),
          h: r2(rect.height),
          height: r2(rect.height),

          // Where the words are, which is not where the box is.
          lineCount: lines.length,
          lines: lines.map(function (line) {
            return {
              x: r2(line.left - content.x),
              y: r2(line.top - content.y),
              w: r2(line.right - line.left),
              h: r2(line.bottom - line.top)
            };
          }),

          // Paragraph-level: what the <p> itself says.
          textAlign: cs.textAlign,
          background: cs.backgroundColor,

          // Run-level: what the span inside it says.
          fontFamily: rcs.fontFamily,
          fontSize: r2(parseFloat(rcs.fontSize)),
          fontWeight: rcs.fontWeight,
          fontStyle: rcs.fontStyle,
          color: rcs.color,
          letterSpacing: rcs.letterSpacing === "normal" ? 0 : r2(parseFloat(rcs.letterSpacing)),
          textTransform: rcs.textTransform,
          textIndent: r2(parseFloat(cs.textIndent) || 0),
          computedLineHeight: cs.lineHeight === "normal" ? null : r2(parseFloat(cs.lineHeight)),
          marginTop: r2(parseFloat(cs.marginTop) || 0),
          marginBottom: r2(parseFloat(cs.marginBottom) || 0),
          paddingLeft: r2(parseFloat(cs.paddingLeft) || 0),
          paddingTop: r2(parseFloat(cs.paddingTop) || 0),
          paddingRight: r2(parseFloat(cs.paddingRight) || 0),
          paddingBottom: r2(parseFloat(cs.paddingBottom) || 0),
          borderTopColor: parseFloat(cs.borderTopWidth) > 0 ? cs.borderTopColor : null,
          borderRightColor: parseFloat(cs.borderRightWidth) > 0 ? cs.borderRightColor : null,
          borderBottomColor: parseFloat(cs.borderBottomWidth) > 0 ? cs.borderBottomColor : null,
          borderLeftColor: parseFloat(cs.borderLeftWidth) > 0 ? cs.borderLeftColor : null,

          // Absolute, for the gap arithmetic the node side does.
          absTop: r2(rect.y),
          absBottom: r2(rect.y + rect.height),
          drawnTop: lines.length ? r2(lines[0].top) : null,
          drawnBottom: lines.length ? r2(lines[lines.length - 1].bottom) : null
        });
      });
    });

    // Every table drawn on the page, cell by cell.
    //
    // Separate from the paragraph slice on purpose: the paragraphs above are
    // the article's own children, so a cell's paragraph is deliberately not
    // one of them. A table case asks about the cell, not about a paragraph
    // that happens to be inside one — where it sits, how wide it is, what it
    // is filled with and where its words are drawn inside it.
    //
    // Nested tables are collected too, in document order, each marked. A
    // table inside a cell is a table.
    var tables = [];

    articles.forEach(function (article, pageIndex) {
      var content = boxes[pageIndex].content;

      [].slice.call(article.querySelectorAll("table")).forEach(function (table) {
        var rect = table.getBoundingClientRect();
        var rows = [].slice.call(table.rows);

        tables.push({
          pageIndex: pageIndex,
          index: tables.length,
          // A table whose parent chain holds another table. Its own rows are
          // all that table.rows returns, so the two never double-count.
          nested: depthOf(table) > 0,
          depth: depthOf(table),

          x: r2(rect.x - content.x),
          y: r2(rect.y - content.y),
          w: r2(rect.width),
          h: r2(rect.height),

          rowCount: rows.length,
          rows: rows.map(function (tr, rowIndex) {
            var rowRect = tr.getBoundingClientRect();
            var column = 0;

            return {
              index: rowIndex,
              y: r2(rowRect.y - content.y),
              h: r2(rowRect.height),
              cells: [].slice.call(tr.cells).map(function (td, cellIndex) {
                var cellRect = td.getBoundingClientRect();
                var cs = view.getComputedStyle(td);
                var span = td.colSpan || 1;
                var at = column;
                var lines = cellLinesOf(td);
                // Alignment lives on the paragraph docx-preview put inside the
                // cell, never on the cell itself: reading it off the td
                // returns the browser's own default for a right-ranged column.
                var para = td.querySelector("p");
                var pcs = para === null ? null : view.getComputedStyle(para);

                column += span;

                return {
                  index: cellIndex,
                  column: at,
                  colSpan: span,
                  text: td.textContent,

                  x: r2(cellRect.x - content.x),
                  y: r2(cellRect.y - content.y),
                  w: r2(cellRect.width),
                  h: r2(cellRect.height),

                  background: cs.backgroundColor,
                  verticalAlign: cs.verticalAlign,
                  textAlign: pcs === null ? null : pcs.textAlign,

                  paddingLeft: r2(parseFloat(cs.paddingLeft) || 0),
                  paddingTop: r2(parseFloat(cs.paddingTop) || 0),
                  paddingRight: r2(parseFloat(cs.paddingRight) || 0),
                  paddingBottom: r2(parseFloat(cs.paddingBottom) || 0),

                  borderTopColor: parseFloat(cs.borderTopWidth) > 0 ? cs.borderTopColor : null,
                  borderRightColor: parseFloat(cs.borderRightWidth) > 0 ? cs.borderRightColor : null,
                  borderBottomColor: parseFloat(cs.borderBottomWidth) > 0 ? cs.borderBottomColor : null,
                  borderLeftColor: parseFloat(cs.borderLeftWidth) > 0 ? cs.borderLeftColor : null,

                  // Where the words are drawn inside the cell, which is not
                  // where the cell is: a right-ranged column's cells are the
                  // width of the column whichever way their text is set.
                  lineCount: lines.length,
                  lines: lines.map(function (line) {
                    return {
                      x: r2(line.left - content.x),
                      y: r2(line.top - content.y),
                      w: r2(line.right - line.left),
                      h: r2(line.bottom - line.top)
                    };
                  })
                };
              })
            };
          })
        });
      });
    });

    // Every shape drawn on the page.
    //
    // docx-preview renders a VML shape as an <svg> holding the drawn element
    // and, once settle has moved it out, a <foreignObject> beside it carrying
    // the words. Both halves are measured: a filled block with no words on it
    // and a block whose words did not draw look identical in a screenshot and
    // are not the same document.
    var shapes = [];

    articles.forEach(function (article, pageIndex) {
      var content = boxes[pageIndex].content;

      [].slice.call(article.querySelectorAll("svg")).forEach(function (svg) {
        var rect = svg.getBoundingClientRect();
        var drawn = svg.querySelector("rect, ellipse");
        var text = svg.querySelector("foreignObject");
        var lines = text === null ? [] : lineRectsOf(text);

        shapes.push({
          pageIndex: pageIndex,
          index: shapes.length,

          x: r2(rect.x - content.x),
          y: r2(rect.y - content.y),
          w: r2(rect.width),
          h: r2(rect.height),

          // Straight off the element docx-preview drew from the file's own
          // fillcolor and strokecolor, not off a computed style: an SVG
          // presentation attribute is what the file said, and the computed
          // value would fold in a stylesheet the document never asked for.
          shape: drawn === null ? null : drawn.tagName,
          fill: drawn === null ? null : drawn.getAttribute("fill"),
          stroke: drawn === null ? null : drawn.getAttribute("stroke"),
          strokeWidth: drawn === null ? null : drawn.getAttribute("stroke-width"),

          // Where the words are, and whether they were drawn at all. A
          // foreignObject left inside the rect reports zero for both.
          text: text === null ? "" : text.textContent,
          textDrawn: text !== null,
          lineCount: lines.length,
          lines: lines.map(function (line) {
            return {
              x: r2(line.left - content.x),
              y: r2(line.top - content.y),
              w: r2(line.right - line.left),
              h: r2(line.bottom - line.top)
            };
          })
        });
      });
    });

    // How tall a running strip actually draws.
    //
    // Not the container's height: docx-preview gives a header and a footer a
    // fixed reserve — a min-height cancelled by an equal negative margin, taken
    // from the section's header distance — so the box is the same size on a
    // page that shows nothing as on one that shows a bar. What a reader sees is
    // the content inside it, so that is what is measured.
    function drawnBox(el) {
      if (!el) return null;
      var top = Infinity, bottom = -Infinity, left = Infinity, right = -Infinity;
      var kids = el.querySelectorAll("*");
      for (var i = 0; i < kids.length; i++) {
        var r = kids[i].getBoundingClientRect();
        if (r.height <= 0 || r.width <= 0) continue;
        if (r.top < top) top = r.top;
        if (r.bottom > bottom) bottom = r.bottom;
        if (r.left < left) left = r.left;
        if (r.right > right) right = r.right;
      }
      if (bottom <= top) return null;
      return { top: top, bottom: bottom, left: left, right: right };
    }

    // The running strips, per page, measured against the paper rather than the
    // text column: a header sits outside the margins by design, and reporting
    // it content-relative would give every one of them a negative number.
    var furniture = boxes.map(function (b) {
      var out = {};
      ["header", "footer"].forEach(function (part) {
        var el = b.el.querySelector(":scope > " + part);
        var drawn = drawnBox(el);
        out[part] = drawn === null
          ? { present: el !== null, drawn: false, text: el ? el.textContent.trim() : "" }
          : {
            present: true,
            drawn: true,
            text: el.textContent.trim(),
            // From the top of the sheet, in px — the frame Word's own
            // HeaderDistance is measured in.
            y: r2(drawn.top - b.page.y),
            x: r2(drawn.left - b.page.x),
            h: r2(drawn.bottom - drawn.top),
            w: r2(drawn.right - drawn.left),
            // And from the bottom, which is what a footer distance means.
            fromBottom: r2(b.page.y + b.page.h - drawn.bottom)
          };
      });
      return out;
    });

    return {
      sections: boxes.map(function (b) {
        return {
          page: { w: r2(b.page.w), h: r2(b.page.h), x: r2(b.page.x), y: r2(b.page.y) },
          content: { w: r2(b.content.w), h: r2(b.content.h) }
        };
      }),
      furniture: furniture,
      paragraphs: paragraphs,
      tables: tables,
      shapes: shapes,
      fontProbe: fontProbe(BODY_FONT),
      // What the paginator did, straight from the page it ran on. Without
      // this a preview that failed to paginate is indistinguishable from one
      // whose document happened to fit.
      paginated: doc.body ? (doc.body.dataset.paginated || null) : null,
      tabsPlaced: doc.body ? (doc.body.dataset.tabs || null) : null,
      documentHeight: doc.documentElement.scrollHeight
    };
  }
})();
</script>
</body>
</html>
`;
}

/** Runs Chrome once with the flags that work headlessly here. */
async function runChrome(chrome, args) {
  const { stdout } = await execFileAsync(chrome, [
    "--headless=new",
    "--disable-gpu",
    "--hide-scrollbars",
    "--force-device-scale-factor=1",
    "--virtual-time-budget=15000",
    ...args,
  ], { maxBuffer: 64 * 1024 * 1024, timeout: 120_000, killSignal: "SIGKILL" });

  return stdout;
}

/** A same-origin static server for the harness and the page it iframes. */
async function startServer(root) {
  const types = { ".html": "text/html; charset=utf-8", ".png": "image/png" };

  for (let attempt = 1; ; attempt += 1) {
    try {
      return await new Promise((resolvePromise, rejectPromise) => {
        const server = createServer(async (request, response) => {
          try {
            const pathname = decodeURIComponent(new URL(request.url, "http://localhost").pathname);
            const file = join(root, pathname.replaceAll("..", ""));
            const body = await readFile(file);
            response.writeHead(200, {
              "Content-Type": types[extname(file)] ?? "application/octet-stream",
            });
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
      await sleep(1000);
    }
  }
}

function decodeMeasurement(dom) {
  const match = /@@MEASURE-B@@([A-Za-z0-9+/=]+)@@END@@/.exec(dom);
  if (match === null) {
    throw new Error("the harness produced no measurement (no marker in the dumped DOM)");
  }
  return JSON.parse(Buffer.from(match[1], "base64").toString("utf8"));
}

/** Crops each page out of the full-page screenshot, for the board. */
async function cropPages(fullPngPath, sections, outDir) {
  const full = PNG.sync.read(await readFile(fullPngPath));
  const written = [];

  for (let index = 0; index < sections.length && index < 8; index += 1) {
    const box = sections[index].page;
    const x = Math.max(0, Math.round(box.x));
    const y = Math.max(0, Math.round(box.y));
    const width = Math.min(Math.round(box.w), full.width - x);
    const height = Math.min(Math.round(box.h), full.height - y);
    if (width <= 0 || height <= 0) continue;

    const crop = new PNG({ width, height });
    PNG.bitblt(full, crop, x, y, width, height, 0, 0);
    const file = `preview-p${index + 1}.png`;
    await writeFile(resolve(outDir, file), PNG.sync.write(crop));
    written.push({ file, w: width, h: height });
  }

  return written;
}

/**
 * Measures one case's preview page and writes `measure-b.json` beside it.
 *
 * Never throws its measurements away: a crash is recorded in the JSON the
 * assertions read, so "the probe died" is a status on the board rather than a
 * stack trace in a log.
 */
export async function runPreviewProbe({ outDir, htmlPath, bodyFont, screenshots = true }) {
  const output = {
    probe: "B",
    error: null,
    viewportWidth: VIEWPORT_WIDTH,
    pxPerMm: PX_PER_MM,
    sections: [],
    furniture: [],
    paragraphs: [],
    tables: [],
    shapes: [],
    fontProbe: null,
    screenshots: [],
  };

  const chrome = chromePath();
  if (chrome === null) {
    output.error = "Chrome not found — set CHROME_PATH";
    await writeFile(resolve(outDir, "measure-b.json"), JSON.stringify(output, null, 2), "utf8");
    return output;
  }

  let server;
  let tempDir;

  try {
    tempDir = await mkdtemp(join(tmpdir(), "conf-b-"));
    await copyFile(htmlPath, join(tempDir, "case.html"));
    await writeFile(join(tempDir, "__probe-b.html"), harnessHtml(bodyFont), "utf8");

    server = await startServer(tempDir);

    const dom = await runChrome(chrome, [
      `--window-size=${VIEWPORT_WIDTH},900`,
      "--dump-dom",
      `http://127.0.0.1:${PORT}/__probe-b.html`,
    ]);

    const measured = decodeMeasurement(dom);
    if (measured.error) {
      throw new Error(`harness error: ${measured.error}`);
    }

    output.sections = measured.sections;
    output.furniture = measured.furniture ?? [];
    output.paragraphs = measured.paragraphs;
    output.tables = measured.tables ?? [];
    output.shapes = measured.shapes ?? [];
    output.fontProbe = measured.fontProbe;
    output.paginated = measured.paginated ?? null;
    output.tabsPlaced = measured.tabsPlaced ?? null;

    if (screenshots && measured.sections.length > 0) {
      const height = Math.max(900, Math.min(measured.documentHeight + 20, 8000));
      const fullPngPath = join(tempDir, "full.png");
      await runChrome(chrome, [
        `--window-size=${VIEWPORT_WIDTH},${height}`,
        `--screenshot=${fullPngPath}`,
        `http://127.0.0.1:${PORT}/case.html`,
      ]);
      output.screenshots = await cropPages(fullPngPath, measured.sections, outDir);
    }
  } catch (error) {
    output.error = String(error?.message ?? error);
  } finally {
    if (server) {
      server.closeAllConnections?.();
      await new Promise((done) => server.close(done));
    }
    if (tempDir) {
      await rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }

  await writeFile(resolve(outDir, "measure-b.json"), JSON.stringify(output, null, 2), "utf8");

  return output;
}

/**
 * The view a case is handed: the measurement, plus the lookups that keep an
 * assertion to one line.
 *
 * Everything geometric is in CSS pixels at 96dpi, page-relative to the content
 * box. `mm` and `pt` convert into that, so a case never writes 3.7795 itself.
 */
export function previewView(measure) {
  // Case-insensitive: `w:caps` prints a label in capitals without changing the
  // text, so the same anchor has to find it whichever engine is asked.
  const find = (anchor) => {
    const wanted = anchor.toLowerCase();
    return measure.paragraphs.find((p) => p.text.toLowerCase().includes(wanted)) ??
      missingParagraph(anchor);
  };

  const lastLine = (p) => p.lines[p.lines.length - 1];
  // A cell, by the words a reader sees in it. Same case-insensitivity as the
  // paragraph side, and for the same reason: a heading cell set in `w:caps`
  // says one thing in the file and prints another.
  const findCell = (anchor) => {
    const wanted = anchor.toLowerCase();

    for (const table of byDepth(measure.tables)) {
      for (const row of table.rows) {
        for (const cell of row.cells) {
          if ((cell.text ?? "").toLowerCase().includes(wanted)) {
            return { ...cell, pageIndex: table.pageIndex, table: table.index, row: row.index };
          }
        }
      }
    }

    return missingCell(anchor);
  };


  return {
    ...measure,

    /** Millimetres, in the pixels everything here is measured in. */
    mm: (value) => value * PX_PER_MM,
    /** Points, likewise. */
    pt: (value) => value * (96 / 72),
    /** A hex colour, in the `rgb(r, g, b)` docx-preview always reserialises to. */
    hex(value) {
      const clean = value.replace("#", "");
      const int = Number.parseInt(clean, 16);
      return `rgb(${(int >> 16) & 255}, ${(int >> 8) & 255}, ${int & 255})`;
    },

    para: find,
    paraAt: (index) => measure.paragraphs[index] ?? missingParagraph(`#${index}`),

    /** Whether the named face resolved rather than falling back silently. */
    fontResolved(name) {
      const probe = measure.fontProbe;
      if (probe === null || probe === undefined) return null;
      if (name !== undefined && probe.name !== name) return null;
      return probe.resolved;
    },

    /** The leading a paragraph is actually drawn at, line top to line top. */
    lineHeight(anchor) {
      const para = find(anchor);
      if (para.lines.length < 2) return para.computedLineHeight;
      const gaps = para.lines.slice(1).map((line, index) => line.y - para.lines[index].y);
      return gaps.reduce((sum, gap) => sum + gap, 0) / gaps.length;
    },

    /** The gap a reader sees under a paragraph: its bottom to the next one's top. */
    gapAfter(anchor) {
      const para = find(anchor);
      const index = measure.paragraphs.indexOf(para);
      const next = measure.paragraphs[index + 1];
      if (index === -1 || next === undefined) return null;
      // Across a page break there is no gap to measure, only a break.
      if (next.pageIndex !== para.pageIndex) return null;
      return Math.round((next.absTop - para.absBottom) * 100) / 100;
    },

    /** Where the drawn text starts, which is not where the box starts. */
    textLeft(anchor) {
      const para = find(anchor);
      // The last line, not the first: a first-line indent moves only the
      // first, and "where the paragraph sits" is where its body sits.
      const line = para.lines.length > 1 ? lastLine(para) : para.lines[0];
      return line === undefined ? null : line.x;
    },

    /** Where the first line starts — the half a first-line indent moves. */
    firstLineLeft(anchor) {
      const line = find(anchor).lines[0];
      return line === undefined ? null : line.x;
    },

    /** The furthest right the drawn text reaches. */
    textRight(anchor) {
      const lines = find(anchor).lines;
      return lines.length === 0 ? null : Math.max(...lines.map((line) => line.x + line.w));
    },

    /** The axis the drawn text is centred on. */
    textCentre(anchor) {
      const line = find(anchor).lines[0];
      return line === undefined ? null : line.x + line.w / 2;
    },

    /** How wide the drawn text is on its first line. */
    textWidth(anchor) {
      const line = find(anchor).lines[0];
      return line === undefined ? null : line.w;
    },

    /** How many pages the preview paginated to. */
    pageCount: () => measure.sections.length,

    /**
     * The running strip drawn on one page, counting pages from one.
     *
     * Measured from the top-left of the *sheet*, not the text column: a header
     * sits outside the margins by design, and against the column every one of
     * them would be a negative number.
     */
    furniture(part, page = 1) {
      const strip = measure.furniture?.[page - 1]?.[part];
      return strip ?? { present: false, drawn: false, text: "", y: null, x: null, h: null };
    },

    /** What the header prints on a given page. */
    headerText: (page = 1) => (measure.furniture?.[page - 1]?.header?.text ?? "").trim(),
    /** What the footer prints on a given page. */
    footerText: (page = 1) => (measure.furniture?.[page - 1]?.footer?.text ?? "").trim(),
    // --- tables -------------------------------------------------------------
    //
    // A cell is not a paragraph, and the lookups above will not find one: the
    // paragraph slice is the article's own children, so a cell's paragraph is
    // deliberately not in it. These are the same questions asked of a cell.

    /** Every table the preview drew, nested ones included, in document order. */
    tables: measure.tables ?? [],

    /** One table, counting from zero. */
    table(index = 0) {
      return (measure.tables ?? [])[index] ?? missingTable(`#${index}`);
    },

    /** Only the tables one is standing inside another. */
    nestedTables() {
      return (measure.tables ?? []).filter((table) => table.nested);
    },
    // --- shapes -------------------------------------------------------------

    /** Every shape the preview drew, in document order. */
    shapes: measure.shapes ?? [],

    /** One shape, counting from zero. */
    shape(index = 0) {
      return (measure.shapes ?? [])[index] ?? missingShape(`#${index}`);
    },

    /**
     * How wide the words on a shape are actually drawn.
     *
     * The assertion `liftShapeText` exists for. A `<foreignObject>` left where
     * docx-preview puts it — inside the `<rect>` — paints nothing, and every
     * line inside it measures zero. Nought here means a filled block with
     * invisible words, which is the failure that looks deliberate.
     */
    shapeTextWidth(anchor) {
      const wanted = anchor.toLowerCase();
      const shape = (measure.shapes ?? [])
        .find((entry) => (entry.text ?? "").toLowerCase().includes(wanted));

      if (shape === undefined) return null;
      return shape.lines.reduce((widest, line) => Math.max(widest, line.w), 0);
    },


    /** The cell whose text contains this anchor. First match wins. */
    cell: findCell,

    /**
     * Every cell whose text contains this anchor.
     *
     * How a repeated header row is counted: the words appear once in the file
     * and twice on the page, so the question "was it drawn again" is a
     * question about how many cells say it.
     */
    cells(anchor) {
      const wanted = anchor.toLowerCase();

      return byDepth(measure.tables).flatMap((table) =>
        table.rows.flatMap((row) =>
          row.cells
            .filter((cell) => (cell.text ?? "").toLowerCase().includes(wanted))
            .map((cell) => ({ ...cell, pageIndex: table.pageIndex, table: table.index, row: row.index }))
        )
      );
    },

    /** One cell by where it sits: table, row, then cell across the row. */
    cellAt(table, row, cell) {
      const found = (measure.tables ?? [])[table]?.rows?.[row]?.cells?.[cell];
      return found ?? missingCell(`#${table}.${row}.${cell}`);
    },

    /** The row a cell sits in, as the preview drew it. */
    row(table, row) {
      const found = (measure.tables ?? [])[table]?.rows?.[row];
      return found ?? { index: -1, missing: true, y: null, h: null, cells: [] };
    },

    /**
     * Where a cell's words start, which is not where the cell starts.
     *
     * The same distinction the paragraph side draws: a right-ranged column's
     * cells are the full width of the column whichever way their text is set,
     * and only the lines inside them move.
     */
    cellTextLeft(anchor) {
      const line = findCell(anchor).lines[0];
      return line === undefined ? null : line.x;
    },

    /** The furthest right a cell's words reach. */
    cellTextRight(anchor) {
      const lines = findCell(anchor).lines;
      return lines.length === 0 ? null : Math.max(...lines.map((line) => line.x + line.w));
    },

    /** The axis a cell's words are centred on. */
    cellTextCentre(anchor) {
      const line = findCell(anchor).lines[0];
      return line === undefined ? null : line.x + line.w / 2;
    },

    /** How far down the cell a cell's first line is drawn, from the cell's top. */
    cellTextTop(anchor) {
      const cell = findCell(anchor);
      const line = cell.lines[0];
      return line === undefined || cell.y === null ? null : line.y - cell.y;
    },

  };
}

/**
 * The tables a page drew, innermost first.
 *
 * A cell's text is everything printed inside it, so an outer cell matches
 * every anchor its inner table matches — and the inner cell is the one an
 * anchor names. Both other probes order theirs the same way, which is what
 * keeps one anchor meaning one cell across all three.
 */
function byDepth(tables) {
  return [...(tables ?? [])].sort((a, b) => (b.depth ?? 0) - (a.depth ?? 0));
}

/**
 * The stand-ins for a table and a cell the preview never drew.
 *
 * Null rather than undefined, for the same reason a missing paragraph is: an
 * assertion against a table that is not there should report the property it
 * wanted as `null`, not die reaching for it.
 */
function missingShape(anchor) {
  return {
    missing: true,
    anchor,
    pageIndex: null,
    index: -1,
    x: null,
    y: null,
    w: null,
    h: null,
    shape: null,
    fill: null,
    stroke: null,
    strokeWidth: null,
    text: "",
    textDrawn: false,
    lineCount: 0,
    lines: [],
  };
}

function missingTable(anchor) {
  return {
    missing: true,
    anchor,
    pageIndex: null,
    index: -1,
    nested: false,
    x: null,
    y: null,
    w: null,
    h: null,
    rowCount: 0,
    rows: [],
  };
}

function missingCell(anchor) {
  return {
    missing: true,
    anchor,
    index: -1,
    column: null,
    colSpan: null,
    text: "",
    x: null,
    y: null,
    w: null,
    h: null,
    background: null,
    verticalAlign: null,
    textAlign: null,
    paddingLeft: null,
    paddingTop: null,
    paddingRight: null,
    paddingBottom: null,
    borderTopColor: null,
    borderRightColor: null,
    borderBottomColor: null,
    borderLeftColor: null,
    lineCount: 0,
    lines: [],
  };
}

function missingParagraph(anchor) {
  return {
    missing: true,
    anchor,
    pageIndex: null,
    text: "",
    x: null,
    y: null,
    w: null,
    h: null,
    height: null,
    lineCount: 0,
    lines: [],
    textAlign: null,
    fontFamily: null,
    fontSize: null,
    fontWeight: null,
    fontStyle: null,
    color: null,
    background: null,
    letterSpacing: null,
    textTransform: null,
    textIndent: null,
    computedLineHeight: null,
    marginTop: null,
    marginBottom: null,
    paddingLeft: null,
    paddingTop: null,
    paddingRight: null,
    paddingBottom: null,
    borderTopColor: null,
    borderRightColor: null,
    borderBottomColor: null,
    borderLeftColor: null,
    absTop: null,
    absBottom: null,
    drawnTop: null,
    drawnBottom: null,
  };
}
