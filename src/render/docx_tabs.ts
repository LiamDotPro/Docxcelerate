/**
 * Putting text on the tab stops the file declares.
 *
 * A tab stop is the only paragraph property in the framework whose effect
 * cannot be written as a style. Every other one — a fill, a border, an indent,
 * a leading — is a number CSS also has a word for, so settling it is a matter
 * of writing that number down. A stop is different: where the text after a tab
 * lands depends on how wide the text *before* it drew, so nothing can be
 * decided until something has laid the line out.
 *
 * That is why this is separate from settling and lives beside the paginator
 * instead. Both need a browser; neither can run in the jsdom a site build
 * renders in.
 *
 * docx-preview has its own implementation and it is not usable here: it is
 * behind an `experimental` flag, and it runs on a 500ms timer after rendering,
 * which is no use to a renderer that lays out into a detached document and
 * serialises it immediately. What it emits for `<w:tab/>` is a bare
 * `<span> </span>` — one space, no class — and that span is what gets widened
 * here, to the width that puts what follows it on the stop.
 *
 * @module
 */

import type { PackedParagraph, PackedTabStop } from "./docx_packed.ts";

/** A point, in the CSS pixels a browser lays out in at 96dpi. */
const PX_PER_PT = 96 / 72;

/** What a leader draws between the text and the stop it runs up to. */
const LEADERS: Record<string, string> = {
  dot: ".",
  hyphen: "-",
  underscore: "_",
  middleDot: "·",
};

/**
 * Positions every tabbed line against the stops its paragraph declares.
 *
 * Call it after `renderAsync` and after settling, in something that lays out.
 * Without a layout every width measures zero and nothing moves, which is the
 * safe direction: the preview stays as docx-preview drew it.
 *
 * @param container The element the preview was rendered into.
 * @param packed What the file says about its paragraphs, from
 * `readPackedParagraphs`.
 * @returns How many tabs were placed.
 */
export function applyTabStops(
  container: Element,
  packed: readonly PackedParagraph[],
): number {
  const withStops = packed.filter((paragraph) => paragraph.tabStops.length > 0);

  if (withStops.length === 0) {
    return 0;
  }

  const view = container.ownerDocument.defaultView;

  if (view === null) {
    return 0;
  }

  const paragraphs = [...container.querySelectorAll("p")] as HTMLElement[];
  const claimed = new Set<Element>();
  let placed = 0;

  for (const source of withStops) {
    const wanted = normalise(source.text);
    const element = paragraphs.find(
      (candidate) => !claimed.has(candidate) && normalise(candidate.textContent ?? "") === wanted,
    );

    if (element === undefined) {
      continue;
    }

    claimed.add(element);
    placed += placeTabs(element, source.tabStops, view);
  }

  return placed;
}

/** Whitespace-insensitive, because a renderer may rewrap but never rewords. */
function normalise(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/**
 * The tab spans in one paragraph, widened onto their stops.
 *
 * docx-preview draws a tab as a span holding a single space, so they are found
 * by that shape. Each takes the first stop to its right, which is what Word
 * does — a tab advances to the next stop, and a line with two tabs uses two
 * stops in order.
 */
function placeTabs(
  paragraph: HTMLElement,
  stops: readonly PackedTabStop[],
  view: Window,
): number {
  const left = paragraph.getBoundingClientRect().left;

  if (left === 0 && paragraph.getBoundingClientRect().width === 0) {
    return 0;
  }

  // docx-preview draws a tab as a span holding one em space — U+2003, not the
  // ordinary space it looks like in a dump. Matching on `" "` found nothing at
  // all, which is a silent nothing: the preview kept the em space, the text
  // after it landed wherever that put it, and every assertion about the stop
  // was measuring a document nobody had touched.
  const tabs = [...paragraph.querySelectorAll("span")]
    .filter((span) => span.children.length === 0)
    .filter((span) => /^[\s ]+$/.test(span.textContent ?? "")) as HTMLElement[];

  let placed = 0;

  for (const tab of tabs) {
    // Where the tab sits now, and therefore how much text precedes it.
    const before = tab.getBoundingClientRect().left - left;
    const stop = stops.find((candidate) => candidate.positionPt * PX_PER_PT > before);

    if (stop === undefined) {
      continue;
    }

    const target = stop.positionPt * PX_PER_PT;
    const width = stop.align === "left" ? target - before : trailingWidth(tab, target, left);

    if (!Number.isFinite(width) || width < 0) {
      continue;
    }

    tab.style.display = "inline-block";
    tab.style.width = `${round(width)}px`;
    tab.style.whiteSpace = "pre";

    if (stop.leader !== "none") {
      drawLeader(tab, stop, view);
    }

    placed += 1;
  }

  return placed;
}

/**
 * How wide a tab must be for what follows it to *end* on the stop.
 *
 * A right stop is the one a contents line is built on: the description runs
 * from the margin, and the page number's last digit lands exactly on the right
 * edge however long the description was. So the width is the distance to the
 * stop less the width of everything after the tab — measured, because that is
 * the only way to know it.
 */
function trailingWidth(tab: HTMLElement, target: number, left: number): number {
  const range = tab.ownerDocument.createRange();

  range.setStartAfter(tab);
  range.setEndAfter(tab.parentNode?.lastChild ?? tab);

  const trailing = range.getBoundingClientRect().width;

  range.detach?.();

  return target - (tab.getBoundingClientRect().left - left) - trailing;
}

/**
 * The dots that run up to a stop.
 *
 * Real leader characters in the paragraph's own face, because that is what Word
 * draws — a dotted border would be a different thing that happens to look
 * similar at one size and not at another.
 *
 * Counted to fit rather than overfilled and clipped. The first version wrote a
 * generous run of dots and hid the overflow, which looks right and measures
 * wrong: `overflow: hidden` clips what is *painted*, and the text's own client
 * rectangles still report the full unclipped run. Anything reading the DOM —
 * this suite's own probe included — then sees a line 58px wider than the sheet
 * and reports a stop that was in fact exactly right. A preview that is correct
 * only until something measures it is the kind of correct this whole project
 * exists to refuse.
 */
function drawLeader(tab: HTMLElement, stop: PackedTabStop, view: Window): void {
  const character = LEADERS[stop.leader];

  if (character === undefined) {
    return;
  }

  const width = Number.parseFloat(tab.style.width) || 0;
  const step = characterWidth(tab, character, view);

  if (step <= 0) {
    return;
  }

  // Floor, so the run stops short of the stop rather than past it. A part of a
  // dot is not a dot, and Word cannot draw one either.
  tab.textContent = character.repeat(Math.max(0, Math.floor(width / step)));
  tab.style.verticalAlign = "bottom";
}

/** How wide one leader character draws, in the face the tab is set in. */
function characterWidth(tab: HTMLElement, character: string, view: Window): number {
  const ruler = tab.ownerDocument.createElement("span");

  ruler.textContent = character.repeat(20);
  ruler.setAttribute("style", "position:absolute;visibility:hidden;white-space:pre");
  // Inside the tab, so it inherits the face, the size and any tracking the
  // paragraph is set in — a dot measured in the page's default font would size
  // the run for a document other than this one.
  tab.appendChild(ruler);

  const width = ruler.getBoundingClientRect().width / 20;

  ruler.remove();
  void view;

  return width;
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}
