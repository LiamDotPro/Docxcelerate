/**
 * Reading facts back out of a packed `.docx`.
 *
 * This exists for one reason: docx-preview drops two things the file plainly
 * declares, and {@linkcode settleDocxPreview} has to write them back. It could
 * recompute them from the theme that produced the document — and did, for a
 * while — but that is a second copy of the packer's arithmetic, and a copy is
 * a thing that drifts. What the preview shows should come from the same place
 * Word reads: the file.
 *
 * Everything here is web-standard. `DecompressionStream` is in Node 20 and in
 * every browser the preview runs in, so the same code unzips in a scaffolded
 * workspace's browser tab and in the jsdom a site build renders in. There is no
 * `node:zlib` import, because half the callers have no Node to import it from.
 *
 * Reading is asynchronous — inflating is — so it is a separate step from
 * settling, which stays synchronous and is called with what this returns.
 * Making settle async instead would be worse: it mutates a DOM somebody is
 * about to measure, and a mutation that lands after the measurement is a bug
 * nobody sees until the numbers are wrong.
 *
 * @module
 */

/** A twip is a twentieth of a point. */
const TWIPS_PER_PT = 20;

/** One run of a paragraph, as the packed file records it. */
export interface PackedRun {
  /**
   * Letter spacing, in points, or `undefined` where the run sets none.
   *
   * docx-preview reads `w:spacing` only when its parent is `w:pPr` — the branch
   * is guarded `if (elem.localName == "pPr")` — so the identical element inside
   * `w:rPr`, which is where Word keeps character spacing, never reaches the DOM.
   */
  letterSpacingPt?: number;
}

/**
 * The gap a paragraph's border holds its words off by, per edge, in points.
 *
 * An edge with no border drawn on it is absent rather than zero: `w:space`
 * lives on the border element, so an undrawn edge has nowhere to record a gap
 * and Word holds nothing off it.
 */
export interface PackedBorderSpace {
  /** The gap above the text, when a top border is drawn. */
  top?: number;
  /** The gap to the right of the text, when a right border is drawn. */
  right?: number;
  /** The gap below the text, when a bottom border is drawn. */
  bottom?: number;
  /** The gap to the left of the text, when a left border is drawn. */
  left?: number;
}

/** One tab stop a paragraph declares, as the packed file records it. */
export interface PackedTabStop {
  /** Where the stop is, in points from the left margin. */
  positionPt: number;
  /** What the text does when it reaches it. */
  align: "left" | "center" | "right" | "decimal" | "bar";
  /** What fills the run-up to it. */
  leader: "none" | "dot" | "hyphen" | "underscore" | "middleDot";
}

/** One body paragraph, as the packed file records it. */
export interface PackedParagraph {
  /** The words it prints, which is how it is matched to what was drawn. */
  text: string;
  /** Its runs, in order. */
  runs: PackedRun[];
  /**
   * The space between each drawn border and the text, in points.
   *
   * `w:pBdr` carries a `w:space` on every edge, and docx-preview's
   * `parseBorderProperties` writes only the `border-*` shorthand. An edge with
   * no border is absent here, because `w:space` lives on the border element and
   * an undrawn edge has nowhere to record a gap.
   */
  borderSpacePt: PackedBorderSpace;
  /**
   * The tab stops the paragraph declares, in order.
   *
   * docx-preview parses these and can position text against them, but only
   * behind its `experimental` flag and only half a second after rendering —
   * which is no use to a renderer that lays out into a detached document and
   * serialises it. So the stops travel here and are applied where there is a
   * layout to apply them to.
   */
  tabStops: PackedTabStop[];
}

/**
 * What a packed document says about its own body paragraphs.
 *
 * @param packed The `.docx` bytes — the same ones handed to `renderAsync`.
 * @returns One entry per body paragraph, in document order.
 *
 * @example
 * ```ts
 * const packed = new Uint8Array(await blob.arrayBuffer());
 * await renderAsync(packed, body, head, { breakPages: true });
 * settleDocxPreview(body, model, await readPackedParagraphs(packed));
 * ```
 */
export async function readPackedParagraphs(packed: Uint8Array): Promise<PackedParagraph[]> {
  const xml = await readPart(packed, "word/document.xml");

  if (xml === undefined) {
    return [];
  }

  return paragraphsOf(xml).map(readParagraph);
}

// ---------------------------------------------------------------------------
// The zip
//
// A `.docx` is a zip, and the only entry wanted here is one small XML part. So
// this reads the central directory rather than unpacking anything: find the
// name, inflate that entry, stop.
// ---------------------------------------------------------------------------

/** One part of the package, as text, or `undefined` when it holds no such part. */
export async function readPart(packed: Uint8Array, wanted: string): Promise<string | undefined> {
  const view = new DataView(
    packed.buffer as ArrayBuffer,
    packed.byteOffset,
    packed.byteLength,
  );
  const latin1 = new TextDecoder("latin1").decode(packed);
  const end = latin1.lastIndexOf("PK");

  if (end === -1) {
    return undefined;
  }

  let at = view.getUint32(end + 16, true);
  const count = view.getUint16(end + 10, true);

  for (let index = 0; index < count; index += 1) {
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const name = latin1.slice(at + 46, at + 46 + nameLength);

    if (name === wanted) {
      const method = view.getUint16(at + 10, true);
      const size = view.getUint32(at + 20, true);
      const offset = view.getUint32(at + 42, true);
      // The local header repeats the name and extra fields, at its own lengths.
      const localName = view.getUint16(offset + 26, true);
      const localExtra = view.getUint16(offset + 28, true);
      const start = offset + 30 + localName + localExtra;
      const data = packed.slice(start, start + size);

      return new TextDecoder().decode(method === 0 ? data : await inflateRaw(data));
    }

    at += 46 + nameLength + extraLength + commentLength;
  }

  return undefined;
}

/** Raw deflate, through the one API both a browser and Node 20 have. */
async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new DecompressionStream("deflate-raw");
  const writer = stream.writable.getWriter();

  void writer.write(data as BufferSource);
  void writer.close();

  const chunks: Uint8Array[] = [];
  const reader = stream.readable.getReader();

  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value as Uint8Array);
  }

  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const out = new Uint8Array(total);
  let at = 0;

  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }

  return out;
}

// ---------------------------------------------------------------------------
// The XML
//
// Regex over string rather than a parser. The shapes wanted are shallow and
// fixed — `w:pPr` holds flat children, a run holds one `w:rPr` — and a reader
// with no dependency is one that keeps working when the parser it would have
// used moves on. It runs in a browser, too, where a parser is weight nobody
// asked to download.
// ---------------------------------------------------------------------------

/** The body's own paragraphs, with the ones inside tables left out. */
function paragraphsOf(xml: string): string[] {
  const body = between(xml, "w:body") ?? xml;
  const found: string[] = [];
  let depth = 0;
  let start = -1;

  for (const match of body.matchAll(/<w:(p|tbl)(?:\s[^>]*)?(\/?)>|<\/w:(p|tbl)>/g)) {
    const [whole, open, selfClose, close] = match;

    // A cell's paragraphs are the table's business, and counting them here
    // would put every later paragraph against the wrong words.
    if (open === "tbl" || close === "tbl") {
      depth += close === "tbl" ? -1 : selfClose === "/" ? 0 : 1;
      continue;
    }
    if (depth > 0) continue;

    if (open === "p") {
      if (selfClose === "/") {
        found.push(whole);
        continue;
      }
      if (start === -1) start = match.index;
    } else if (close === "p" && start !== -1) {
      found.push(body.slice(start, match.index + whole.length));
      start = -1;
    }
  }

  return found;
}

function readParagraph(xml: string): PackedParagraph {
  const properties = element(xml, "w:pPr");
  const borders = properties === null ? null : element(properties, "w:pBdr");

  return {
    text: textOf(xml),
    runs: elements(xml, "w:r").map((run) => {
      const spacing = numberAttribute(element(element(run, "w:rPr") ?? "", "w:spacing") ?? "", "w:val");

      return spacing === undefined ? {} : { letterSpacingPt: spacing / TWIPS_PER_PT };
    }),
    borderSpacePt: {
      top: borderSpace(borders, "top"),
      right: borderSpace(borders, "right"),
      bottom: borderSpace(borders, "bottom"),
      left: borderSpace(borders, "left"),
    },
    tabStops: readTabStops(properties),
  };
}

/** The stops a paragraph declares, in twips converted to points. */
function readTabStops(properties: string | null): PackedTabStop[] {
  const tabs = properties === null ? null : element(properties, "w:tabs");

  if (tabs === null) {
    return [];
  }

  const aligns = ["left", "center", "right", "decimal", "bar"] as const;
  const leaders = ["none", "dot", "hyphen", "underscore", "middleDot"] as const;

  return elements(tabs, "w:tab").map((stop) => {
    const align = attribute(stop, "w:val") ?? "left";
    const leader = attribute(stop, "w:leader") ?? "none";

    return {
      positionPt: (numberAttribute(stop, "w:pos") ?? 0) / TWIPS_PER_PT,
      align: (aligns as readonly string[]).includes(align)
        ? align as PackedTabStop["align"]
        : "left",
      leader: (leaders as readonly string[]).includes(leader)
        ? leader as PackedTabStop["leader"]
        : "none",
    };
  });
}

/** The gap one edge holds, when that edge is drawn at all. */
function borderSpace(borders: string | null | undefined, edge: string): number | undefined {
  if (borders === null || borders === undefined) return undefined;

  const found = element(borders, `w:${edge}`);
  if (found === null) return undefined;

  const style = attribute(found, "w:val");
  if (style === undefined || style === "none" || style === "nil") return undefined;

  // `w:space` is already in whole points — the one place Word does not count
  // in twips.
  return numberAttribute(found, "w:space") ?? 0;
}

function attribute(xml: string, name: string): string | undefined {
  return new RegExp(`${name}="([^"]*)"`).exec(xml)?.[1];
}

function numberAttribute(xml: string, name: string): number | undefined {
  const raw = attribute(xml, name);
  if (raw === undefined) return undefined;

  const value = Number.parseFloat(raw);
  return Number.isFinite(value) ? value : undefined;
}

/** The first `<name …/>` or `<name …>…</name>`, whole. */
function element(xml: string, name: string): string | null {
  const empty = new RegExp(`<${name}(\\s[^>]*)?/>`).exec(xml);
  const paired = new RegExp(`<${name}(\\s[^>]*)?>([\\s\\S]*?)</${name}>`).exec(xml);

  if (paired !== null && (empty === null || paired.index < empty.index)) {
    return paired[0];
  }
  return empty === null ? null : empty[0];
}

/** Every `<name …/>` or `<name …>…</name>`, in order. */
function elements(xml: string, name: string): string[] {
  const pattern = new RegExp(`<${name}(?:\\s[^>]*)?(?:/>|>[\\s\\S]*?</${name}>)`, "g");
  return [...xml.matchAll(pattern)].map((match) => match[0]);
}

/** What is inside `<name>…</name>`. */
function between(xml: string, name: string): string | null {
  return new RegExp(`<${name}(?:\\s[^>]*)?>([\\s\\S]*)</${name}>`).exec(xml)?.[1] ?? null;
}

/** The words a run of XML prints, tabs included as the tabs they are. */
function textOf(xml: string): string {
  return [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g)]
    .map((match) => (match[1] === undefined ? "\t" : decodeEntities(match[1])))
    .join("");
}

function decodeEntities(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&#39;", "'")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}
