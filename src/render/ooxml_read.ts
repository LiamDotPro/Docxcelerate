/**
 * Reading facts out of OOXML, by regex rather than by parser.
 *
 * Every reader here is deliberately shallow. The shapes being read are fixed
 * and flat — `w:pPr` holds flat children, a run holds one `w:rPr`, a chart's
 * cache holds a run of `c:pt` — and a reader with no dependency is one that
 * keeps working when the parser it would have used moves on. It runs in a
 * browser too, where a parser is weight nobody asked to download.
 *
 * The one shape this cannot read is one that contains itself. A table holds
 * tables and a paragraph holds none, so `elements` is safe for the second and
 * wrong for the first: matched lazily, `<w:tbl>…</w:tbl>` ends at the first
 * closing tag, which for a table holding a table is halfway through the outer
 * one. Anything that nests is walked by depth where it is read, not here.
 *
 * @module
 */

/**
 * The value of one attribute, wherever it appears in a run of XML.
 *
 * @param xml The XML to search.
 * @param name The attribute's name, prefix included.
 * @returns Its value, or `undefined` when the attribute is not there.
 */
export function attribute(xml: string, name: string): string | undefined {
  return new RegExp(`${name}="([^"]*)"`).exec(xml)?.[1];
}

/**
 * The same, as a number.
 *
 * @param xml The XML to search.
 * @param name The attribute's name.
 * @returns Its value, or `undefined` when it is absent or not a number.
 */
export function numberAttribute(xml: string, name: string): number | undefined {
  const raw = attribute(xml, name);

  if (raw === undefined) {
    return undefined;
  }

  const value = Number.parseFloat(raw);

  return Number.isFinite(value) ? value : undefined;
}

/**
 * The first `<name …/>` or `<name …>…</name>`, whole.
 *
 * @param xml The XML to search.
 * @param name The element's name, prefix included.
 * @returns The element as written, or `null` when there is none.
 */
export function element(xml: string, name: string): string | null {
  const empty = new RegExp(`<${name}(\\s[^>]*)?/>`).exec(xml);
  const paired = new RegExp(`<${name}(\\s[^>]*)?>([\\s\\S]*?)</${name}>`).exec(xml);

  if (paired !== null && (empty === null || paired.index < empty.index)) {
    return paired[0];
  }

  return empty === null ? null : empty[0];
}

/**
 * Every `<name …/>` or `<name …>…</name>`, in order.
 *
 * Only for an element that cannot contain itself — see the module note.
 *
 * @param xml The XML to search.
 * @param name The element's name.
 * @returns Each element as written.
 */
export function elements(xml: string, name: string): string[] {
  const pattern = new RegExp(`<${name}(?:\\s[^>]*)?(?:/>|>[\\s\\S]*?</${name}>)`, "g");

  return [...xml.matchAll(pattern)].map((match) => match[0]);
}

/**
 * What is inside `<name>…</name>`, greedily — so the outermost pair wins.
 *
 * @param xml The XML to search.
 * @param name The element's name.
 * @returns The contents, or `null` when the element is not there.
 */
export function between(xml: string, name: string): string | null {
  return new RegExp(`<${name}(?:\\s[^>]*)?>([\\s\\S]*)</${name}>`).exec(xml)?.[1] ?? null;
}

/**
 * XML text with its entities turned back into the characters they stand for.
 *
 * @param value The text as the file writes it.
 * @returns The text as a reader sees it.
 */
export function decodeEntities(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&#39;", "'")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}
