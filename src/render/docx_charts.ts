/**
 * Putting a document's charts into the package after it is packed.
 *
 * A chart is four things in a `.docx`, and only one of them is in
 * `document.xml`: the drawing that reserves the space, the chart part it points
 * at, the workbook that part opens, and the entries in the content types and
 * the relationships that let a reader find either. The library that lays the
 * document out writes one zip and offers no way to be handed a part, so the
 * package is opened here and the rest is added to it.
 *
 * The join between the two halves is a token. The renderer writes a drawing
 * whose relationship id is the chart node's own id — an id no relationship
 * could otherwise have — and this walks the packed parts looking for those
 * tokens, allocating a real id per part as it goes. Allocating them here rather
 * than at render time is what keeps them from colliding: the ids an image or a
 * running strip took are not known until the file exists, and a chart that
 * claimed `rId7` before an image did would point at the image.
 *
 * A document with no charts is returned untouched, bytes and all. Repacking a
 * file to change nothing in it is a risk taken for no reason.
 *
 * @module
 */

import type { DocumentModel, DocumentNode, GraphNode } from "../domain/types.ts";
import { cleanMinimalDocumentStyle } from "../project/style.ts";
import { chartPartOf, hasChartData } from "./chart_part.ts";
import { readZipEntries, writeZipEntries, type ZipEntry } from "./ooxml_zip.ts";

/**
 * The relationship id a chart's drawing carries until the package is opened.
 *
 * @param id The chart node's id.
 * @returns The token to write in the drawing's `r:id`.
 */
export function chartRelationshipToken(id: string): string {
  return `dxclChart_${id}`;
}

/** Finds the tokens a packed part carries, in the order they appear in it. */
const TOKEN_PATTERN = /r:id="dxclChart_([^"]*)"/g;

/** The relationship type a drawing uses to reach a chart part. */
const CHART_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
/** The relationship type a chart uses to reach its workbook. */
const PACKAGE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
/** What a chart part declares itself as in the content types. */
const CHART_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
/** What an embedded workbook declares itself as. */
const WORKBOOK_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/**
 * Adds a document's charts to the package that was packed for it.
 *
 * @param packed The `.docx` bytes, straight from the packer.
 * @param doc The document that produced them, which is where the chart data is.
 * @returns The package with its charts in it, or the bytes unchanged when the
 * document has none.
 */
export async function addChartsToPackage(
  packed: Uint8Array,
  doc: DocumentModel,
): Promise<Uint8Array> {
  const charts = chartsOf(doc);

  if (charts.size === 0) {
    return packed;
  }

  const entries = await readZipEntries(packed);
  const added: ZipEntry[] = [];
  const overrides: string[] = [];
  // Zipping a workbook is asynchronous and replacing a token is not, so the
  // workbooks are collected as the parts are walked and zipped once the walk
  // is over. Held here rather than beside the module, so that two documents
  // packed at once cannot end up with each other's charts.
  const workbooks: { name: string; parts: readonly ZipEntry[] }[] = [];
  const style = doc.style ?? cleanMinimalDocumentStyle;
  let count = 0;

  // A snapshot, because committing a relationships part appends to `entries`
  // when the part it belongs to had none — a running strip carrying nothing
  // but text is the case. Walking the array while it grows would visit what
  // was just added, which is harmless today only because a `.rels` part holds
  // no tokens.
  for (const entry of [...entries]) {
    const xml = textOf(entry);

    if (xml === undefined || !xml.includes("dxclChart_")) {
      continue;
    }

    const relationships = relationshipsFor(entries, entry.name);
    let next = nextRelationshipId(relationships.xml);

    const rewritten = xml.replaceAll(TOKEN_PATTERN, (whole, id: string) => {
      const node = charts.get(id);

      // A token with no node behind it cannot be turned into a chart, and a
      // relationship pointing at a part that was never written is a file Word
      // offers to repair. Leaving the token alone is the safer failure: the
      // drawing draws as an empty frame rather than the document refusing to
      // open. This is unreachable while the renderer and this module agree,
      // and is here because they are edited separately.
      if (node === undefined) {
        return whole;
      }

      count += 1;
      const relationshipId = `rId${next}`;
      next += 1;

      const built = chartPartOf(node, style);
      const chartName = `chart${count}`;

      added.push(
        textPart(`word/charts/${chartName}.xml`, built.chart),
        textPart(
          `word/charts/_rels/${chartName}.xml.rels`,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="${PACKAGE_RELATIONSHIP}" Target="../embeddings/${chartName}.xlsx"/>` +
            `</Relationships>`,
        ),
      );
      overrides.push(
        `<Override PartName="/word/charts/${chartName}.xml" ContentType="${CHART_CONTENT_TYPE}"/>`,
        `<Override PartName="/word/embeddings/${chartName}.xlsx" ContentType="${WORKBOOK_CONTENT_TYPE}"/>`,
      );
      // The workbook is a package of its own, so it is zipped with the same
      // writer and dropped in whole.
      workbooks.push({
        name: `word/embeddings/${chartName}.xlsx`,
        parts: built.workbook,
      });

      relationships.xml = withChild(
        relationships.xml,
        "Relationships",
        `<Relationship Id="${relationshipId}" Type="${CHART_RELATIONSHIP}" Target="charts/${chartName}.xml"/>`,
      );

      return `r:id="${relationshipId}"`;
    });

    replace(entries, entry.name, rewritten);
    relationships.commit();
  }

  if (count === 0) {
    return packed;
  }

  for (const workbook of workbooks) {
    added.push({ name: workbook.name, bytes: await writeZipEntries(workbook.parts) });
  }

  replace(
    entries,
    "[Content_Types].xml",
    withChild(
      textOf(entries.find((entry) => entry.name === "[Content_Types].xml")) ?? "",
      "Types",
      overrides.join(""),
    ),
  );

  return await writeZipEntries([...entries, ...added]);
}

/** Every chart in a document that has something to plot, by id. */
function chartsOf(doc: DocumentModel): Map<string, GraphNode> {
  const found = new Map<string, GraphNode>();

  const walk = (nodes: readonly DocumentNode[] | undefined): void => {
    for (const node of nodes ?? []) {
      if (node.kind === "graph") {
        if (hasChartData(node)) {
          found.set(node.id, node);
        }
        continue;
      }

      if ("children" in node) {
        walk(node.children);
      }
    }
  };

  walk(doc.nodes);
  walk(doc.header);
  walk(doc.footer);
  walk(doc.firstHeader);
  walk(doc.firstFooter);
  walk(doc.evenHeader);
  walk(doc.evenFooter);

  return found;
}

/**
 * The relationships part belonging to one part, created if it has none.
 *
 * A running strip that carries nothing but text has no relationships part at
 * all, and a chart in a footer is the first thing to need one.
 */
function relationshipsFor(
  entries: ZipEntry[],
  part: string,
): { xml: string; commit: () => void } {
  const slash = part.lastIndexOf("/");
  const directory = slash === -1 ? "" : part.slice(0, slash + 1);
  const name = part.slice(slash + 1);
  const path = `${directory}_rels/${name}.rels`;
  const existing = entries.find((entry) => entry.name === path);

  const state = {
    xml: existing === undefined
      ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
      : new TextDecoder().decode(existing.bytes),
    commit: (): void => {
      if (existing === undefined) {
        entries.push(textPart(path, state.xml));
        return;
      }

      replace(entries, path, state.xml);
    },
  };

  return state;
}

/**
 * A child added inside an element, whether or not it already has any.
 *
 * A relationships part with nothing in it is written `<Relationships …/>`, and
 * an append that only knew how to insert before a closing tag did nothing at
 * all to it. That is the worst shape a failure here can take: the chart part
 * was written, the drawing pointed at a relationship id, and the relationship
 * naming the part was silently dropped — a package Word will not open. So this
 * handles both forms and throws when it is handed neither, because there is no
 * such thing as adding a part that a reader cannot reach.
 */
function withChild(xml: string, element: string, child: string): string {
  const closing = `</${element}>`;

  if (xml.includes(closing)) {
    return xml.replace(closing, child + closing);
  }

  const empty = new RegExp(`<${element}(\\s[^>]*?)?/>`).exec(xml);

  if (empty !== null) {
    return xml.replace(empty[0], `<${element}${empty[1] ?? ""}>${child}${closing}`);
  }

  throw new Error(`Cannot add to <${element}>: the part holds no such element.`);
}

/** The first relationship number nothing in a part has taken. */
function nextRelationshipId(rels: string): number {
  const used = [...rels.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]));

  return Math.max(0, ...used) + 1;
}

/** A part's text, or nothing where the part is not XML. */
function textOf(entry: ZipEntry | undefined): string | undefined {
  if (entry === undefined || !entry.name.endsWith(".xml") && !entry.name.endsWith(".rels")) {
    return undefined;
  }

  return new TextDecoder().decode(entry.bytes);
}

/** Writes a part's text back, in the place the part already had. */
function replace(entries: ZipEntry[], name: string, xml: string): void {
  const at = entries.findIndex((entry) => entry.name === name);

  if (at !== -1) {
    entries[at] = textPart(name, xml);
  }
}

/** A part from its text. */
function textPart(name: string, xml: string): ZipEntry {
  return { name, bytes: new TextEncoder().encode(xml) };
}
