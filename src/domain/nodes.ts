/**
 * The nodes a built document is made of.
 *
 * Every node is JSON and carries no code, which is what lets a document be
 * built on one machine and rendered on another. A node whose content is not
 * settled yet says so — `mode`, `when`, `prompts` and `derivers` are how it
 * carries the question rather than the answer.
 *
 * @module
 */

import type { Condition, DataReference, DeriverInvocation, PromptSpec } from "./expressions.ts";
import type { DocumentStyle, TableAlign, TextAlign } from "./style.ts";

/** A plain JSON object, used wherever the model carries caller-defined data. */
export type JsonObject = Record<string, unknown>;

/** Discriminator naming which kind of node a {@linkcode DocumentNode} is. */
export type NodeKind =
  | "section"
  | "paragraph"
  | "image"
  | "graph"
  | "table"
  | "tableRow"
  | "tableCell"
  | "shape"
  | "tableOfContents"
  | "pageBreak"
  | "pageNumber"
  | "repeat";

/**
 * Whether a node's content is fixed at build time or written per document.
 *
 * `static` content ships in the model as-is. `dynamic` content is left for the
 * engine to fill from the node's prompts and the data it is given.
 */
export type NodeMode = "static" | "dynamic";

/**
 * The chart a {@linkcode GraphNode} draws.
 *
 * `bar` stands columns up, which is what a reader means by a bar chart;
 * `barHorizontal` lays them down, for categories whose names are too long to
 * sit under a column. Stacking is not a type of its own — a stacked bar is
 * still a bar, and saying it twice would let a document ask for a stacked pie.
 *
 * `radar` closes the category axis into a ring, so several series are read as
 * shapes rather than as rows: it answers "which of these is strong where", and
 * it is the wrong chart for "how much" because a reader cannot compare areas.
 * `bubble` is a scatter whose points carry a third figure as their size —
 * `sizes` on the series — for the case where two measures are not enough and a
 * third axis would be a lie.
 */
export type GraphType =
  | "bar"
  | "barHorizontal"
  | "line"
  | "area"
  | "pie"
  | "doughnut"
  | "scatter"
  | "radar"
  | "bubble";

/**
 * Whether the series stand beside one another, stack, or fill the plot.
 *
 * `true` stacks them, so the top of a stack is the total. `"percent"` stacks
 * them to the same height, so what is read is each series' share of its
 * category and the totals are deliberately thrown away — a 100% stacked chart,
 * as Word calls it. `false` and absent both stand them side by side.
 *
 * A third state on the same property rather than a chart type of its own, for
 * the reason {@linkcode GraphType} gives: a percent stack is still a bar.
 */
export type GraphStacking = boolean | "percent";

/**
 * What each data label prints.
 *
 * `value` is the point's own figure, `percent` its share of the whole, and
 * `category` the name it is counted against. The two `categoryAnd…` forms print
 * the name and then the figure, separated by a comma — "Rent, 42%".
 *
 * A share only exists on a pie or a doughnut, where the slices add up to
 * something. Word has no percentage to print on any other chart, so asking for
 * one there is an error rather than a label left blank.
 */
export type GraphDataLabels =
  | "value"
  | "percent"
  | "category"
  | "categoryAndValue"
  | "categoryAndPercent";

/** Where a chart's key sits, or that it has none. */
export type GraphLegend = "none" | "top" | "bottom" | "left" | "right";

/** One plotted run of numbers, and what to call it. */
export interface GraphSeries {
  /**
   * The name printed in the key.
   *
   * A chart with one series often needs no key — its title already says what
   * is plotted — which is why this is optional rather than the series' id.
   */
  label?: string;
  /**
   * The numbers, in the same order as the chart's categories.
   *
   * `null` is a gap rather than a zero: a month with no reading yet is not a
   * month that read nothing, and a line drawn through it would state a figure
   * nobody measured.
   */
  values: (number | null)[];
  /**
   * The colour this series is drawn in, as RGB hex without the `#`.
   *
   * Left out, the series takes its place in the theme's chart palette, which
   * is what keeps a document re-themeable: a colour written here is one the
   * next theme cannot change.
   */
  color?: string;
  /**
   * Whether this series' line is dashed or dotted rather than solid.
   *
   * What sets a forecast, a budget or a target apart from the actuals beside
   * it without leaning on colour alone, which a greyscale print throws away.
   * Read by the charts that draw a series as a line — `line`, `scatter` and
   * `radar` — and by no other: a bar, an area or a slice is a filled shape
   * with no line to dash. Left out, the line is solid.
   */
  dash?: "dashed" | "dotted";
  /**
   * How big each point is drawn, for a `bubble` chart and nothing else.
   *
   * The third figure a bubble carries: the categories are the x, `values` the
   * y, and these the area. A point with no size is not drawn, the same way a
   * point with no value is not — a bubble of no size is not a bubble of size
   * zero. Read by no other chart type, which is what keeps a scatter given
   * sizes a scatter rather than a bubble chart nobody asked for.
   */
  sizes?: (number | null)[];
}

/** The numbers a {@linkcode GraphNode} plots, and what they are counted against. */
export interface GraphData {
  /**
   * What the values are counted against — months, regions, quarters.
   *
   * A `scatter` and a `bubble` read these as their x values instead, so they
   * are numbers written as text there; anything that will not parse counts as
   * its position. Absent, the categories are the positions themselves.
   */
  categories?: string[];
  /** The runs of numbers to plot, in the order they are drawn and keyed. */
  series: GraphSeries[];
}

/**
 * The theme a style came from, by id.
 *
 * A string rather than a union of the shipped themes: the model is data that
 * travels, and a document set in a theme someone wrote themselves should say so
 * rather than claim to be one of ours. The shipped ids are listed by
 * {@linkcode https://docxcelerate.com/themes | the theme catalog} and typed as
 * `ShippedThemeId` in `docxcelerate/themes`.
 */

export interface BaseNode {
  /** Identifier, unique within the document. */
  id: string;
  /** Which kind of node this is. */
  kind: NodeKind;
  /** Heading shown for the node, where its kind renders one. */
  title?: string;
  /**
   * Which block style the document's theme should draw this node in.
   *
   * A name, not an appearance: `"band"`, `"panel"`, `"badge"`. What the name
   * looks like is the style's business, which is what keeps a colour out of a
   * component and lets one theme swap for another without touching a node.
   * Unknown names draw as nothing, so a document is never broken by a theme
   * that has not heard of one.
   */
  variant?: string;
  /** Whether the node may be dropped when there is nothing to say. */
  optional?: boolean;
  /** A test that decides, per document, whether the node is included. */
  when?: Condition;
  /** Instructions for an engine writing this node's content. */
  prompts?: PromptSpec[];
  /**
   * Settings for the engine writing this node — a model name, a token limit —
   * passed to its `AiClient` untouched. Nothing in docxcelerate reads them.
   * Only a generated node carries them.
   */
  aiOptions?: JsonObject;
  /** Computations to run before the node is written. */
  derivers?: DeriverInvocation[];
}

/** A titled group of nodes. */
export interface SectionNode extends BaseNode {
  /** Discriminator. */
  kind: "section";
  /**
   * Whether a renderer prints the section's title above its children.
   *
   * `false` keeps the title as the section's name — the id it derives, the TOC
   * entry, the address a request targets — without printing it. For a section
   * whose content already says what it is: a charges table whose header row
   * reads "Description", an address block headed by its own label. Absent
   * means printed, which is what a heading is for.
   */
  showTitle?: boolean;
  /** The nodes the section contains, in order. */
  children: DocumentNode[];
}

/** A block of prose, either written at build time or by the engine. */
export interface ParagraphNode extends BaseNode {
  /** Discriminator. */
  kind: "paragraph";
  /** Whether `text` is final or is to be written per document. */
  mode: NodeMode;
  /** The prose, when the paragraph is static or has a fallback. */
  text?: string;
  /**
   * How the paragraph's lines sit in the text column.
   *
   * On the node rather than only in the theme, because alignment is often
   * what the thing *is* rather than how it looks — a date ranged right is
   * ranged right in every theme, the way a money column is. A theme can still
   * say it for a named block, and the node wins when both do, exactly as a
   * cell wins over its column.
   */
  align?: TextAlign;
  /**
   * Pictures set in the line rather than above it.
   *
   * A mark beside a line of credit is one line; given a paragraph of its own
   * it becomes a picture with a caption under it, and a one-line footer bar
   * three lines deep. Each picture records where in `text` it sits, so `text`
   * stays exactly what it was — everything that reads a paragraph's words
   * still reads all of them, in order.
   */
  inlineImages?: InlineImage[];
  /**
   * The same words as `text`, cut where the setting changes.
   *
   * Present only when some of the words are set differently from the rest — a
   * term in bold, a phrase that links somewhere. `text` is still all of them,
   * joined, so everything that reads a paragraph's words goes on reading one
   * string.
   *
   * Each run carries its own words rather than a range over `text`, and that
   * is the decision the field rests on. A published paragraph still has
   * `{{tokens}}` in it, and whatever settles them changes its length: a range
   * counted at build time would be pointing at different words by the time a
   * document was written. A run's words are settled in place and joined again.
   *
   * A renderer honours the runs only while they still spell `text`. Something
   * that rewrote the one without the other has made the runs stale, and stale
   * words set in bold are worse than the right words set plain.
   */
  runs?: ParagraphRun[];
  /**
   * The paragraph's place in a list, when it is an item of one.
   *
   * On the paragraph rather than on a node around the items, so that an item
   * is everything a paragraph already is: it has an id, a condition can drop
   * it, a loop produces one per entry, an engine can write it. A list is its
   * items standing next to one another — which is also what it is in Word —
   * and a renderer counts them, the way it counts pages.
   *
   * Items that stand together are one list. Anything that is not an item
   * ends it, and a numbered list after that starts again at one.
   */
  list?: ParagraphList;
}

/**
 * How a row's stated height binds: `"atLeast"` is a floor the row grows past
 * when its contents need more, `"exact"` is a fixed depth that clips them.
 */
export type RowHeightRule = "atLeast" | "exact";

/**
 * Which way a cell's words run when they are turned on their side: `"up"` reads
 * from the foot of the cell to its head, as a spine label does, and `"down"`
 * from the head to the foot.
 */
export type CellTextDirection = "up" | "down";

/** Which kind of list a paragraph is an item of: marked with a bullet, or counted. */
export type ListKind = "bullet" | "number";

/**
 * The numerals a numbered list counts in at one level, and how they are set.
 *
 * The name is Word's own number format (`w:numFmt`), so `"lowerLetter"` counts
 * a, b, c and is followed by a full stop. The `Paren` form puts the numeral in
 * brackets with no stop, which is how a contract writes a sub-clause: (a), (i).
 */
export type ListStyle =
  | "decimal"
  | "lowerLetter"
  | "upperLetter"
  | "lowerRoman"
  | "upperRoman"
  | "decimalParen"
  | "lowerLetterParen"
  | "upperLetterParen"
  | "lowerRomanParen"
  | "upperRomanParen";

/**
 * The glyph a bulleted item is marked with, in place of the bullet.
 *
 * `"dash"` is an en dash (–), `"tick"` a check mark (✓), `"cross"` a ballot
 * x (✗) and `"square"` a small square (▪). Each is a plain Unicode glyph in
 * the paragraph's own face, like the bullet, so no symbol font is needed.
 */
export type ListMarker = "dash" | "tick" | "cross" | "square";

/** A paragraph's place in a list. */
export interface ParagraphList {
  /** Whether the item is marked or counted. */
  kind: ListKind;
  /**
   * How far under another item this one stands. Absent means not at all;
   * `1` is a point under an item, and a renderer goes as deep as nine.
   */
  level?: number;
  /**
   * The numerals this item's level counts in, for a numbered item. The first
   * item at a level in a list that says decides it for the whole list. Absent
   * means the default: figures, then letters, then roman numerals by depth,
   * each followed by a full stop. A bulleted item has none.
   */
  style?: ListStyle;
  /**
   * The number this item's level counts from, for a numbered item that opens
   * the level in its list. Absent means 1. A bulleted item has none.
   */
  start?: number;
  /**
   * The glyph this item is marked with, for a bulleted item. Each item says its
   * own, so ticks and crosses can stand in one list. Absent means the bullet:
   * •, ◦, ▪ by depth. A numbered item has none.
   */
  marker?: ListMarker;
}

/**
 * A stretch of a paragraph's words, and how it is set.
 *
 * What a run says is emphasis, destination and a named theme block, never appearance: there is no
 * colour and no size here, for the reason there is none on a node. A linked
 * phrase is drawn in the theme's accent because the theme says what a link
 * looks like.
 */
export interface ParagraphRun {
  /** The words. */
  text: string;
  /** Set heavier than the words around it — or, `false`, lighter than a bold block's. */
  bold?: boolean;
  /** Set in italics. */
  italic?: boolean;
  /**
   * Set with a rule beneath: `true` for a single rule, `"double"` for the
   * accounting double rule under a grand total. Left out or `false`, no rule.
   */
  underline?: boolean | "double";
  /** Set struck through. */
  strike?: boolean;
  /** Raised above the line or dropped below it, smaller. Absent, the words sit on it. */
  script?: "superscript" | "subscript";
  /**
   * An address the words link to: `https:`, `http:`, `mailto:` or `tel:`.
   *
   * Anything else prints as the words alone. The address may come out of a
   * request, and a preview draws a link as something a reader can click.
   */
  href?: string;
  /**
   * Which block style the theme draws this run in. A name, never an appearance:
   * the theme's `blocks` say what it looks like. Only the block's character
   * settings apply, and an unknown name draws the run plain.
   */
  variant?: string;
}

/** A picture set in a paragraph's line, and where along the text it sits. */
export interface InlineImage {
  /** The offset in the paragraph's `text` the picture is placed at. */
  at: number;
  /** The picture itself. */
  image: ImageNode;
}

/** A picture, either supplied at build time or produced by the engine. */
export interface ImageNode extends BaseNode {
  /** Discriminator. */
  kind: "image";
  /** Whether the image is fixed or is to be produced per document. */
  mode: NodeMode;
  /**
   * Where the picture comes from.
   *
   * A `data:` URI carries the bytes in the model, which is the only form that
   * survives being handed to an engine. A path or URL still draws on screen,
   * where a browser can fetch it, but cannot be packed into a Word file.
   */
  path?: string;
  /**
   * A raster to pack in place of an SVG.
   *
   * Word will not embed an SVG without one. Screen renderers ignore this and
   * draw the SVG itself, which is the sharper of the two.
   */
  fallbackPath?: string;
  /** Alternative text describing the image. */
  alt?: string;
  /**
   * Rendered width, in points, or `"full"` for the width of whatever it stands in:
   * the text column of the page as turned, or the inside of its table cell or shape.
   *
   * Left out, it follows the picture's own proportions from `height`, or is 120pt
   * when neither is given. With `"full"` the height follows the picture's
   * proportions unless `height` is given too, which draws exactly that box.
   */
  width?: number | "full";
  /** Rendered height, in points. Left out, it follows the picture's own proportions from the width. */
  height?: number;
  /**
   * Where the picture sits across the text column, when it is not flush left.
   *
   * On the node for the same reason a paragraph's is: a centred picture is centred
   * in every theme. Absent writes no alignment at all.
   */
  align?: TableAlign;
  /**
   * An address the picture links to: `https:`, `http:`, `mailto:` or `tel:`.
   *
   * Anything else draws the picture alone. The address may come out of a
   * request, and a preview draws a link as something a reader can click.
   */
  href?: string;
  /** What to show while the image has not been produced yet. */
  placeholder?: string;
}

/**
 * A chart, either given its data at build time or handed one by the engine.
 *
 * Packed as a real Word chart rather than a picture of one: the numbers travel
 * with the document, so a reader can select it, restyle it and open its data,
 * and it stays sharp at any zoom and on any printer. That is the whole reason
 * the node carries data instead of an image — a chart flattened to pixels at
 * build time is a chart nobody downstream can do anything with.
 */
export interface GraphNode extends BaseNode {
  /** Discriminator. */
  kind: "graph";
  /** Whether the data is fixed or is to be produced per document. */
  mode: NodeMode;
  /** Which chart to draw. */
  graphType: GraphType;
  /** The series to plot, when they are known at build time. */
  data?: GraphData;
  /** A caption printed beneath the chart. */
  caption?: string;
  /**
   * What the chart shows, for a reader who cannot see it.
   *
   * Absent means the drawing is described by its title, then its caption, then
   * its id — what every chart was described by before this existed.
   */
  alt?: string;
  /** What to show while the chart has no data yet. */
  placeholder?: string;
  /** How wide the chart is drawn, in points. The text column unless it is said. */
  width?: number;
  /** How deep it is drawn, in points. Seven twelfths of its width unless it is said. */
  height?: number;
  /**
   * Where the chart sits across the text column, when it is not flush left.
   *
   * On the node for the same reason a paragraph's is: a centred chart is centred
   * in every theme. Absent writes no alignment at all.
   */
  align?: TableAlign;
  /**
   * Where the key sits.
   *
   * Absent means the renderer decides, which is a key under the plot for
   * anything with more than one series and none for a chart with one — a key
   * naming the only thing on the chart is a line of text saying what the title
   * already said. A pie is the exception: its key names the slices, so it
   * keeps one however few there are.
   */
  legend?: GraphLegend;
  /**
   * Whether the series stack rather than stand beside one another, and whether
   * they stack to a total or to a share. See {@linkcode GraphStacking}.
   */
  stacked?: GraphStacking;
  /**
   * How the values are printed, as an OOXML number format — `"#,##0"`,
   * `"0.0%"`, `"£#,##0"`.
   *
   * The same codes Word's own number formatting uses, because they are what
   * the value axis and the data labels are given verbatim.
   */
  numberFormat?: string;
  /** What the category axis counts along. */
  categoryAxisTitle?: string;
  /** What the value axis measures. */
  valueAxisTitle?: string;
  /**
   * Where the value axis starts, in the units of the values. Left out, Word
   * picks it, which for most data is zero. A value below it is drawn clipped
   * at the axis; on a percent stack the axis is a share, so `0.5` is 50%.
   * Ignored on a pie and a doughnut, which have no axis.
   */
  valueAxisMin?: number;
  /**
   * Where the value axis ends, in the units of the values. Left out, Word
   * picks it. Must be above `valueAxisMin` when both are given.
   */
  valueAxisMax?: number;
  /**
   * The distance between ticks and gridlines on the value axis. Left out, Word
   * picks it. Must be greater than zero.
   */
  valueAxisStep?: number;
  /**
   * Whether each point prints a label, and what it prints. See
   * {@linkcode GraphDataLabels}.
   *
   * Off for most charts — a number on every point is a table drawn badly — and
   * worth turning on for a pie, where the slices are the reading. `true` prints
   * the figure the chart is read by: the share on a pie or a doughnut, the value
   * everywhere else. `false` and absent print nothing.
   */
  dataLabels?: boolean | GraphDataLabels;
}

/**
 * One column's shape, which every row shares.
 *
 * Widths are declared once, on the table, rather than per cell. A row that set
 * its own would be a row that disagrees with the row above it, and a table
 * whose columns do not line up is not a table.
 */
export interface TableColumn {
  /**
   * Width in millimetres, a share of the text column such as `"30%"`, or
   * `"auto"` to share out what the others leave. Left out, it is `"auto"`.
   *
   * Millimetres rather than points because a page is measured in them: a
   * 26mm money column against a 210mm page is a proportion a reader can check.
   * A percentage is for the table that has to keep its proportions on any
   * paper: it is a share of the text column, or of the paper for a `bleed`
   * table, worked out when the page is known.
   */
  width?: number | "auto" | `${number}%`;
  /** How this column's cells are aligned. Left unless it is said. */
  align?: TableAlign;
}

/**
 * A grid of cells, with the columns declared once.
 *
 * The rows are children rather than a field, so everything that works on a
 * node works on a row: a loop produces rows per entry, a condition drops one
 * per document, and each carries an id. That is the whole reason a row is a
 * node and not a tuple — an invoice's lines are a `.map()`, and the engine has
 * to be able to walk them without the table being a special case.
 */
export interface TableNode extends BaseNode {
  /** Discriminator. */
  kind: "table";
  /** The columns, left to right. */
  columns: TableColumn[];
  /** Where the table sits across the text column. Left unless it is said. */
  align?: TableAlign;
  /**
   * Whether the whole table stays on one page rather than breaking between rows.
   *
   * For a short table that reads as one thing — a totals box, bank details, a
   * signature block — where a page break between two of its rows leaves a
   * subtotal on one page and the total it leads to on the next. Word moves the
   * whole table over instead. Every row is also kept whole, as `keepTogether`
   * on a row does. A table taller than a page breaks anyway, having nowhere
   * else to go. Left out, Word may break the table between any two rows.
   */
  keepTogether?: boolean;
  /** The rows, and any loops that produce them. */
  children: DocumentNode[];
}

/** One row of a {@linkcode TableNode}. */
export interface TableRowNode extends BaseNode {
  /** Discriminator. */
  kind: "tableRow";
  /**
   * Whether this row heads the table.
   *
   * A header row is drawn as one and repeats at the top of every page the
   * table runs onto, which is a thing only the renderer can do — a second
   * header written into the body would be a row of text that says the same
   * words in the wrong place.
   */
  header?: boolean;
  /**
   * Whether the row stays whole on one page rather than breaking across two.
   *
   * For a row holding several lines — an invoice line with a wrapped
   * description, a clause in a schedule — where a break through the middle
   * leaves half a row at the foot of one page and half at the head of the next.
   * Word moves the whole row over instead. A row taller than a page breaks
   * anyway, having nowhere else to go.
   * Left out, Word may break the row wherever the page ends.
   */
  keepTogether?: boolean;
  /** The row's stated height, in points. Absent, the row is as tall as its contents. */
  height?: number;
  /**
   * How `height` binds. Absent means `"atLeast"`; only `"exact"` is carried,
   * so the model says only what departs from the default.
   */
  heightRule?: RowHeightRule;
  /** The cells, left to right. */
  children: DocumentNode[];
}

/** One cell of a {@linkcode TableRowNode}. */
export interface TableCellNode extends BaseNode {
  /** Discriminator. */
  kind: "tableCell";
  /** How many columns this cell runs across. One unless it is said. */
  span?: number;
  /** Alignment, when this cell departs from its column's. */
  align?: TableAlign;
  /** Turns the cell's words on their side. Absent, they run across. */
  textDirection?: CellTextDirection;
  /** What the cell holds — paragraphs, usually, but any node fits. */
  children: DocumentNode[];
}


/**
 * A drawn box with words on it.
 *
 * The one thing in the model that is a *shape* rather than text: Word draws a
 * real rectangle, and the paragraphs inside it sit on top of the fill rather
 * than beside it. A callout, a status block, a coloured panel a heading stands
 * in — the things a document reaches for when a paragraph with a background is
 * not enough because the box has to be a stated size.
 *
 * That size is the whole difference from a filled table cell, which is the
 * other way to draw a box here and the right one when the box should grow with
 * its text. A shape is a box you have decided the dimensions of; a cell is a
 * box the text decides. Neither is a substitute for the other, and a document
 * that wants the second should use a `<Table>` with one cell in it.
 *
 * What it looks like — the fill, the rule around it, the room inside, how the
 * words sit against its height — is the theme's, through `variant`, exactly as
 * it is for every other node. A shape names what it is; the style says what
 * that looks like.
 */
export interface ShapeNode extends BaseNode {
  /** Discriminator. */
  kind: "shape";
  /**
   * How wide the box is drawn, in points.
   *
   * The text column's full width unless it is said, which is the width a
   * banner or a callout wants and the only width that needs no arithmetic
   * from the document.
   */
  width?: number;
  /**
   * How deep the box is drawn, in points.
   *
   * A shape does not grow to fit its words — that is what makes it a shape —
   * so this is a decision the document or its theme has to make. The theme's
   * `heightPt` is taken when the node says nothing, because a block that
   * declares a depth is already saying exactly this.
   */
  height?: number;
  /** What is drawn on the box: paragraphs, usually, but any node fits. */
  children: DocumentNode[];
}
/** A table of contents, built from the sections around it. */
export interface TableOfContentsNode extends BaseNode {
  /** Discriminator. */
  kind: "tableOfContents";
}

/**
 * Where one page ends and the next begins.
 *
 * Only ever written where the break is part of what the document *is* — an
 * invoice whose payment details belong on their own page, a contract whose
 * signature block must not be orphaned. Breaking to control where a paragraph
 * happens to land is a job for the margins, not for a node.
 */
export interface PageBreakNode extends BaseNode {
  /** Discriminator. */
  kind: "pageBreak";
}

/** Which page this is, counted while the document is laid out. */
export type PageNumberFormat = "current" | "total" | "currentOfTotal";

/**
 * The page number, filled in by whatever lays the pages out.
 *
 * A build cannot know it: how many pages a document runs to depends on the
 * page size, the font and how much the engine wrote into every dynamic node.
 * So the node says which form it wants and the renderer counts.
 */
export interface PageNumberNode extends BaseNode {
  /** Discriminator. */
  kind: "pageNumber";
  /** Which form to print. Defaults to `currentOfTotal`. */
  format?: PageNumberFormat;
  /** What sits between the two numbers in `currentOfTotal`. Defaults to ` / `. */
  separator?: string;
  /**
   * Words printed before the number: `"Page "` for "Page 3 of 10".
   *
   * Printed as written, spaces included, and set in the same type as the
   * numbers. Absent prints nothing before them.
   */
  prefix?: string;
  /**
   * How the line sits across the column, the same as a paragraph's.
   *
   * Absent writes no `w:jc`, so the line inherits whatever the style around it
   * decides, which is what a page number did before it could say. A theme can
   * also say it through `variant`, and when both do the node wins.
   */
  align?: TextAlign;
}

/**
 * A body repeated once per entry in a request-time collection.
 *
 * A build cannot unroll this the way it unrolls a branch: the length of
 * `source` is not known until a document is written. So the loop itself is
 * what gets published, and the engine walks it. Each pass binds the entry under
 * `as` and its position under `indexAs`, and its one-based place among the entries that pass `where` under `positionAs`, both readable through `ctx`, and
 * suffixes child ids with the index so they stay unique across passes.
 */
export interface RepeatNode extends BaseNode {
  /** Discriminator. */
  kind: "repeat";
  /** The collection to walk, resolved per document. */
  source: DataReference;
  /** The `ctx` key each entry is bound to. */
  as: string;
  /** The `ctx` key each entry's zero-based position is bound to. */
  indexAs: string;
  /**
   * The `ctx` key each walked entry's position, counting from one, is bound to.
   *
   * Counted over the entries that pass `where`, so a filtered loop numbers what
   * the recipient sees. Absent on a loop published before positions existed,
   * which binds none.
   */
  positionAs?: string;
  /**
   * A test each entry has to pass to be walked at all.
   *
   * This is what a `.filter()` before the `.map()` becomes. The build cannot
   * apply it — which entries there are belongs to the request — so the test
   * travels with the loop and the engine applies it per entry.
   */
  where?: Condition;
  /** The nodes repeated for every entry, in order. */
  children: DocumentNode[];
}

/** Any node a document can contain. */
export type DocumentNode =
  | SectionNode
  | ParagraphNode
  | ImageNode
  | GraphNode
  | TableNode
  | TableRowNode
  | TableCellNode
  | ShapeNode
  | TableOfContentsNode
  | PageBreakNode
  | PageNumberNode
  | RepeatNode;

/**
 * What a document says about itself to the software that files it — Word's
 * File > Info, Explorer and Finder, and any DMS that indexes on properties.
 *
 * Written to `docProps/core.xml` and `docProps/custom.xml`. None of it is
 * printed. The title is not here: it is the document's own `title`, and is
 * written as the file's title whether or not it is printed.
 */
export interface DocumentProperties {
  /** Who the document is from — Word's Author, and its Last Modified By. */
  author?: string;
  /** What the document is about, in a phrase. */
  subject?: string;
  /** Words to find it by, as one string — Word shows them as Tags. */
  keywords?: string;
  /** A longer note about the document — Word's Comments. */
  description?: string;
  /**
   * Named values a DMS can index and filter on, such as an invoice number.
   * Every value is text. One that comes out empty is left out of the file,
   * rather than filed as a blank.
   */
  custom?: Record<string, string>;
}

/**
 * A built document: the JSON an engine is handed to write one copy from.
 *
 * @example Rendering a built document to a DOCX blob
 * ```ts
 * import { createDocxBlob } from "@docxcelerate/docxcelerate/docx";
 * import type { DocumentModel } from "@docxcelerate/docxcelerate";
 *
 * const model: DocumentModel = {
 *   schemaVersion: "docxcelerate.letter/v0",
 *   id: "welcome",
 *   title: "Welcome",
 *   nodes: [{ id: "hello", kind: "paragraph", mode: "static", text: "Hello." }],
 * };
 *
 * const blob = await createDocxBlob(model);
 * ```
 */
export interface DocumentModel {
  /** The model version, so a reader knows what it is looking at. */
  schemaVersion: "docxcelerate.letter/v0";
  /** Identifier for the document. */
  id: string;
  /** The document's title. */
  title: string;
  /** How the document looks; a renderer default applies when absent. */
  style?: DocumentStyle;
  /**
   * The language the document is written in, as a BCP 47 tag such as `fr-FR`.
   *
   * Absent on a document that never said, which renders as it always has. When
   * present, Word proofs the text in this language, derivers are told it, and
   * dynamic nodes are asked to be written in it.
   */
  locale?: string;
  /**
   * Whether Word opens the document with Track Changes on.
   *
   * Absent on a document that never asked, which packs as it always has. It is
   * a setting, not content: there is nothing to draw until someone edits the
   * file in Word, so a preview shows the document exactly as it is.
   */
  trackChanges?: boolean;
  /** Anything the caller wants to carry alongside the document. */
  metadata?: JsonObject;
  /** What the file says about itself, beyond its title. Absent writes only the title. */
  properties?: DocumentProperties;
  /** The body of the document, in order. */
  nodes: DocumentNode[];
  /**
   * Nodes drawn at the top of every page.
   *
   * Running furniture, not the first thing in the body: it repeats, and it sits
   * outside the text the margins measure. A letterhead that should appear once
   * belongs in `nodes`.
   */
  header?: DocumentNode[];
  /** Nodes drawn at the foot of every page. */
  footer?: DocumentNode[];
  /**
   * Nodes drawn at the top of the first page, in place of `header`.
   *
   * Present only when the document said its first page differs — a letter
   * whose letterhead *is* the top of page one does not want the running strip
   * repeating above it. An empty array means the first page shows nothing
   * where the other pages show `header`.
   */
  firstHeader?: DocumentNode[];
  /** Nodes drawn at the foot of the first page, in place of `footer`. */
  firstFooter?: DocumentNode[];
  /**
   * Nodes drawn at the top of left-hand pages, in place of `header`.
   *
   * A document printed on both sides and bound has two kinds of page, not one.
   * The reference belongs at the *outside* edge of each — right on a recto,
   * left on a verso — so it is always the corner a thumb reaches, and a folio
   * that sat in the same place on every sheet would sit in the gutter on half
   * of them.
   *
   * Naming either of these turns on Word's `w:evenAndOddHeaders`, which is a
   * setting for the whole document rather than for one section: from then on
   * `header` and `footer` are what a *right-hand* page shows.
   */
  evenHeader?: DocumentNode[];
  /** Nodes drawn at the foot of left-hand pages, in place of `footer`. */
  evenFooter?: DocumentNode[];
}
