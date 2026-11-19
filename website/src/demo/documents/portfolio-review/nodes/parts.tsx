import { Cell, type Nodes, Paragraph, Row, Table, Text } from "docxcelerate/template";
import { type BarGrid, barGrid, type Mark, type Upright, uprightGrid } from "../geometry.ts";

export type { Mark, Upright };

/**
 * The pieces every page is built from: an opening, a rubric, a bar.
 *
 * Kept in one file because they are vocabulary rather than content — each is
 * a few lines, and a reader looking for "how is a section opened" should find
 * it beside "how is a bar drawn".
 */

/** "1  Summary" over its standfirst and a dark rule. */
export const Opening: Nodes<{ id: string; number: string; title: string; standfirst?: string }> = (
  { id, number, title, standfirst },
) => (
  <Table id={`${id}-opening`} variant="opening" columns={[{ width: 15 }, { width: "auto" }]}>
    <Row>
      <Cell variant="openingNumber">{number}</Cell>
      <Cell>
        <Paragraph variant="openingTitle">{title}</Paragraph>
        {standfirst && <Paragraph variant="standfirst">{standfirst}</Paragraph>}
      </Cell>
    </Row>
  </Table>
);

/** A tracked label over a hairline, with an optional note on the right. */
export const Rubric: Nodes<{ id: string; label: string; note?: string; variant?: string }> = (
  { id, label, note, variant = "rubric" },
) => {
  // `note ?? ""` rather than `note && …`: the branch compiler turns a
  // conditional into a condition, and a paragraph holds only text.
  const tab = note ? "\t" : "";
  return (
    <Paragraph id={`${id}-rubric`} variant={variant}>
      {label}
      {tab}
      <Text variant="rubricNote">{note ?? ""}</Text>
    </Paragraph>
  );
};

/**
 * A horizontal bar drawn as table cells.
 *
 * Cells, not shapes: Word and docx-preview both lay a table out to the
 * twentieth of a point, so a bar sixty per cent long is sixty per cent long in
 * both. See {@linkcode barGrid} for how marks become columns.
 */
export const Bar: Nodes<{
  id: string;
  width: number;
  heightPt: number;
  insetPt?: number;
  track: string;
  marks: Mark[];
  /** Points the cell adds under the bar, taken off its last row. */
  trailingPt?: number;
}> = ({ id, width, heightPt, insetPt = 0, track, marks, trailingPt = 0 }) => (
  <Grid id={id} grid={barGrid(width, heightPt, insetPt, track, marks, trailingPt)} />
);

/** A laid-out bar grid, as a table. */
const Grid: Nodes<{ id: string; grid: BarGrid }> = ({ id, grid }) => (
  <Table id={id} variant="barNone" columns={grid.widths.map((width) => ({ width }))}>
    {grid.rows.map((row) => (
      <Row height={row.height} heightRule="exact">
        {row.cells.map((variant) => <Cell variant={variant}></Cell>)}
      </Row>
    ))}
  </Table>
);

/** Empty room of an exact depth, for the space above a table. */
export const Gap: Nodes<{ pt: 2 | 3 | 4 | 6 | 8 | 10 | 12 | 14 | 16 | 20 | 24 | 28 | 34 | 45 }> = ({ pt }) => (
  <Paragraph variant={`gap${pt}`} />
);

/** One line of a bar list: what it is, its bar, and its figure. */
export interface BarLine {
  label: string;
  value: string;
  marks: Mark[];
  /** The label's and the figure's variants, when this line is set apart. */
  labelVariant?: string;
  valueVariant?: string;
}

/**
 * Labelled bars down the page: label, bar, figure, a fixed gap between lines.
 *
 * One table, the gaps as rows of an exact depth: separate tables side by side
 * are joined into one by Word anyway, and a row is a gap both engines agree on.
 */
export const BarList: Nodes<{
  id: string;
  lines: BarLine[];
  labelWidth: number;
  valueWidth: number;
  heightPt: number;
  insetPt?: number;
  gapPt: number;
  track?: string;
  gutter?: number;
}> = ({ id, lines, labelWidth, valueWidth, heightPt, insetPt, gapPt, track = "barTrack", gutter = 2.82 }) => {
  const barWidth = 178 - labelWidth - valueWidth - 2 * gutter;
  return (
    <Table
      id={id}
      variant="barList"
      columns={[{ width: labelWidth }, { width: gutter }, { width: barWidth }, { width: gutter }, { width: valueWidth, align: "right" }]}
    >
      {lines.map((line, index) => (
        <>
          {index > 0 && (
            <Row height={gapPt} heightRule="exact">
              <Cell variant="barNone"></Cell>
              <Cell variant="barNone"></Cell>
              <Cell variant="barNone"></Cell>
              <Cell variant="barNone"></Cell>
              <Cell variant="barNone"></Cell>
            </Row>
          )}
          <Row>
            <Cell variant={line.labelVariant ?? "barLabel"}>{line.label}</Cell>
            <Cell></Cell>
            <Cell>
              <Bar id={`${id}-${index}`} width={barWidth} heightPt={heightPt} insetPt={insetPt} track={track} marks={line.marks} trailingPt={1} />
            </Cell>
            <Cell></Cell>
            <Cell variant={line.valueVariant ?? "barValue"}>{line.value}</Cell>
          </Row>
        </>
      ))}
    </Table>
  );
};

/** A figure in a key-figures grid. */
export interface Figure {
  label: string;
  value: string;
  note: string;
  rise?: boolean;
  /** The cell's variant, when this figure is set apart from the rest. */
  cell?: string;
}

/**
 * Figures in a grid, each under a dark rule: a label, the figure, a note.
 *
 * The grid is one table, with its gutters as empty columns and the gap
 * between rows as an empty row of an exact depth, so the rules break where
 * the canvas breaks them.
 */
export const KeyFigures: Nodes<{
  id: string;
  figures: Figure[];
  columns: number;
  gutterPt: number;
  rowGapPt?: number;
  valueVariant?: string;
  cellVariant?: string;
  labelVariant?: string;
  noteVariant?: string;
}> = ({
  id,
  figures,
  columns,
  gutterPt,
  rowGapPt = 20,
  valueVariant = "keyValue",
  cellVariant = "keyFigure",
  labelVariant = "keyLabel",
  noteVariant = "keyNote",
}) => {
  const gutter = gutterPt * 0.35278;
  const width = (178 - gutter * (columns - 1)) / columns;
  const rows: Figure[][] = [];
  for (let i = 0; i < figures.length; i += columns) rows.push(figures.slice(i, i + columns));
  const grid = Array.from({ length: columns * 2 - 1 }, (_, i) => ({ width: i % 2 === 0 ? width : gutter }));

  return (
    <Table id={id} variant="keyGap" columns={grid}>
      {rows.map((row, index) => (
        <>
          {index > 0 && (
            <Row height={rowGapPt} heightRule="exact">
              {grid.map(() => <Cell variant="barNone"></Cell>)}
            </Row>
          )}
          <Row>
            {row.map((figure, column) => (
              <>
                {column > 0 && <Cell></Cell>}
                <Cell variant={figure.cell ?? cellVariant}>
                  <Paragraph variant={labelVariant}>{figure.label}</Paragraph>
                  <Paragraph variant={valueVariant}>{figure.value}</Paragraph>
                  <Paragraph variant={figure.rise ? "keyRise" : noteVariant}>
                    {figure.rise ? `▲ ${figure.note}` : figure.note}
                  </Paragraph>
                </Cell>
              </>
            ))}
          </Row>
        </>
      ))}
    </Table>
  );
};

/**
 * Upright bars standing side by side on a common foot, as table cells.
 *
 * Each slot is a column, in millimetres, holding a bar or empty. Each band
 * between two bar heights is a row, so a taller bar is filled above a shorter
 * one's top and both are filled below it; `hanging` turns the bands over for
 * bars that drop from a baseline instead.
 */
export const Uprights: Nodes<{ id: string; slots: { width: number; bar?: Upright }[]; hanging?: boolean }> = (
  { id, slots, hanging = false },
) => <Grid id={id} grid={uprightGrid(slots, hanging)} />;
