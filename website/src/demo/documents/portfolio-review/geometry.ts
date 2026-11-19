/**
 * The arithmetic behind the drawn bars, kept out of the components.
 *
 * A component's conditionals are compiled into decisions a published document
 * can carry; these are not decisions about a client, they are where a cell
 * edge falls. Written here they stay ordinary JavaScript.
 */

/** One thing drawn on a bar, in millimetres from its left end. */
export interface Mark {
  from: number;
  to: number;
  variant: string;
  /** Lifted off the track's top and bottom, like a value bar inside its track. */
  inset?: boolean;
}

/** A bar laid out as table rows: each row's depth, and the fill of each column. */
export interface BarGrid {
  widths: number[];
  rows: { height: number; cells: string[] }[];
}

/**
 * Every mark's ends become column edges, and each row paints each stretch
 * with whatever mark is on top there: three rows when something is inset
 * (above it, through it, below it), one when nothing is.
 */
export function barGrid(
  width: number,
  heightPt: number,
  insetPt: number,
  track: string,
  marks: Mark[],
  trailingPt = 0,
): BarGrid {
  const clamp = (x: number) => Math.min(width, Math.max(0, x));
  const edges = [...new Set([0, width, ...marks.flatMap((m) => [clamp(m.from), clamp(m.to)])])]
    .map((x) => Math.round(x * 100) / 100)
    .sort((a, b) => a - b)
    .filter((x, i, all) => i === 0 || x - all[i - 1]! > 0.05);
  const spans = edges.slice(1).map((to, i) => ({ from: edges[i]!, to }));
  const paint = (middle: boolean) =>
    spans.map(({ from, to }) => {
      const centre = (from + to) / 2;
      const top = marks.filter((m) => (middle || !m.inset) && clamp(m.from) <= centre && centre <= clamp(m.to)).pop();
      return top ? top.variant : track;
    });
  const inset = marks.some((m) => m.inset) && insetPt > 0;
  const rows = inset
    ? [
      { height: insetPt, cells: paint(false) },
      { height: heightPt - 2 * insetPt, cells: paint(true) },
      { height: insetPt, cells: paint(false) },
    ]
    : [{ height: heightPt, cells: paint(true) }];
  // What follows the bar inside its cell — the point-deep paragraph Word puts
  // after any table in a cell — comes off the bar's last row, so the cell is
  // as deep as the bar the design draws and no deeper.
  rows[rows.length - 1]!.height -= Math.min(trailingPt, rows[rows.length - 1]!.height - 0.5);
  return { widths: spans.map(({ from, to }) => to - from), rows };
}

/** One upright bar: how tall, in points, and its fill. */
export interface Upright {
  heightPt: number;
  variant: string;
}

/**
 * Upright bars as rows of bands: each band between two bar heights is a row,
 * and a bar fills every band at or below its own top. `hanging` turns the
 * bands over for bars that drop from a baseline.
 */
export function uprightGrid(slots: { width: number; bar?: Upright }[], hanging: boolean): BarGrid {
  const levels = [...new Set(slots.map((slot) => (slot.bar ? slot.bar.heightPt : 0)))]
    .filter((h) => h > 0)
    .sort((a, b) => b - a);
  const bands = levels.map((top, i) => ({ top, height: top - (levels[i + 1] ?? 0) }));
  const ordered = hanging ? [...bands].reverse() : bands;
  return {
    widths: slots.map((slot) => slot.width),
    rows: ordered.map((band) => ({
      height: band.height,
      cells: slots.map((slot) => (slot.bar && slot.bar.heightPt >= band.top ? slot.bar.variant : "barNone")),
    })),
  };
}
