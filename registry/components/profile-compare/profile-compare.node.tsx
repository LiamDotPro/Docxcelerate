import {
  Graph,
  Paragraph,
  Section,
  type Section as SectionComponent,
  useFormat,
  useState,
} from "docxcelerate/template";

/**
 * One subject against a benchmark across several measures, as a radar and as
 * the sentence naming where they differ most.
 *
 * A radar answers "strong where", not "how much" — and it only answers that if
 * every spoke shares a scale. **That is the decision this component exists to
 * make.** Handed a hundred pounds, four incidents and nine per cent, a radar
 * drawn straight from the numbers puts the hundred at the rim and the four at
 * the centre, and draws a shape that says nothing except which measure happens
 * to use the largest units. So each measure is scored against its own `best`
 * before it is plotted, the axis is scored out of 100, and the caption says so.
 *
 * Three more decisions come with it.
 *
 * **A reader cannot compare the areas of two rings, and is not asked to.**
 * The shape is the reading; the sentence carries the figures. Anyone who wants
 * to know by how much gets it in words, in the measure's own units, rather
 * than by eyeballing a polygon.
 *
 * **Fewer than three measures is not a radar.** Two spokes draw a line through
 * the middle, and one draws a spike. Below three the component prints the
 * comparison as prose and draws nothing.
 *
 * **A measure with no benchmark is still plotted.** The subject's ring is the
 * point; the benchmark is the context. Where a benchmark is missing the second
 * ring simply has a gap there, which is honest, rather than a zero, which
 * would read as a benchmark of nothing.
 *
 * Installed by `dxcl add profile-compare`.
 */

/** How many spokes a ring needs before it is a shape rather than a spike. */
const MIN_SPOKES = 3;

/** What a radar's axis runs to once every measure has been scored onto it. */
const SCALE = 100;

/** What this component reads. Add these fields to your document data type. */
export interface ProfileCompareData {
  profile: {
    /** What the first ring is called: a name, a site, a period. */
    subject: string;
    /** What the second is called: `"Sector average"`, `"Last year"`, `"Target"`. */
    benchmark: string;
    /**
     * One entry per spoke, in the order they should be read round the ring.
     *
     * Three at least. The order is the chart's order, so group related
     * measures next to one another — a radar's shape is read as a whole, and
     * neighbours are what make it one.
     */
    measures: Array<{
      /** What the measure is called, as it should be printed on the spoke. */
      label: string;
      /** The subject's reading, in whatever units the measure uses. */
      value: number;
      /** The benchmark's reading, in the same units. Absent draws a gap. */
      against?: number | null;
      /**
       * What counts as full marks for this measure.
       *
       * Every spoke is scored against its own `best` and plotted out of 100,
       * which is what lets measures in different units share one ring. Absent,
       * the larger of the two readings is used — which makes the chart
       * self-scaling and the outer point always the rim.
       */
      best?: number;
      /** What the measure is counted in, for the sentence: `"hours"`, `"%"`. */
      unit?: string;
    }>;
  };
}

export const ProfileCompare: SectionComponent = () => {
  const format = useFormat();
  const [profile] = useState((data: ProfileCompareData) => data.profile);

  const measures = (profile.measures ?? []).filter((measure) =>
    Number.isFinite(measure.value)
  );

  if (measures.length < MIN_SPOKES) {
    return (
      <Section id="profile-compare" title={`${profile.subject} against ${profile.benchmark}`}>
        <Paragraph id="profile-compare-none">
          {measures.length === 0
            ? "We have nothing recorded to compare yet."
            : `We are comparing ${format.list(measures.map((measure) => measure.label))}, ` +
              `which is too few to draw as a profile. ` +
              format.list(
                measures.map((measure) =>
                  `${measure.label} is ${format.number(measure.value)}${
                    measure.unit === undefined ? "" : ` ${measure.unit}`
                  }`
                ),
              ) + "."}
        </Paragraph>
      </Section>
    );
  }

  /** Full marks for one spoke: what it says, or whichever reading is larger. */
  const bestOf = (measure: (typeof measures)[number]) => {
    if (measure.best !== undefined && Number.isFinite(measure.best) && measure.best > 0) {
      return measure.best;
    }

    const against = Number.isFinite(measure.against) ? Number(measure.against) : 0;

    return Math.max(measure.value, against, 0);
  };

  /** One reading, scored onto the shared axis. */
  const score = (value: number | null | undefined, best: number) => {
    if (value === null || value === undefined || !Number.isFinite(value)) {
      return null;
    }

    // A spoke whose best is nothing has no scale to be scored against, and
    // dividing by it would put every reading on it at infinity.
    return best === 0 ? 0 : Math.min((value / best) * SCALE, SCALE);
  };

  const scored = measures.map((measure) => {
    const best = bestOf(measure);

    return {
      ...measure,
      mine: score(measure.value, best) ?? 0,
      theirs: score(measure.against, best),
    };
  });

  // Where the two rings differ most, measured on the shared scale — which is
  // the difference the shape actually shows. Ranking on the raw figures would
  // name whichever measure happens to use the largest units.
  const widest = [...scored]
    .filter((measure) => measure.theirs !== null)
    .sort((left, right) =>
      Math.abs(right.mine - (right.theirs ?? 0)) - Math.abs(left.mine - (left.theirs ?? 0))
    )[0];

  // A unit that is a symbol closes up against its figure — "82%", not "82 %".
  // A unit that is a word takes a space. Nothing else about the two differs,
  // and getting it wrong is visible in every sentence the component prints.
  const said = (value: number, unit: string | undefined) => {
    const gap = unit === undefined || /^[A-Za-z]/.test(unit) ? " " : "";

    return `${format.number(value)}${unit === undefined ? "" : gap + unit}`;
  };

  return (
    <Section id="profile-compare" title={`${profile.subject} against ${profile.benchmark}`}>
      <Paragraph id="profile-compare-summary">
        {widest === undefined
          ? `We have no ${profile.benchmark.toLowerCase()} to compare against, so the chart ` +
            `shows ${profile.subject} alone.`
          : `The widest gap is ${widest.label}: ${said(widest.value, widest.unit)} against ` +
            `${said(Number(widest.against), widest.unit)} for the ` +
            `${profile.benchmark.toLowerCase()}.`}
      </Paragraph>

      <Graph
        id="profile-compare-chart"
        title={`Scored out of ${SCALE}`}
        graphType="radar"
        legend="bottom"
        data={{
          categories: scored.map((measure) => measure.label),
          series: [
            { label: profile.subject, values: scored.map((measure) => measure.mine) },
            { label: profile.benchmark, values: scored.map((measure) => measure.theirs) },
          ],
        }}
        caption="Each measure is scored against its own best, so the spokes share a scale. Read the shape; the figures are above."
      />
    </Section>
  );
};
