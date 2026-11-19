/**
 * What the engine writes for the preview client, node by node.
 *
 * The site shows the document twice: as a proof, with placeholders where the
 * engine writes, and as the engine would write it for Ms Nkemelu. This is the
 * second — the canvas's own ENGINE text, word for word, handed back through a
 * fixture AI client so the page is the real build path with a deterministic
 * model. A node missing here fails the build rather than printing a gap.
 */
const LETTER = [
  "Your portfolio earned £12,150 this quarter, a return of 3.0%, and ended September worth £412,860. That is ahead of its benchmark, which returned 2.6%, and most of the difference came in August, when global and emerging-market shares rose together.",
  "A good quarter for shares has a side effect. Equities now make up 64.2% of your portfolio against the 60% we agreed, outside the range we set to keep your risk where you chose it. I recommend selling £17,300 of the Global Equity Index Fund and moving it into bonds and property. Page 3 sets out the trades.",
  "Your plans have not changed since we spoke in July, so neither has your strategy. The £12,000 you asked for in January is already held in cash, and nothing in this recommendation touches it. You also have £8,400 of this year's ISA allowance left, which we can talk about at your review.",
  "On the projections on page 13 you remain on track to retire in 2036 with the income you told us you need. If anything in your circumstances changes before January, please let me know, so that my advice keeps up with it.",
];

const HIGHLIGHTS = [
  "Your portfolio returned 3.0% against 2.6% for its benchmark, led by global and emerging-market shares.",
  "Equities have drifted 4.2 points above target; we recommend moving £17,300 into bonds and property.",
  "You remain on track for retirement in 2036, and £8,400 of your ISA allowance is still unused.",
];

const REBALANCE = "Rebalance";
const RECOMMENDED = "recommended this quarter";

export const engineText: Record<string, string> = {
  "cover-decision": REBALANCE,
  "cover-decision-note": RECOMMENDED,
  "cover-headline":
    "A strong quarter, led by global and emerging-market shares. Equities have run ahead of your agreed mix, so we recommend a modest rebalance.",
  letter: LETTER.join("\n\n"),
  decision: REBALANCE,
  "decision-note": RECOMMENDED,
  rebalance: "Equities are 4.2 points above target, outside their agreed range.",
  "rebalance-why":
    "A strong quarter for shares moved the portfolio away from the risk level you chose. Moving the excess into bonds and property restores it, without touching the cash set aside for January.",
  highlights: HIGHLIGHTS.join("\n\n"),
  "performance-caption":
    "Over twelve months the portfolio returned 9.4% against 8.7% for its benchmark. More than half of that lead was built this quarter, when global and emerging-market equities rose together.",
  "return-drivers":
    "Equities did the work, adding 2.6 percentage points between them; the Global Equity Index and Emerging Markets funds contributed more than half of the total on their own. Bonds added 0.3 as expectations of a rate cut firmed. Smaller companies and property were the only drags, and together cost about 0.1.",
  drift:
    "Equities finished the quarter 4.2 points above target and bonds 3.9 below, both outside the three-point ranges we agreed. Property and cash are within range. The proposed trades bring every class back to target.",
  "global-equity-note": "Lifted by US technology earnings in August; the largest single contributor.",
  "uk-income-note": "Steady dividends from banks and utilities, and £1,104 of the quarter's income.",
  "emerging-note": "The best performer, helped by a weaker dollar and recovering demand in Asia.",
  "smaller-cos-note": "The only equity fund to fall, as higher borrowing costs weighed on smaller firms.",
  "global-bond-note": "Gained as markets priced in a rate cut; the hedge removed the currency swing.",
  "gilt-2031-note": "Held for stability. The price barely moved, as intended.",
  "property-note": "Slipped on office valuations, partly offset by logistics and housing.",
  "cash-note": "Holds your January withdrawal of £12,000 and the next two quarters' fees.",
  "global-equity-why":
    "The core of your equity allocation: broad, low-cost exposure to about 4,000 companies. Its weight in large US technology firms explains most of this quarter's gain, and is why the rebalance trims it first.",
  "global-bond-why":
    "It steadies the portfolio when shares fall, and the currency hedge keeps its returns in pounds. It is where most of the rebalance goes.",
  "uk-income-why":
    "Your main source of income. Its dividends paid £1,104 this quarter, nearly half of the £2,316 sent to your account.",
  "smaller-cos-why":
    "The one fund that fell this quarter. Smaller firms borrow more and feel higher rates first; we hold it for the long run and are not adding to it now.",
  "emerging-why":
    "The quarter's best performer, helped by a weaker dollar. It is also the most volatile part of your portfolio, which is why it stays under a tenth of it.",
  "gilt-2031-why":
    "It pays a fixed 4.25% and returns your capital in July 2031, five years before you plan to retire. It is there to be dull, and it is.",
  "income-summary":
    "Your investments paid £8,857 over the last twelve months, about 2.1% of the portfolio's value. This quarter's £2,316 was the strongest of the year, helped by interim dividends from the UK Equity Income Fund.",
  "plan-decision": "On track",
  "plan-decision-note": "for retirement in 2036",
  "plan-check":
    "On the central projection, your portfolio reaches about £827k by 2036, above the £750k you told us you need. Only the weak projection falls short, by about £106k, so we see no reason to change your contributions or your retirement date.",
};
