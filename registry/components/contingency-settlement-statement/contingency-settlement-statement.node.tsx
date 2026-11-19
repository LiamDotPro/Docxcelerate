import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import contingencySettlementStatement from "../derivers/contingency-settlement-statement.deriver.ts";

/**
 * The closing statement a contingent or damages-based fee requires (ABA Model
 * Rule 1.5(c); the UK DBA and CFA equivalents): the gross recovery, each
 * deduction in the order the agreement and the law apply it, and the net to the
 * client, as signed amounts in one column that adds up.
 *
 * Decisions already made for you:
 *
 * - Tiered percentages are marginal by band. The top rate is never applied to
 *   the whole recovery, and the fee is the sum of the rounded band charges, so
 *   the working in the method sentence adds up exactly.
 * - The fee base follows the agreement: gross, or the recovery net of expenses.
 *   When expenses exist the statement also shows what the other base would have
 *   charged.
 * - The rate comes from the stage the case reached, read from the data and never
 *   guessed. A missing stage throws when the schedule or the California cap needs one.
 * - A statutory cap (NY Judiciary Law 474-a, CA B&P 6146, the UK DBA Regulations
 *   2013, the CFA Order 2013) is computed as an amount, compared with the agreed
 *   fee, and the lower is charged. The DBA cap counts VAT and counsel's fees and
 *   leaves out future loss; costs paid by the other side are credited, never
 *   beyond what the client owes. The cap is chosen in the data, never inferred.
 * - Liens come off the client's share after the fee. Only a lien marked
 *   `beforeFee` comes off first and lowers the fee base.
 * - A recovery too small to cover everything prints as a shortfall, with the
 *   agreement's own words on who carries it, or an "incomplete" warning when
 *   those words were not given.
 *
 * All of it is decided in one deriver, because every step is arithmetic,
 * sorting or formatting on request data, and publishing refuses all of those.
 * The node computes nothing: every test is written inline so a publish build
 * keeps each arm.
 *
 * The lines are sixteen fixed slots rather than a `.map()`. The resolver reads a
 * published loop's collection from request data and ignores the derived scope,
 * so a `.map()` over the deriver's rows would publish and then resolve to no
 * rows; conditions and `{{derived.…}}` tokens do honour the derived scope.
 *
 * The node publishes 20 branches against the default `branchLimit` of 32: 16 row
 * slots, other-base, cap and two shortfall arms.
 *
 * Installed by `dxcl add contingency-settlement-statement`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface ContingencySettlementStatementData {
  settlement: {
    /**
     * Which kind of agreement the fee is paid under.
     * `contingency`: a percentage of the recovery (US-style contingent fee).
     * `DBA`: an England & Wales damages-based agreement (DBA Regulations 2013).
     * `CFA`: an England & Wales conditional fee agreement: base costs plus a
     * success fee. Required: anything else throws.
     */
    regime: "contingency" | "DBA" | "CFA";
    /**
     * ISO 4217 code every amount is in: "USD", "GBP". It also sets the minor
     * unit (two decimals for USD and GBP, none for JPY). Required: throws when
     * missing or not a code Intl knows.
     */
    currency: string;
    /**
     * The total awarded or agreed, in major units, before anything is taken
     * off it. That includes any benefits the compensator withholds for the
     * CRU; list those as a lien with `cru: true`. Required: throws when missing,
     * not finite, or negative. 0 is allowed and prints a shortfall if anything
     * is owed.
     */
    grossRecovery: number;
    /**
     * How far the case got, which picks the percentage. Missing is allowed only
     * when no schedule row names a stage, and when `statutoryCap` is not
     * "ca-medical-negligence". Otherwise it throws: the stage is never guessed.
     */
    stage?: "pre-suit" | "filed" | "trial" | "appeal";
    /**
     * The agreement's percentages. For `contingency` and `DBA`, each row is a
     * band of the fee base. For `CFA`, each row is the success-fee uplift on
     * base costs. Required, non-empty.
     */
    schedule: Array<{
      /**
       * The stage this row applies at. Rows for the given `stage` are used.
       * When none match, rows with no stage are used. When none of those
       * exist either, it throws.
       */
      stage?: "pre-suit" | "filed" | "trial" | "appeal";
      /**
       * Upper edge of this band, cumulative, in major units: 250000 means
       * "the first 250,000". Missing means the open top band ("everything
       * above"). Among the rows used, exactly one must have no `upTo`, and the
       * `upTo` values must be positive and distinct. Otherwise it throws.
       * CFA: must be missing on every row, otherwise it throws.
       */
      upTo?: number;
      /**
       * The rate, as a percentage from 0 to 100. A third is `100 / 3` and
       * prints as 33.33%. For `DBA` it is the solicitor's charge before VAT.
       * For `CFA` it is the uplift on base costs and may not exceed 100.
       * Out of range throws.
       */
      percent: number;
    }>;
    /**
     * What the agreement applies the percentage to. "gross" is the whole
     * recovery. "net" is the recovery after the case expenses below. Required
     * for `contingency` and `DBA`, where missing throws. Ignored for `CFA`.
     */
    base?: "gross" | "net";
    /**
     * A statutory fee cap that applies to this matter. `contingency` only:
     * given with another regime, it throws. Missing means no statutory cap.
     * The cap is chosen here and is never inferred from a jurisdiction.
     */
    statutoryCap?: "ny-medical-malpractice" | "ca-medical-negligence";
    /**
     * Which DBA or CFA cap applies. Required for `DBA` and `CFA`, where
     * missing throws. Ignored for `contingency`.
     */
    claimType?: "personal-injury" | "employment" | "other";
    /**
     * The personal-injury heads of damage, in major units. Required when
     * `claimType` is "personal-injury" under `DBA` or `CFA`, where missing
     * throws. Ignored otherwise.
     */
    damagesBreakdown?: {
      /** General damages for pain, suffering and loss of amenity. */
      general: number;
      /** Past (pre-settlement) financial loss. */
      pastLoss: number;
      /**
       * Future financial loss. It is never in a cap base. Printed in the cap
       * sentence as excluded when above 0. Missing means 0.
       */
      futureLoss?: number;
    };
    /**
     * CFA base costs (hourly charges), before VAT, in major units. Required for
     * `CFA`, where missing throws. Ignored otherwise.
     */
    baseCosts?: number;
    /**
     * VAT on the solicitor's charges, as a percentage: 20. `DBA` and `CFA` only.
     * Missing means 0, and no VAT row is printed.
     */
    vatPercent?: number;
    /**
     * Counsel's fees, VAT included, in major units. `DBA` and `CFA` only. Under
     * a DBA they are part of the capped payment. Missing means 0, and no row
     * is printed.
     */
    counselFees?: number;
    /**
     * Costs the opponent has paid or must pay, VAT included, in major units.
     * `DBA` and `CFA` only. They are credited against what the client owes and
     * never credited beyond it. Missing means 0, and no row is printed.
     */
    recoverableCostsPaid?: number;
    /**
     * Case expenses and disbursements the client reimburses. Under a DBA they
     * sit outside the cap. Each is printed as a row with its label as given.
     * Missing or empty means none.
     */
    expenses?: Array<{ label: string; amount: number }>;
    /**
     * Liens and subrogation claims paid from the recovery: medical, Medicare,
     * health plan, workers' compensation, CRU. Missing or empty means none.
     */
    liens?: Array<{
      /** Who holds it, printed as given. Required, non-blank. */
      holder: string;
      /** The amount claimed, in major units. */
      amount: number;
      /**
       * The amount actually payable after a negotiated or statutory reduction,
       * such as Medicare's procurement-cost reduction or a hospital compromise.
       * Used instead of `amount`, and the row says "reduced from <amount>".
       * Throws when greater than `amount`. 0 is allowed (waived).
       */
      reducedAmount?: number;
      /**
       * True only where the agreement or the lien statute says the lien comes
       * off before the fee is worked out. It then reduces the fee base. Missing
       * means false: paid from the client's share after the fee.
       */
      beforeFee?: boolean;
      /**
       * True for benefits recoverable by the Compensation Recovery Unit. In a
       * UK PI claim they are netted off past loss, never general damages, for
       * the DBA and CFA caps. Missing means false.
       */
      cru?: boolean;
    }>;
    /**
     * The agreement's words on who carries a shortfall, printed verbatim after
     * the shortfall sentence. Read only when there is a shortfall. Missing or
     * empty there prints the "incomplete" warning instead.
     */
    shortfallTerms?: string;
  };
}

/**
 * Prints the settlement table, the method sentence, and whichever of the
 * other-base, cap and shortfall paragraphs apply.
 */
export const ContingencySettlementStatement: Nodes = async () => {
  const [settlement] = useState((d: ContingencySettlementStatementData) => d.settlement);
  const s = await useDeriver(contingencySettlementStatement, [settlement]);

  return (
    <Section id="contingency-settlement" title="Settlement statement">
      <Table
        id="contingency-settlement-table"
        columns={[{ width: "auto" }, { width: 38, align: "right" }]}
      >
        <Row id="contingency-settlement-header" header>
          <Cell>Item</Cell>
          <Cell>Amount</Cell>
        </Row>
        <Row id="contingency-settlement-gross">
          <Cell>Gross recovery</Cell>
          <Cell>{s.gross}</Cell>
        </Row>
        {s.rows[0].label && (
          <Row id="contingency-settlement-line-1">
            <Cell>{s.rows[0].label}</Cell>
            <Cell>{s.rows[0].amount}</Cell>
          </Row>
        )}
        {s.rows[1].label && (
          <Row id="contingency-settlement-line-2">
            <Cell>{s.rows[1].label}</Cell>
            <Cell>{s.rows[1].amount}</Cell>
          </Row>
        )}
        {s.rows[2].label && (
          <Row id="contingency-settlement-line-3">
            <Cell>{s.rows[2].label}</Cell>
            <Cell>{s.rows[2].amount}</Cell>
          </Row>
        )}
        {s.rows[3].label && (
          <Row id="contingency-settlement-line-4">
            <Cell>{s.rows[3].label}</Cell>
            <Cell>{s.rows[3].amount}</Cell>
          </Row>
        )}
        {s.rows[4].label && (
          <Row id="contingency-settlement-line-5">
            <Cell>{s.rows[4].label}</Cell>
            <Cell>{s.rows[4].amount}</Cell>
          </Row>
        )}
        {s.rows[5].label && (
          <Row id="contingency-settlement-line-6">
            <Cell>{s.rows[5].label}</Cell>
            <Cell>{s.rows[5].amount}</Cell>
          </Row>
        )}
        {s.rows[6].label && (
          <Row id="contingency-settlement-line-7">
            <Cell>{s.rows[6].label}</Cell>
            <Cell>{s.rows[6].amount}</Cell>
          </Row>
        )}
        {s.rows[7].label && (
          <Row id="contingency-settlement-line-8">
            <Cell>{s.rows[7].label}</Cell>
            <Cell>{s.rows[7].amount}</Cell>
          </Row>
        )}
        {s.rows[8].label && (
          <Row id="contingency-settlement-line-9">
            <Cell>{s.rows[8].label}</Cell>
            <Cell>{s.rows[8].amount}</Cell>
          </Row>
        )}
        {s.rows[9].label && (
          <Row id="contingency-settlement-line-10">
            <Cell>{s.rows[9].label}</Cell>
            <Cell>{s.rows[9].amount}</Cell>
          </Row>
        )}
        {s.rows[10].label && (
          <Row id="contingency-settlement-line-11">
            <Cell>{s.rows[10].label}</Cell>
            <Cell>{s.rows[10].amount}</Cell>
          </Row>
        )}
        {s.rows[11].label && (
          <Row id="contingency-settlement-line-12">
            <Cell>{s.rows[11].label}</Cell>
            <Cell>{s.rows[11].amount}</Cell>
          </Row>
        )}
        {s.rows[12].label && (
          <Row id="contingency-settlement-line-13">
            <Cell>{s.rows[12].label}</Cell>
            <Cell>{s.rows[12].amount}</Cell>
          </Row>
        )}
        {s.rows[13].label && (
          <Row id="contingency-settlement-line-14">
            <Cell>{s.rows[13].label}</Cell>
            <Cell>{s.rows[13].amount}</Cell>
          </Row>
        )}
        {s.rows[14].label && (
          <Row id="contingency-settlement-line-15">
            <Cell>{s.rows[14].label}</Cell>
            <Cell>{s.rows[14].amount}</Cell>
          </Row>
        )}
        {s.rows[15].label && (
          <Row id="contingency-settlement-line-16">
            <Cell>{s.rows[15].label}</Cell>
            <Cell>{s.rows[15].amount}</Cell>
          </Row>
        )}
        <Row id="contingency-settlement-net">
          <Cell>{s.netLabel}</Cell>
          <Cell>{s.net}</Cell>
        </Row>
      </Table>
      <Paragraph id="contingency-settlement-method">{s.method}</Paragraph>
      {s.otherBase && <Paragraph id="contingency-settlement-other-base">{s.otherBase}</Paragraph>}
      {s.cap && <Paragraph id="contingency-settlement-cap">{s.cap}</Paragraph>}
      {s.outcome === "shortfall" && settlement.shortfallTerms && (
        <Paragraph id="contingency-settlement-shortfall">
          The recovery does not cover the fee, expenses and liens, which leaves a shortfall of {s.net}.{" "}
          {settlement.shortfallTerms}
        </Paragraph>
      )}
      {s.outcome === "shortfall" && !settlement.shortfallTerms && (
        <Paragraph id="contingency-settlement-shortfall-incomplete">
          Statement incomplete: the recovery does not cover the fee, expenses and liens, which leaves a shortfall of{" "}
          {s.net}, and the agreement's terms on who carries a shortfall were not given. Add them before this statement
          is sent.
        </Paragraph>
      )}
    </Section>
  );
};
