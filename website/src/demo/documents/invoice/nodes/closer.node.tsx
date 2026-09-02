import {
  Cell,
  type Nodes,
  Paragraph,
  Row,
  Table,
  useDeriver,
  useState,
} from "docxcelerate/template";
import { invoiceDates } from "../derivers.ts";
import type { InvoiceData } from "../types.ts";

/**
 * What is left to say once the total is settled: where the rest of it lives.
 *
 * An invoice that turns the page has to say so on the page it turns from,
 * or the reader takes the total for the end of the document and files it
 * without the account details. The due date sits opposite, because that is
 * the other thing someone reads off the bottom of page one.
 */
export const Closer: Nodes = async () => {
  const [state] = useState((data: InvoiceData) => ({
    issueDate: data.issueDate,
    dueDate: data.dueDate,
  }));
  const dates = await useDeriver(invoiceDates, [state.issueDate, state.dueDate]);

  return (
    <Table id="closer" columns={[{ width: "auto" }, { width: 46, align: "right" }]}>
      <Row>
        <Cell id="closer-note">
          <Paragraph variant="muted">
            Payment details, terms and a scan-to-pay code are on page 2.
          </Paragraph>
        </Cell>
        <Cell id="closer-due">
          <Paragraph variant="label">Due {dates.due}</Paragraph>
        </Cell>
      </Row>
    </Table>
  );
};
