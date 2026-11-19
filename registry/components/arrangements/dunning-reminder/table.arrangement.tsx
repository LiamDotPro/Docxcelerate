import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { DunningReminderLines } from "./lines.ts";

/**
 * Number, due date, days overdue and amount in columns, with a total row. A
 * reminder covering several invoices is read by scanning amounts down one
 * column, which is why this is the default.
 */
export const table = (lines: DunningReminderLines): Yield => (
  <>
    <Table
      id="dunning-reminder-invoices"
      columns={[{ width: "auto" }, { width: 32 }, { width: 24 }, { width: 32, align: "right" }]}
    >
      <Row id="dunning-reminder-invoices-header" header>
        <Cell>Invoice</Cell>
        <Cell>Due date</Cell>
        <Cell>Overdue</Cell>
        <Cell>Amount</Cell>
      </Row>
    {lines.invoices[0].number && (
      <Row id="dunning-reminder-invoice-1">
        <Cell>{lines.invoices[0].number}</Cell>
        <Cell>{lines.invoices[0].dueDate}</Cell>
        <Cell>{lines.invoices[0].overdue}</Cell>
        <Cell>{lines.invoices[0].amount}</Cell>
      </Row>
    )}
    {lines.invoices[1].number && (
      <Row id="dunning-reminder-invoice-2">
        <Cell>{lines.invoices[1].number}</Cell>
        <Cell>{lines.invoices[1].dueDate}</Cell>
        <Cell>{lines.invoices[1].overdue}</Cell>
        <Cell>{lines.invoices[1].amount}</Cell>
      </Row>
    )}
    {lines.invoices[2].number && (
      <Row id="dunning-reminder-invoice-3">
        <Cell>{lines.invoices[2].number}</Cell>
        <Cell>{lines.invoices[2].dueDate}</Cell>
        <Cell>{lines.invoices[2].overdue}</Cell>
        <Cell>{lines.invoices[2].amount}</Cell>
      </Row>
    )}
    {lines.invoices[3].number && (
      <Row id="dunning-reminder-invoice-4">
        <Cell>{lines.invoices[3].number}</Cell>
        <Cell>{lines.invoices[3].dueDate}</Cell>
        <Cell>{lines.invoices[3].overdue}</Cell>
        <Cell>{lines.invoices[3].amount}</Cell>
      </Row>
    )}
    {lines.invoices[4].number && (
      <Row id="dunning-reminder-invoice-5">
        <Cell>{lines.invoices[4].number}</Cell>
        <Cell>{lines.invoices[4].dueDate}</Cell>
        <Cell>{lines.invoices[4].overdue}</Cell>
        <Cell>{lines.invoices[4].amount}</Cell>
      </Row>
    )}
    {lines.invoices[5].number && (
      <Row id="dunning-reminder-invoice-6">
        <Cell>{lines.invoices[5].number}</Cell>
        <Cell>{lines.invoices[5].dueDate}</Cell>
        <Cell>{lines.invoices[5].overdue}</Cell>
        <Cell>{lines.invoices[5].amount}</Cell>
      </Row>
    )}
    {lines.invoices[6].number && (
      <Row id="dunning-reminder-invoice-7">
        <Cell>{lines.invoices[6].number}</Cell>
        <Cell>{lines.invoices[6].dueDate}</Cell>
        <Cell>{lines.invoices[6].overdue}</Cell>
        <Cell>{lines.invoices[6].amount}</Cell>
      </Row>
    )}
    {lines.invoices[7].number && (
      <Row id="dunning-reminder-invoice-8">
        <Cell>{lines.invoices[7].number}</Cell>
        <Cell>{lines.invoices[7].dueDate}</Cell>
        <Cell>{lines.invoices[7].overdue}</Cell>
        <Cell>{lines.invoices[7].amount}</Cell>
      </Row>
    )}
      <Row id="dunning-reminder-total">
        <Cell>Total overdue</Cell>
        <Cell></Cell>
        <Cell></Cell>
        <Cell>{lines.total}</Cell>
      </Row>
    </Table>
    {lines.overflow && <Paragraph id="dunning-reminder-overflow">{lines.overflow}</Paragraph>}
  </>
);
