import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { FinnishVirtualBarcodeLines } from "./lines.ts";

/**
 * The IBAN and reference as label/value rows, so their values line up with
 * other label/value blocks on the invoice. The code follows under the table.
 */
export const table = (lines: FinnishVirtualBarcodeLines): Yield => (
  <>
    {(lines.iban || lines.reference) && (
      <Table id="virtual-barcode-details" columns={[{ width: 45 }, { width: "auto" }]}>
        {lines.iban && (
          <Row id="virtual-barcode-iban">
            <Cell>Saajan tilinumero (IBAN)</Cell>
            <Cell>{lines.iban}</Cell>
          </Row>
        )}
        {lines.reference && (
          <Row id="virtual-barcode-reference">
            <Cell>Viitenumero</Cell>
            <Cell>{lines.reference}</Cell>
          </Row>
        )}
      </Table>
    )}
    {/* Not a row: beside a 45mm label the value cell is about 115mm less cell padding, too narrow for 54 digits. */}
    {lines.code && <Paragraph id="virtual-barcode-label">Virtuaaliviivakoodi</Paragraph>}
    {lines.code && <Paragraph id="virtual-barcode-code">{lines.code}</Paragraph>}
  </>
);
