import { Paragraph, type Yield } from "docxcelerate/template";
import type { ExemptionCertificateCheckLines } from "./lines.ts";

/**
 * One paragraph per line, for a narrow column or reasons long enough to wrap
 * badly in a table cell.
 */
export const stacked = (lines: ExemptionCertificateCheckLines): Yield => (
  <>
    <Paragraph id="exemption-check-summary">{lines.summary}</Paragraph>
    {lines.rows[0].description && (
      <Paragraph id="exemption-check-line-1">
        {lines.rows[0].description}, {lines.rows[0].amount}: {lines.rows[0].status}. {lines.rows[0].reason}.
      </Paragraph>
    )}
    {lines.rows[1].description && (
      <Paragraph id="exemption-check-line-2">
        {lines.rows[1].description}, {lines.rows[1].amount}: {lines.rows[1].status}. {lines.rows[1].reason}.
      </Paragraph>
    )}
    {lines.rows[2].description && (
      <Paragraph id="exemption-check-line-3">
        {lines.rows[2].description}, {lines.rows[2].amount}: {lines.rows[2].status}. {lines.rows[2].reason}.
      </Paragraph>
    )}
    {lines.rows[3].description && (
      <Paragraph id="exemption-check-line-4">
        {lines.rows[3].description}, {lines.rows[3].amount}: {lines.rows[3].status}. {lines.rows[3].reason}.
      </Paragraph>
    )}
    {lines.rows[4].description && (
      <Paragraph id="exemption-check-line-5">
        {lines.rows[4].description}, {lines.rows[4].amount}: {lines.rows[4].status}. {lines.rows[4].reason}.
      </Paragraph>
    )}
    {lines.rows[5].description && (
      <Paragraph id="exemption-check-line-6">
        {lines.rows[5].description}, {lines.rows[5].amount}: {lines.rows[5].status}. {lines.rows[5].reason}.
      </Paragraph>
    )}
    {lines.rows[6].description && (
      <Paragraph id="exemption-check-line-7">
        {lines.rows[6].description}, {lines.rows[6].amount}: {lines.rows[6].status}. {lines.rows[6].reason}.
      </Paragraph>
    )}
    {lines.rows[7].description && (
      <Paragraph id="exemption-check-line-8">
        {lines.rows[7].description}, {lines.rows[7].amount}: {lines.rows[7].status}. {lines.rows[7].reason}.
      </Paragraph>
    )}
    {lines.rows[8].description && (
      <Paragraph id="exemption-check-line-9">
        {lines.rows[8].description}, {lines.rows[8].amount}: {lines.rows[8].status}. {lines.rows[8].reason}.
      </Paragraph>
    )}
    {lines.rows[9].description && (
      <Paragraph id="exemption-check-line-10">
        {lines.rows[9].description}, {lines.rows[9].amount}: {lines.rows[9].status}. {lines.rows[9].reason}.
      </Paragraph>
    )}
    {lines.rows[10].description && (
      <Paragraph id="exemption-check-line-11">
        {lines.rows[10].description}, {lines.rows[10].amount}: {lines.rows[10].status}. {lines.rows[10].reason}.
      </Paragraph>
    )}
    {lines.rows[11].description && (
      <Paragraph id="exemption-check-line-12">
        {lines.rows[11].description}, {lines.rows[11].amount}: {lines.rows[11].status}. {lines.rows[11].reason}.
      </Paragraph>
    )}
    {lines.rows[12].description && (
      <Paragraph id="exemption-check-line-13">
        {lines.rows[12].description}, {lines.rows[12].amount}: {lines.rows[12].status}. {lines.rows[12].reason}.
      </Paragraph>
    )}
    {lines.rows[13].description && (
      <Paragraph id="exemption-check-line-14">
        {lines.rows[13].description}, {lines.rows[13].amount}: {lines.rows[13].status}. {lines.rows[13].reason}.
      </Paragraph>
    )}
    {lines.rows[14].description && (
      <Paragraph id="exemption-check-line-15">
        {lines.rows[14].description}, {lines.rows[14].amount}: {lines.rows[14].status}. {lines.rows[14].reason}.
      </Paragraph>
    )}
    {lines.rows[15].description && (
      <Paragraph id="exemption-check-line-16">
        {lines.rows[15].description}, {lines.rows[15].amount}: {lines.rows[15].status}. {lines.rows[15].reason}.
      </Paragraph>
    )}
    <Paragraph id="exemption-check-basis">{lines.basis}</Paragraph>
  </>
);
