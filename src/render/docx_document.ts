import { Document, HeadingLevel, Packer, PageOrientation, Paragraph, TextRun } from "docx";
import type { DocumentModel, DocumentNode, DocumentStyle } from "../domain/types.ts";
import { cleanMinimalDocumentStyle } from "../project/style.ts";

export function createDocxDocument(doc: DocumentModel): Document {
  const style = doc.style ?? cleanMinimalDocumentStyle;

  return new Document({
    styles: createDocxStyles(style),
    sections: [
      {
        properties: {
          page: {
            size: {
              width: pageWidthTwips(style),
              height: pageHeightTwips(style),
              orientation: style.page.orientation === "landscape"
                ? PageOrientation.LANDSCAPE
                : PageOrientation.PORTRAIT,
            },
            margin: {
              top: mmToTwips(style.page.margins.topMm),
              right: mmToTwips(style.page.margins.rightMm),
              bottom: mmToTwips(style.page.margins.bottomMm),
              left: mmToTwips(style.page.margins.leftMm),
            },
          },
        },
        children: [
          new Paragraph({
            text: doc.title,
            heading: HeadingLevel.TITLE,
          }),
          ...doc.nodes.flatMap((node) => renderNode(node, style)),
        ],
      },
    ],
  });
}

export async function createDocxBlob(doc: DocumentModel): Promise<Blob> {
  return await Packer.toBlob(createDocxDocument(doc));
}

function renderNode(node: DocumentNode, style: DocumentStyle): Paragraph[] {
  if (node.kind === "section") {
    return [
      new Paragraph({
        text: node.title ?? node.id,
        heading: HeadingLevel.HEADING_1,
      }),
      ...node.children.flatMap((child) => renderNode(child, style)),
    ];
  }

  if (node.kind === "paragraph") {
    return [
      new Paragraph({
        children: [new TextRun(node.text ?? "")],
        spacing: {
          after: ptToTwips(style.paragraph.spacingAfterPt),
          line: Math.round(style.typography.bodyLineHeight * 240),
        },
      }),
    ];
  }

  if (node.kind === "image") {
    return [
      new Paragraph({
        children: [new TextRun(`[image: ${node.alt ?? node.path ?? node.placeholder ?? node.id}]`)],
      }),
    ];
  }

  if (node.kind === "graph") {
    return [
      new Paragraph({
        children: [
          new TextRun(`[${node.graphType} graph: ${node.caption ?? node.placeholder ?? node.id}]`),
        ],
      }),
    ];
  }

  // A built document has already walked its loops, so this is a published one
  // being packed directly. The body is all there is to show: one pass, standing
  // for however many the request will ask for.
  if (node.kind === "repeat") {
    return node.children.flatMap((child) => renderNode(child, style));
  }

  return [
    new Paragraph({
      children: [new TextRun(node.title ?? "Table of contents")],
    }),
  ];
}

function createDocxStyles(style: DocumentStyle) {
  return {
    default: {
      document: {
        run: {
          font: style.typography.bodyFont,
          size: ptToHalfPoints(style.typography.bodySizePt),
          color: style.typography.color,
        },
        paragraph: {
          spacing: {
            after: ptToTwips(style.paragraph.spacingAfterPt),
            line: Math.round(style.typography.bodyLineHeight * 240),
          },
        },
      },
      title: {
        run: {
          font: style.typography.headingFont,
          size: ptToHalfPoints(style.title.fontSizePt),
          bold: style.title.weight === "bold",
          color: style.typography.color,
        },
        paragraph: {
          spacing: {
            before: ptToTwips(style.title.spacingBeforePt),
            after: ptToTwips(style.title.spacingAfterPt),
          },
        },
      },
      heading1: {
        run: {
          font: style.typography.headingFont,
          size: ptToHalfPoints(style.sectionHeading.fontSizePt),
          bold: style.sectionHeading.weight === "bold",
          color: style.typography.color,
        },
        paragraph: {
          spacing: {
            before: ptToTwips(style.sectionHeading.spacingBeforePt),
            after: ptToTwips(style.sectionHeading.spacingAfterPt),
          },
        },
      },
    },
  };
}

function pageWidthTwips(style: DocumentStyle): number {
  return mmToTwips(style.page.size === "A4" ? 210 : 215.9);
}

function pageHeightTwips(style: DocumentStyle): number {
  return mmToTwips(style.page.size === "A4" ? 297 : 279.4);
}

function mmToTwips(value: number): number {
  return Math.round(value * 56.6929133858);
}

function ptToTwips(value: number): number {
  return Math.round(value * 20);
}

function ptToHalfPoints(value: number): number {
  return Math.round(value * 2);
}
