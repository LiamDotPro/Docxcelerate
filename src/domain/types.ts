export type JsonObject = Record<string, unknown>;

export type NodeKind =
  | "section"
  | "paragraph"
  | "image"
  | "graph"
  | "tableOfContents"
  | "repeat";

export type NodeMode = "static" | "dynamic";

export type GraphType = "bar" | "line" | "pie";

export type DocumentStylePreset = "clean-minimal";

export type DocumentPageSize = "A4" | "LETTER";

export type DocumentPageOrientation = "portrait" | "landscape";

export type DocumentFontWeight = "regular" | "bold";

export interface DocumentPageMargins {
  topMm: number;
  rightMm: number;
  bottomMm: number;
  leftMm: number;
}

export interface DocumentPageStyle {
  size: DocumentPageSize;
  orientation: DocumentPageOrientation;
  margins: DocumentPageMargins;
}

export interface DocumentTypographyStyle {
  bodyFont: string;
  headingFont: string;
  bodySizePt: number;
  bodyLineHeight: number;
  color: string;
}

export interface DocumentParagraphStyle {
  spacingAfterPt: number;
}

export interface DocumentTextBlockStyle {
  fontSizePt: number;
  weight: DocumentFontWeight;
  spacingBeforePt: number;
  spacingAfterPt: number;
}

export interface DocumentStyle {
  preset: DocumentStylePreset;
  page: DocumentPageStyle;
  typography: DocumentTypographyStyle;
  paragraph: DocumentParagraphStyle;
  title: DocumentTextBlockStyle;
  sectionHeading: DocumentTextBlockStyle;
}

export type PromptKind = "general" | "info" | "negative" | "system";

export type ReferenceScope = "data" | "ctx" | "derived";

export interface DataReference {
  scope: ReferenceScope;
  path: string;
}

export type ComparisonOperator = "eq" | "ne" | "gt" | "gte" | "lt" | "lte";

/**
 * A decision the engine makes per document, rather than one the build makes once.
 *
 * `truthy` and `not` came first and are still emitted for the shapes they cover,
 * so an engine that predates the richer forms keeps understanding the common
 * case. The compiler that turns an `if` in a component into one of these picks
 * the narrowest form that fits.
 */
export type Condition =
  | { type: "truthy"; ref: DataReference }
  | { type: "not"; ref: DataReference }
  | {
    type: "compare";
    operator: ComparisonOperator;
    left: ValueExpression;
    right: ValueExpression;
  }
  | { type: "and"; conditions: Condition[] }
  | { type: "or"; conditions: Condition[] }
  | { type: "negate"; condition: Condition };

export type ValueExpression =
  | { type: "literal"; value: string | number | boolean }
  | { type: "ref"; ref: DataReference };

export interface PromptSpec {
  kind: PromptKind;
  text: string;
}

export interface DeriverInvocation {
  output: string;
  name: string;
  inputs: ValueExpression[];
}

export interface BaseNode {
  id: string;
  kind: NodeKind;
  title?: string;
  optional?: boolean;
  when?: Condition;
  prompts?: PromptSpec[];
  derivers?: DeriverInvocation[];
}

export interface SectionNode extends BaseNode {
  kind: "section";
  children: DocumentNode[];
}

export interface ParagraphNode extends BaseNode {
  kind: "paragraph";
  mode: NodeMode;
  text?: string;
}

export interface ImageNode extends BaseNode {
  kind: "image";
  mode: NodeMode;
  path?: string;
  alt?: string;
  width?: number;
  height?: number;
  placeholder?: string;
}

export interface GraphNode extends BaseNode {
  kind: "graph";
  mode: NodeMode;
  graphType: GraphType;
  data?: JsonObject;
  caption?: string;
  placeholder?: string;
}

export interface TableOfContentsNode extends BaseNode {
  kind: "tableOfContents";
}

/**
 * A body repeated once per entry in a request-time collection.
 *
 * A build cannot unroll this the way it unrolls a branch: the length of
 * `source` is not known until a document is written. So the loop itself is
 * what gets published, and the engine walks it. Each pass binds the entry under
 * `as` and its position under `indexAs`, both readable through `ctx`, and
 * suffixes child ids with the index so they stay unique across passes.
 */
export interface RepeatNode extends BaseNode {
  kind: "repeat";
  source: DataReference;
  as: string;
  indexAs: string;
  children: DocumentNode[];
}

export type DocumentNode =
  | SectionNode
  | ParagraphNode
  | ImageNode
  | GraphNode
  | TableOfContentsNode
  | RepeatNode;

export interface DocumentModel {
  schemaVersion: "docxcelerate.letter/v0";
  id: string;
  title: string;
  style?: DocumentStyle;
  metadata?: JsonObject;
  nodes: DocumentNode[];
}

export interface RuntimeState {
  ctx: JsonObject;
  derived: JsonObject;
  dataProvider: DataProvider;
  aiClient: AiClient;
}

export interface DataProvider {
  get(path: string): Promise<unknown> | unknown;
}

export interface AiGenerationRequest {
  node: ParagraphNode;
  prompt: string;
  prompts: PromptSpec[];
  state: RuntimeState;
}

export interface AiImageGenerationRequest {
  node: ImageNode;
  prompt: string;
  prompts: PromptSpec[];
  state: RuntimeState;
}

export interface AiImageResult {
  path: string;
  alt?: string;
  width?: number;
  height?: number;
}

export interface AiGraphGenerationRequest {
  node: GraphNode;
  prompt: string;
  prompts: PromptSpec[];
  state: RuntimeState;
}

export interface AiGraphResult {
  graphType?: GraphType;
  data?: JsonObject;
  caption?: string;
}

export interface AiClient {
  generateParagraph(request: AiGenerationRequest): Promise<string> | string;
  generateImage?(request: AiImageGenerationRequest): Promise<AiImageResult> | AiImageResult;
  generateGraph?(request: AiGraphGenerationRequest): Promise<AiGraphResult> | AiGraphResult;
}
