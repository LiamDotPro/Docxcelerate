export type JsonObject = Record<string, unknown>;

export type NodeKind = "section" | "paragraph" | "image" | "graph" | "tableOfContents";

export type NodeMode = "static" | "dynamic";

export type GraphType = "bar" | "line" | "pie";

export type LetterStylePreset = "clean-minimal";

export type LetterPageSize = "A4" | "LETTER";

export type LetterPageOrientation = "portrait" | "landscape";

export type LetterFontWeight = "regular" | "bold";

export interface LetterPageMargins {
  topMm: number;
  rightMm: number;
  bottomMm: number;
  leftMm: number;
}

export interface LetterPageStyle {
  size: LetterPageSize;
  orientation: LetterPageOrientation;
  margins: LetterPageMargins;
}

export interface LetterTypographyStyle {
  bodyFont: string;
  headingFont: string;
  bodySizePt: number;
  bodyLineHeight: number;
  color: string;
}

export interface LetterParagraphStyle {
  spacingAfterPt: number;
}

export interface LetterTextBlockStyle {
  fontSizePt: number;
  weight: LetterFontWeight;
  spacingBeforePt: number;
  spacingAfterPt: number;
}

export interface LetterStyle {
  preset: LetterStylePreset;
  page: LetterPageStyle;
  typography: LetterTypographyStyle;
  paragraph: LetterParagraphStyle;
  title: LetterTextBlockStyle;
  sectionHeading: LetterTextBlockStyle;
}

export type PromptKind = "general" | "info" | "negative" | "system";

export type ReferenceScope = "data" | "ctx" | "derived";

export interface DataReference {
  scope: ReferenceScope;
  path: string;
}

export type Condition =
  | { type: "truthy"; ref: DataReference }
  | { type: "not"; ref: DataReference };

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
  children: LetterNode[];
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

export type LetterNode = SectionNode | ParagraphNode | ImageNode | GraphNode | TableOfContentsNode;

export interface LetterDocument {
  schemaVersion: "docxcelerate.letter/v0";
  id: string;
  title: string;
  style?: LetterStyle;
  metadata?: JsonObject;
  nodes: LetterNode[];
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
