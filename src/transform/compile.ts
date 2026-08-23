/**
 * Compiling ordinary control flow into decisions a document can carry.
 *
 * A conditional in a component is the authoring surface for a condition. With
 * data in hand it means exactly what it says. Without — publishing to an engine,
 * against stand-in values for a request nobody has made — it cannot: the stand-in
 * is an object, so `if (state.overdue)` is true whatever the field would have
 * held, and `!state.overdue` is false before anything can see which field was
 * meant. JavaScript coerces to a boolean without asking, and no runtime can
 * intercept it. A build that let that through would publish one arm and send it
 * to every recipient.
 *
 * All three ways of writing one are compiled, because all three coerce the same
 * way: the `if` a component returns from, the `cond ? A : B` it writes inside its
 * own JSX, and the `cond && <Node />` it writes more often than either.
 *
 * So the source is rewritten before it runs. The work splits cleanly in two, and
 * that split is the whole reason this is tractable:
 *
 * - **The syntax supplies the operators.** `!`, `===`, `>`, `&&` and `||` are
 *   right there in the tree, so no analysis is needed to know a comparison was
 *   written.
 * - **The values supply the paths.** Each leaf is wrapped in a thunk, and what
 *   it evaluates to knows where it came from — a stand-in carries its own path,
 *   a real value is just a value.
 *
 * Neither half has to trace anything back through `useState` to the request,
 * which is the analysis this deliberately avoids.
 *
 * @module
 */

import ts from "typescript";

/** What {@linkcode transformDocumentSource} takes beyond the source itself. */
export interface TransformOptions {
  /** The file being compiled, used in messages and to pick the syntax. */
  fileName?: string;
  /** Where the emitted helpers are imported from. Defaults to the package. */
  runtimeModule?: string;
}

/** What a compile produced. */
export interface TransformResult {
  /** The rewritten source. */
  code: string;
  /** How many conditionals became branches. */
  branches: number;
  /** Whether anything was rewritten at all. */
  changed: boolean;
}

/**
 * The mark a compiled module carries.
 *
 * A publish build refuses a module without it. That refusal is the whole point:
 * a document project compiled by a plain `tsc` still runs, and still produces a
 * document — it just quietly takes one arm of every decision. Without a mark there
 * is nothing to notice that by, and the failure surfaces as a recipient reading
 * the wrong letter.
 */
export const compiledMarker = "__docxcelerateCompiled";

const helpers = {
  branch: "branch",
  test: "__test",
  not: "__not",
  compare: "__compare",
  and: "__and",
  or: "__or",
} as const;

const comparisons = new Map<ts.SyntaxKind, string>([
  [ts.SyntaxKind.EqualsEqualsEqualsToken, "eq"],
  [ts.SyntaxKind.EqualsEqualsToken, "eq"],
  [ts.SyntaxKind.ExclamationEqualsEqualsToken, "ne"],
  [ts.SyntaxKind.ExclamationEqualsToken, "ne"],
  [ts.SyntaxKind.GreaterThanToken, "gt"],
  [ts.SyntaxKind.GreaterThanEqualsToken, "gte"],
  [ts.SyntaxKind.LessThanToken, "lt"],
  [ts.SyntaxKind.LessThanEqualsToken, "lte"],
]);

/**
 * Rewrites the conditionals of a document module into published branches.
 *
 * Three shapes are rewritten, and they are the three whose outcomes are both
 * expressible as arms of a branch and whose arms yield nodes: an `if` that
 * returns, a ternary, and a `&&` before a node. Everything else is ordinary
 * control flow with nothing to publish — an `if` that assigns to a variable, a
 * ternary picking a string for a prop — and is left exactly as written.
 *
 * @param source The module's source text.
 * @param options The file name, and where the helpers are imported from.
 * @returns The rewritten source, and how much was rewritten.
 */
export function transformDocumentSource(
  source: string,
  options: TransformOptions = {},
): TransformResult {
  const fileName = options.fileName ?? "document.tsx";
  const file = ts.createSourceFile(
    fileName,
    source,
    ts.ScriptTarget.ESNext,
    true,
    fileName.endsWith(".tsx") || fileName.endsWith(".jsx")
      ? ts.ScriptKind.TSX
      : ts.ScriptKind.TS,
  );

  let branches = 0;
  const used = new Set<string>();
  const factory = ts.factory;

  const call = (name: string, args: ts.Expression[]): ts.CallExpression => {
    used.add(name);
    return factory.createCallExpression(factory.createIdentifier(name), undefined, args);
  };

  const thunk = (body: ts.Expression): ts.ArrowFunction =>
    factory.createArrowFunction(undefined, undefined, [], undefined, undefined, body);

  /**
   * Turns a condition expression into a call that yields a test.
   *
   * The shape is read off the syntax; the leaves are left to say for themselves
   * what they refer to, which is what the thunks are for.
   */
  const testOf = (expression: ts.Expression): ts.Expression => {
    if (ts.isParenthesizedExpression(expression)) {
      return testOf(expression.expression);
    }

    if (
      ts.isPrefixUnaryExpression(expression) &&
      expression.operator === ts.SyntaxKind.ExclamationToken
    ) {
      return call(helpers.not, [testOf(expression.operand)]);
    }

    if (ts.isBinaryExpression(expression)) {
      const operator = comparisons.get(expression.operatorToken.kind);

      if (operator) {
        return call(helpers.compare, [
          thunkedValue(expression.left),
          factory.createStringLiteral(operator),
          thunkedValue(expression.right),
        ]);
      }

      if (expression.operatorToken.kind === ts.SyntaxKind.AmpersandAmpersandToken) {
        return call(helpers.and, [testOf(expression.left), testOf(expression.right)]);
      }

      if (expression.operatorToken.kind === ts.SyntaxKind.BarBarToken) {
        return call(helpers.or, [testOf(expression.left), testOf(expression.right)]);
      }
    }

    return call(helpers.test, [expression]);
  };

  // A comparison side is passed as the value itself. It is not a test, so it is
  // not wrapped in one — `expr` on the other side works out whether it is a real
  // value or a stand-in that knows its own path.
  const thunkedValue = (expression: ts.Expression): ts.Expression =>
    ts.isParenthesizedExpression(expression) ? thunkedValue(expression.expression) : expression;

  const transformer: ts.TransformerFactory<ts.SourceFile> = (context) => {
    const visit = (node: ts.Node): ts.Node => {
      if (ts.isBlock(node) || ts.isSourceFile(node)) {
        const statements = rewriteStatements(
          node.statements,
          (child) => ts.visitNode(child, visit) as ts.Statement,
        );

        if (statements) {
          return ts.isBlock(node)
            ? factory.updateBlock(node, statements)
            : factory.updateSourceFile(node as ts.SourceFile, statements);
        }
      }

      if (ts.isConditionalExpression(node)) {
        const compiled = choiceOf(node, visit);

        if (compiled) {
          return compiled;
        }
      }

      if (ts.isBinaryExpression(node)) {
        const compiled = guardOf(node, visit);

        if (compiled) {
          return compiled;
        }
      }

      return ts.visitEachChild(node, visit, context);
    };

    /**
     * The expression one arm of a decision yields.
     *
     * A `return` is the arm itself. A nested `if` is another decision, so it
     * becomes another branch — which is all an `else if` chain is, and why one
     * of any length needs no case of its own. An arm that is absent falls back
     * to whatever followed the `if`, because returning on one path and carrying
     * on down the function is the same decision written the short way.
     */
    const armOf = (
      statement: ts.Statement | undefined,
      trailing: readonly ts.Statement[],
      visitStatement: (statement: ts.Statement) => ts.Statement,
    ): ts.Expression | undefined => {
      if (!statement) {
        return trailing.length === 1 && ts.isReturnStatement(trailing[0])
          ? trailing[0].expression ?? nothing()
          : undefined;
      }

      if (ts.isReturnStatement(statement)) {
        return statement.expression ?? nothing();
      }

      if (ts.isIfStatement(statement)) {
        return branchOf(statement, trailing, visitStatement);
      }

      if (ts.isBlock(statement)) {
        if (statement.statements.length === 1) {
          return armOf(statement.statements[0], trailing, visitStatement);
        }

        // A block that does some work and then decides. Only the decision can
        // become an arm, so anything before it stays where it is and this is
        // left for the ordinary walk to reach.
        return undefined;
      }

      return undefined;
    };

    /** One `if`, compiled into the branch that carries both its outcomes. */
    const branchOf = (
      statement: ts.IfStatement,
      trailing: readonly ts.Statement[],
      visitStatement: (statement: ts.Statement) => ts.Statement,
    ): ts.Expression | undefined => {
      const whenTrue = armOf(statement.thenStatement, trailing, visitStatement);

      if (whenTrue === undefined) {
        return undefined;
      }

      const whenFalse = armOf(statement.elseStatement, trailing, visitStatement);

      if (whenFalse === undefined) {
        return undefined;
      }

      // An ordinary function that happens to return early is not a document
      // conditional. `formatList` picking a string, a guard clause returning a
      // number — those are control flow with nothing to publish, and rewriting
      // them would turn a string into a branch element. What marks a component
      // conditional is that an arm yields a node.
      if (!containsJsx(whenTrue) && !containsJsx(whenFalse)) {
        return undefined;
      }

      branches += 1;

      return call(helpers.branch, [
        testOf(statement.expression),
        thunk(visitExpression(whenTrue, visitStatement)),
        thunk(visitExpression(whenFalse, visitStatement)),
      ]);
    };

    /**
     * A `cond ? A : B` whose arms are nodes, compiled into the same branch.
     *
     * The ternary is the other half of the authoring surface an `if` opens. A
     * component that decides between two nodes inside its own JSX has nowhere
     * to put a statement, so it writes the choice as an expression — and that
     * expression coerces its stand-in to a boolean exactly as the `if` would
     * have, silently, before anything can see which field was meant.
     *
     * The line is drawn where it is for `if`: an arm has to yield a node. A
     * ternary picking between two strings for a prop is a value, and the
     * derivers are what carry a value that varies per document.
     */
    const choiceOf = (
      node: ts.ConditionalExpression,
      visitNode: (node: ts.Node) => ts.Node,
    ): ts.Expression | undefined => {
      if (!containsJsx(node.whenTrue) && !containsJsx(node.whenFalse)) {
        return undefined;
      }

      branches += 1;

      return call(helpers.branch, [
        testOf(node.condition),
        thunk(visited(node.whenTrue, visitNode)),
        thunk(visited(node.whenFalse, visitNode)),
      ]);
    };

    /**
     * A `cond && <Node />`, compiled into a branch with one arm.
     *
     * This is the ternary with its else left off, and it means what leaving the
     * else off an `if` means: the node is absent in that case. It is written far
     * more often than either, because it is the shortest way to say a paragraph
     * only sometimes belongs in the document.
     *
     * `||` and `??` are left alone. Neither yields the right-hand side when the
     * test holds, so neither is a way of writing a document decision — whatever
     * they are doing in a component, it is not this.
     */
    const guardOf = (
      node: ts.BinaryExpression,
      visitNode: (node: ts.Node) => ts.Node,
    ): ts.Expression | undefined => {
      if (node.operatorToken.kind !== ts.SyntaxKind.AmpersandAmpersandToken) {
        return undefined;
      }

      if (!containsJsx(node.right)) {
        return undefined;
      }

      branches += 1;

      return call(helpers.branch, [
        testOf(node.left),
        thunk(visited(node.right, visitNode)),
        thunk(nothing()),
      ]);
    };

    /**
     * Rewrites a run of statements whose tail is a conditional return.
     *
     * Both shapes a component writes end up here: `if (…) return A; return B;`
     * is the early-return form and `if (…) { return A } else { return B }` is
     * the explicit one.
     */
    const rewriteStatements = (
      statements: ts.NodeArray<ts.Statement>,
      visitStatement: (statement: ts.Statement) => ts.Statement,
    ): ts.Statement[] | undefined => {
      for (const [index, statement] of statements.entries()) {
        if (!ts.isIfStatement(statement)) {
          continue;
        }

        const compiled = branchOf(statement, statements.slice(index + 1), visitStatement);

        if (compiled === undefined) {
          continue;
        }

        return [
          ...statements.slice(0, index).map(visitStatement),
          factory.createReturnStatement(compiled),
        ];
      }

      return undefined;
    };

    return (file) => ts.visitNode(file, visit) as ts.SourceFile;
  };

  const result = ts.transform(file, [transformer]);
  const printed = ts.createPrinter({ newLine: ts.NewLineKind.LineFeed })
    .printFile(result.transformed[0]);

  result.dispose();

  if (used.size === 0) {
    return { code: source, branches: 0, changed: false };
  }

  return {
    code: `${importOf(used, options.runtimeModule ?? "docxcelerate/template")}\n` +
      `${stampOf()}\n${printed}`,
    branches,
    changed: true,
  };
}

/** A source file that holds a decision nothing compiled. */
export interface UncompiledSource {
  /** The file, as it was named to the check. */
  fileName: string;
  /** How many decisions in it would have been compiled. */
  branches: number;
}

/**
 * Finds the sources that carry a decision but were never compiled.
 *
 * The transform is its own detector: a file it would have rewritten, that does
 * not already carry the mark, is one that reached the build another way. That
 * matters because the failure is silent — an uncompiled conditional still runs, still
 * produces a document, and still takes one arm of every decision, so a project
 * built by a plain `tsc` publishes a document that is wrong for everybody whose
 * data would have taken the other arm.
 *
 * @param files The sources to check, each with the name to report it under.
 * @returns The files that need compiling, in the order they were given.
 */
export function findUncompiledSources(
  files: Iterable<{ fileName: string; source: string }>,
): UncompiledSource[] {
  const uncompiled: UncompiledSource[] = [];

  for (const { fileName, source } of files) {
    if (isCompiledSource(source)) {
      continue;
    }

    const result = transformDocumentSource(source, { fileName });

    if (result.changed) {
      uncompiled.push({ fileName, branches: result.branches });
    }
  }

  return uncompiled;
}

/**
 * Refuses a build whose sources were not compiled.
 *
 * @param files The sources to check.
 * @throws If any of them holds a decision that nothing compiled.
 */
export function assertCompiledSources(
  files: Iterable<{ fileName: string; source: string }>,
): void {
  const uncompiled = findUncompiledSources(files);

  if (uncompiled.length === 0) {
    return;
  }

  const listed = uncompiled
    .map(({ fileName, branches }) => `  ${fileName} (${branches} decision${branches === 1 ? "" : "s"})`)
    .join(String.fromCharCode(10));

  throw new Error(
    [
      "These files decide something a document has to carry, but nothing compiled them:",
      listed,
      "A conditional that is not compiled still runs — it just takes one arm, once, and publishes",
      "that arm to every recipient. Add the Docxcelerate transform to the build that produced",
      "them: docxcelerateTransform() for Vite, docxcelerateEsbuildTransform() for esbuild,",
      "both from docxcelerate/transform.",
    ].join(String.fromCharCode(10)),
  );
}

/**
 * Whether a module was compiled by this transform.
 *
 * @param source The module's source text.
 * @returns `true` when it carries the mark.
 */
export function isCompiledSource(source: string): boolean {
  return source.includes(compiledMarker);
}

/**
 * An arm that returns nothing, for an `if` whose other arm returns a node.
 *
 * A component that returns on one path and falls off the end on the other is
 * saying the node is absent in that case, which is exactly what an arm yielding
 * nothing means.
 */
function nothing(): ts.Expression {
  return ts.factory.createIdentifier("undefined");
}

/**
 * Whether an expression yields part of a document.
 *
 * Read off the syntax rather than the types, so it costs nothing and needs no
 * program. A component conditional written any other way — building the node
 * into a variable first, then returning the variable — is not recognised, and
 * is the shape the compiled-source check exists to catch.
 */
function containsJsx(expression: ts.Expression): boolean {
  let found = false;

  const walk = (node: ts.Node): void => {
    if (found) {
      return;
    }

    if (ts.isJsxElement(node) || ts.isJsxSelfClosingElement(node) || ts.isJsxFragment(node)) {
      found = true;
      return;
    }

    ts.forEachChild(node, walk);
  };

  walk(expression);

  return found;
}

/** Visits inside an arm, so a decision nested in one is compiled too. */
function visited(
  expression: ts.Expression,
  visitNode: (node: ts.Node) => ts.Node,
): ts.Expression {
  return ts.visitNode(expression, visitNode) as ts.Expression;
}

/** Visits inside an arm, so an `if` nested in one is compiled too. */
function visitExpression(
  expression: ts.Expression,
  visitStatement: (statement: ts.Statement) => ts.Statement,
): ts.Expression {
  const wrapped = ts.factory.createExpressionStatement(expression);
  const visited = visitStatement(wrapped);

  return ts.isExpressionStatement(visited) ? visited.expression : expression;
}

function importOf(used: Set<string>, module: string): string {
  const names = [...used].sort().join(", ");

  return `import { ${names} } from ${JSON.stringify(module)};`;
}

function stampOf(): string {
  return `const ${compiledMarker} = true;\nexport { ${compiledMarker} };`;
}
