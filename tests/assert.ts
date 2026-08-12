import assert from "node:assert/strict";

/**
 * Small shims over node:assert so test bodies read the same way they did
 * under Deno's std assertions.
 */

export function assertEquals<T>(actual: T, expected: T): void {
  assert.deepStrictEqual(actual, expected);
}

export function assertStringIncludes(actual: string, expected: string): void {
  assert.ok(
    actual.includes(expected),
    `Expected string to include ${JSON.stringify(expected)}.`,
  );
}

export function assertRejects(
  fn: () => Promise<unknown>,
  errorClass?: new (...args: never[]) => Error,
  messageIncludes?: string,
): Promise<void> {
  return assert.rejects(fn, (error: unknown) => {
    if (errorClass && !(error instanceof errorClass)) {
      return false;
    }

    if (messageIncludes) {
      const message = error instanceof Error ? error.message : String(error);
      assert.ok(
        message.includes(messageIncludes),
        `Expected error message to include ${JSON.stringify(messageIncludes)}, got ${
          JSON.stringify(message)
        }.`,
      );
    }

    return true;
  });
}
