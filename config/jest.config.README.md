# Jest config — deferred test suites

> Inline JSON commentary (e.g. a `_comment_*` sibling key on `testPathIgnorePatterns`) is NOT supported — Jest's schema validator emits a `Validation Warning: Unknown option …` on startup for any unrecognised top-level key. Capture rationale in this README and reference it from PR descriptions / commit messages instead.

Per Found.D13 follow-up: `testPathIgnorePatterns` in `jest.config.json` excludes 9 pre-existing test suites that fail to load due to `scope` field schema drift (the `scope` field was removed from `ISearchHistoryEntry` and `IUrlState` after these tests were written).

These suites are NOT broken specs — they are pre-existing test rot uncovered by the now-working Heft Jest harness. Real fixes belong to:

- **T3.D9** (Sprint 6 — `disposeStore` regression test + lifecycle smoke harness): re-enable `tests/store/slices/*.test.ts` and `tests/middleware/urlSyncMiddleware.test.ts` after fixing the `scope` references. Audit Roadmap Matrix row T3.D9 explicitly depends on Found.D13 — this deferral is the dependency contract.
- **T5.D2** (Sprint 5 — Network tab + per-call timing): re-enable `tests/services/{TokenService,SearchService}.test.ts` after the DebugCollector wire surfaces the new contract.

When re-enabling, remove the corresponding `testPathIgnorePatterns` entry AND fix the underlying schema drift in `tests/utils/testHelpers.ts` (line 181) and `tests/middleware/urlSyncMiddleware.test.ts` (lines 360, 417, 432, 485).

> Note: `tests/utils/testHelpers.ts` is a shared helper module imported by the deferred slice tests, not a spec itself. It is naturally excluded from discovery by the `testMatch` glob (`*.test.ts(x)`) so it does not need a `testPathIgnorePatterns` entry of its own — the schema drift inside it only surfaces transitively when the deferred slice suites get re-enabled and import from it.

## CLI footgun: `--test-path-pattern` (kebab-case)

Heft's Jest plugin accepts `--test-path-pattern` (kebab-case) only. Passing Jest's native `--testPathPattern` (camelCase) is silently consumed by Heft before Jest sees it, with no error and no filter applied. Always use the kebab-case form when invoking via `npm test`.
