/**
 * Trail-marker smoke test for the Jest harness post-Found.D13.
 * Real lifecycle assertions land in T3.D9 (Sprint 6 dep on this fix).
 * Until then this test exists only to prove npm test runs at least one
 * spec end-to-end through the SPFx-Heft Jest pipeline.
 */
describe('Foundations Found.D13 — Jest harness smoke', () => {
  it('runs at least one spec to completion', () => {
    expect(1 + 1).toBe(2);
  });

  it('resolves ts-jest TypeScript transform', () => {
    const value: number = 42;
    expect(value).toBe(42);
  });
});
