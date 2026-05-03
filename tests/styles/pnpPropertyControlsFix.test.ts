describe('pnpPropertyControlsFix', () => {
  beforeEach(() => {
    document.head.querySelectorAll('#sp-search-pnp-property-controls-fix').forEach(n => n.remove());
    jest.resetModules();
  });

  it('injects the style tag with non-empty content', async () => {
    const { ensurePnpPropertyControlStyles } = await import('../../src/styles/pnpPropertyControlsFix');
    ensurePnpPropertyControlStyles();
    const tag = document.getElementById('sp-search-pnp-property-controls-fix');
    expect(tag).not.toBeNull();
    expect(tag!.textContent).toContain('.collectionData_f8375039');
    expect(tag!.textContent).toContain('.tableRow_f8375039');
  });

  it('is idempotent across multiple calls within one module instance', async () => {
    const { ensurePnpPropertyControlStyles } = await import('../../src/styles/pnpPropertyControlsFix');
    ensurePnpPropertyControlStyles();
    ensurePnpPropertyControlStyles();
    expect(document.querySelectorAll('#sp-search-pnp-property-controls-fix')).toHaveLength(1);
  });

  it('re-injects after a fresh module load (proves module-level injected flag is what guards repeats)', async () => {
    const first = await import('../../src/styles/pnpPropertyControlsFix');
    first.ensurePnpPropertyControlStyles();
    document.head.querySelectorAll('#sp-search-pnp-property-controls-fix').forEach(n => n.remove());
    jest.resetModules();
    const second = await import('../../src/styles/pnpPropertyControlsFix');
    second.ensurePnpPropertyControlStyles();
    expect(document.getElementById('sp-search-pnp-property-controls-fix')).not.toBeNull();
  });
});
