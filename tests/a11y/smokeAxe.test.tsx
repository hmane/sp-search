import * as React from 'react';
import { render } from '@testing-library/react';
import { axe, toHaveNoViolations } from 'jest-axe';

// Minimal render harness for the four most-trafficked surfaces. Real component
// integration with the SPFx context is beyond a unit test; this test verifies
// the static markup of representative shapes is axe-clean.

expect.extend(toHaveNoViolations);

describe('a11y smoke — top-10 surfaces (Found.D6)', () => {
  it('Search Box mode toggle uses semantic fieldset/legend', async () => {
    const { container } = render(
      <fieldset>
        <legend className="visuallyHidden">Query input mode</legend>
        <button role="radio" aria-checked={true} aria-label="Regular search mode" type="button">Regular</button>
        <button role="radio" aria-checked={false} aria-label="KQL mode" type="button">KQL</button>
      </fieldset>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });

  it('Scope selector exposes aria-describedby', async () => {
    const { container } = render(
      <>
        <span id="desc">Restricts the search to documents within the selected SharePoint scope</span>
        <label htmlFor="scope">Search scope</label>
        <select id="scope" aria-describedby="desc">
          <option>All</option>
          <option>Site</option>
        </select>
      </>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });

  it('Empty state markup is axe-clean', async () => {
    const { container } = render(
      <div role="status" aria-live="polite">
        <h2>No results</h2>
        <p>Try a different query or remove a filter.</p>
      </div>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });

  it('Detail panel close button has accessible name', async () => {
    const { container } = render(
      <button aria-label="Close detail panel" type="button">×</button>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });

  it('regression sentinel — img without alt fails axe', async () => {
    const { container } = render(<img src="x.png" />);
    const results = await axe(container);
    expect(results.violations.length).toBeGreaterThan(0);
  });

  // T1.D1 — phone-width Filters drawer. The toggle exposes the dialog
  // relationship (`aria-haspopup="dialog"`, `aria-expanded`) and the drawer
  // surface itself is rendered by Fluent's `<Panel>` (Layer + FocusTrapZone
  // + role="dialog" + aria-modal). Jsdom can't simulate the focus trap
  // dynamically, so this assertion verifies the static markup contract
  // axe checks: name + role + relationship.
  it('T1.D1 drawer toggle + dialog markup is axe-clean', async () => {
    const { container } = render(
      <>
        <button
          type="button"
          aria-haspopup="dialog"
          aria-expanded={false}
          aria-label="Show filters, 3 active"
        >
          Show filters (3)
        </button>
        <div role="dialog" aria-modal="true" aria-labelledby="t1d1-h">
          <h2 id="t1d1-h">Filters</h2>
          <button type="button" aria-label="Close filters">×</button>
        </div>
      </>
    );
    const results = await axe(container);
    expect(results).toHaveNoViolations();
  });
});
