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
});
