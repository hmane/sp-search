import { safeNavigate } from '../../src/libraries/spSearchStore/utils/safeNavigate';

describe('safeNavigate (Found.D4)', () => {
  let assignedTo: string | null = null;
  let originalLocation: Location;

  beforeEach(() => {
    assignedTo = null;
    originalLocation = window.location;
    Object.defineProperty(window, 'location', {
      writable: true,
      configurable: true,
      value: {
        ...window.location,
        assign: (url: string) => {
          assignedTo = url;
        },
      },
    });
  });

  afterEach(() => {
    Object.defineProperty(window, 'location', {
      writable: true,
      configurable: true,
      value: originalLocation,
    });
  });

  it('allows https:// URLs', () => {
    expect(safeNavigate('https://example.com/doc.pdf')).toBe(true);
    expect(assignedTo).toBe('https://example.com/doc.pdf');
  });

  it('allows http:// URLs', () => {
    expect(safeNavigate('http://example.com/')).toBe(true);
    expect(assignedTo).toBe('http://example.com/');
  });

  it('allows root-relative paths', () => {
    expect(safeNavigate('/sites/SPSearch/Pages/Search.aspx')).toBe(true);
    expect(assignedTo).toBe('/sites/SPSearch/Pages/Search.aspx');
  });

  it('rejects javascript: URLs', () => {
    expect(safeNavigate('javascript:alert(1)')).toBe(false);
    expect(assignedTo).toBeNull();
  });

  it('rejects data: URLs', () => {
    expect(safeNavigate('data:text/html,<script>alert(1)</script>')).toBe(false);
    expect(assignedTo).toBeNull();
  });

  it('rejects empty / null / undefined', () => {
    expect(safeNavigate('')).toBe(false);
    expect(safeNavigate(null as unknown as string)).toBe(false);
    expect(safeNavigate(undefined as unknown as string)).toBe(false);
    expect(assignedTo).toBeNull();
  });

  it('rejects whitespace-only', () => {
    expect(safeNavigate('   ')).toBe(false);
    expect(assignedTo).toBeNull();
  });
});
