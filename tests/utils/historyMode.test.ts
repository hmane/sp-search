import { shouldPushHistory } from '../../src/libraries/spSearchStore/utils/historyMode';
import type { IUrlSnapshotForHistory } from '../../src/libraries/spSearchStore/utils/historyMode';

/**
 * T2.D8 — distinguish navigational changes (queryText, currentVerticalKey)
 * worth a new browser-history entry from incremental tweaks (filter
 * toggles, pagination, sort, layout switch) that should `replaceState`.
 * The audit's acceptance signal: "3-query Back/Forward sequence works;
 * filter toggles do not add history".
 */

function snap(overrides: Partial<IUrlSnapshotForHistory> = {}): IUrlSnapshotForHistory {
  return {
    queryText: overrides.queryText || '',
    currentVerticalKey: overrides.currentVerticalKey || 'all',
  };
}

describe('shouldPushHistory — T2.D8', () => {
  describe('navigational changes → push (new history entry)', () => {
    it('queryText changes from empty to non-empty', () => {
      expect(shouldPushHistory(snap({ queryText: '' }), snap({ queryText: 'budget' }))).toBe(true);
    });

    it('queryText changes from one query to another', () => {
      expect(shouldPushHistory(snap({ queryText: 'budget' }), snap({ queryText: 'sales' }))).toBe(true);
    });

    it('queryText cleared (non-empty to empty)', () => {
      expect(shouldPushHistory(snap({ queryText: 'budget' }), snap({ queryText: '' }))).toBe(true);
    });

    it('currentVerticalKey changes (tab switch)', () => {
      expect(shouldPushHistory(snap({ currentVerticalKey: 'all' }), snap({ currentVerticalKey: 'documents' }))).toBe(true);
    });

    it('both queryText and vertical change in the same update', () => {
      expect(shouldPushHistory(
        snap({ queryText: '', currentVerticalKey: 'all' }),
        snap({ queryText: 'budget', currentVerticalKey: 'documents' })
      )).toBe(true);
    });
  });

  describe('incremental changes → replace (no new history entry)', () => {
    it('returns false when nothing changed (the no-op case)', () => {
      const same = snap({ queryText: 'budget', currentVerticalKey: 'documents' });
      expect(shouldPushHistory(same, same)).toBe(false);
    });

    it('returns false when queryText is identical (incremental change elsewhere)', () => {
      expect(shouldPushHistory(
        snap({ queryText: 'budget', currentVerticalKey: 'documents' }),
        snap({ queryText: 'budget', currentVerticalKey: 'documents' })
      )).toBe(false);
    });
  });

  describe('initial-load semantics', () => {
    it('returns false when previous snapshot is undefined (initial load — no history entry yet)', () => {
      expect(shouldPushHistory(undefined, snap({ queryText: 'budget' }))).toBe(false);
    });
  });
});
