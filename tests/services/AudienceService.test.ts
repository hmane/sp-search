import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { resolveUserGroupIds, isInAudience } from '../../src/libraries/spSearchStore/services/AudienceService';

/**
 * Tests for AudienceService.
 *
 * Regression focus: the Graph `/me/memberOf` URL must NOT put `@odata.type` in
 * `$select` — that annotation is not a selectable property and Graph rejects the
 * request with HTTP 400, which made group resolution always fail (fail-closed →
 * audience-targeted items hidden for everyone). `@odata.type` still comes back on
 * the heterogeneous `directoryObject` collection automatically, so the
 * group/directoryRole discriminator below keeps working.
 *
 * `resolveUserGroupIds()` caches per page session; this suite exercises it once.
 * `isInAudience()` is a pure function.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const httpGet = (SPContext as any).http.get as jest.Mock;

describe('AudienceService', () => {
  describe('resolveUserGroupIds', () => {
    it('queries /me/memberOf with $select=id (no @odata.type) and extracts group + directoryRole ids', async () => {
      httpGet.mockResolvedValue({
        ok: true,
        data: {
          value: [
            { id: 'g1', '@odata.type': '#microsoft.graph.group' },
            { id: 'r1', '@odata.type': '#microsoft.graph.directoryRole' },
            { id: 'd1', '@odata.type': '#microsoft.graph.device' }, // not a group/role → ignored
          ],
        },
      });

      const ids = await resolveUserGroupIds();

      expect(httpGet).toHaveBeenCalledTimes(1);
      const calledUrl: string = httpGet.mock.calls[0][0];
      expect(calledUrl).toContain('/me/memberOf');
      expect(calledUrl).toContain('$select=id');
      expect(calledUrl).not.toContain('@odata.type'); // would 400
      expect(ids).toEqual(['g1', 'r1']);
    });
  });

  describe('isInAudience', () => {
    it('returns true when no audience is configured (empty/undefined)', () => {
      expect(isInAudience(undefined, [])).toBe(true);
      expect(isInAudience([], ['g1'])).toBe(true);
    });

    it('returns true when the user is in at least one target group', () => {
      expect(isInAudience(['g1', 'g2'], ['g0', 'g2'])).toBe(true);
    });

    it('returns false when the user is in none of the target groups', () => {
      expect(isInAudience(['g1', 'g2'], ['g3', 'g4'])).toBe(false);
      expect(isInAudience(['g1'], [])).toBe(false);
    });
  });
});
