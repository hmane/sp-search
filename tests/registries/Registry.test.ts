import { Registry } from '../../src/libraries/spSearchStore/registries/Registry';

interface IMockProvider {
  id: string;
  name: string;
}

describe('Registry', () => {
  let registry: Registry<IMockProvider>;

  beforeEach(() => {
    registry = new Registry<IMockProvider>('TestProvider');
  });

  describe('register', () => {
    it('should add a provider', () => {
      const provider: IMockProvider = { id: 'prov-1', name: 'Provider One' };
      registry.register(provider);

      expect(registry.get('prov-1')).toEqual(provider);
    });

    it('should allow registering multiple providers with different IDs', () => {
      registry.register({ id: 'prov-1', name: 'Provider One' });
      registry.register({ id: 'prov-2', name: 'Provider Two' });
      registry.register({ id: 'prov-3', name: 'Provider Three' });

      expect(registry.getAll()).toHaveLength(3);
    });

    it('should warn and keep first registration on duplicate ID', () => {
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation();

      const first: IMockProvider = { id: 'dup', name: 'First' };
      const second: IMockProvider = { id: 'dup', name: 'Second' };

      registry.register(first);
      registry.register(second);

      expect(warnSpy).toHaveBeenCalledTimes(1);
      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining('already contains "dup"')
      );

      // First registration wins
      expect(registry.get('dup')!.name).toBe('First');

      warnSpy.mockRestore();
    });

    it('should override existing registration with force=true', () => {
      const first: IMockProvider = { id: 'prov', name: 'First' };
      const second: IMockProvider = { id: 'prov', name: 'Override' };

      registry.register(first);
      registry.register(second, true);

      expect(registry.get('prov')!.name).toBe('Override');
    });

    it('should not warn when using force=true', () => {
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation();

      registry.register({ id: 'prov', name: 'First' });
      registry.register({ id: 'prov', name: 'Override' }, true);

      expect(warnSpy).not.toHaveBeenCalled();

      warnSpy.mockRestore();
    });
  });

  describe('freeze', () => {
    it('should mark the registry as frozen', () => {
      expect(registry.isFrozen()).toBe(false);
      registry.freeze();
      expect(registry.isFrozen()).toBe(true);
    });

    it('should prevent new registrations after freeze', () => {
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation();

      registry.register({ id: 'before-freeze', name: 'Allowed' });
      registry.freeze();
      registry.register({ id: 'after-freeze', name: 'Blocked' });

      expect(warnSpy).toHaveBeenCalledTimes(1);
      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining('registry is frozen')
      );

      // The blocked registration should not be present
      expect(registry.get('after-freeze')).toBeUndefined();

      // The existing registration should still be there
      expect(registry.get('before-freeze')).toBeDefined();

      warnSpy.mockRestore();
    });

    it('should prevent force registrations after freeze', () => {
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation();

      registry.register({ id: 'existing', name: 'Original' });
      registry.freeze();
      registry.register({ id: 'existing', name: 'Force Override' }, true);

      // Even force should not work after freeze
      expect(registry.get('existing')!.name).toBe('Original');

      warnSpy.mockRestore();
    });

    it('should allow reading after freeze', () => {
      registry.register({ id: 'prov-1', name: 'Provider One' });
      registry.register({ id: 'prov-2', name: 'Provider Two' });
      registry.freeze();

      expect(registry.get('prov-1')).toBeDefined();
      expect(registry.getAll()).toHaveLength(2);
    });

    it('should be idempotent — calling freeze twice is safe', () => {
      registry.freeze();
      registry.freeze();
      expect(registry.isFrozen()).toBe(true);
    });
  });

  describe('get', () => {
    it('should return the registered provider by ID', () => {
      const provider: IMockProvider = { id: 'prov-1', name: 'Provider One' };
      registry.register(provider);
      expect(registry.get('prov-1')).toEqual(provider);
    });

    it('should return undefined for a non-existent ID', () => {
      expect(registry.get('nonexistent')).toBeUndefined();
    });

    it('should return undefined for an empty string ID', () => {
      expect(registry.get('')).toBeUndefined();
    });

    it('should return the same reference that was registered', () => {
      const provider: IMockProvider = { id: 'ref-test', name: 'Reference' };
      registry.register(provider);
      expect(registry.get('ref-test')).toBe(provider);
    });
  });

  describe('getAll', () => {
    it('should return all registered providers', () => {
      registry.register({ id: 'a', name: 'A' });
      registry.register({ id: 'b', name: 'B' });
      registry.register({ id: 'c', name: 'C' });

      const all = registry.getAll();
      expect(all).toHaveLength(3);
      expect(all.map(p => p.id)).toEqual(expect.arrayContaining(['a', 'b', 'c']));
    });

    it('should return an empty array when nothing is registered', () => {
      expect(registry.getAll()).toEqual([]);
    });

    it('should return a new array instance on each call', () => {
      registry.register({ id: 'a', name: 'A' });
      const all1 = registry.getAll();
      const all2 = registry.getAll();
      expect(all1).not.toBe(all2);
      expect(all1).toEqual(all2);
    });
  });

  describe('isFrozen', () => {
    it('should return false initially', () => {
      expect(registry.isFrozen()).toBe(false);
    });

    it('should return true after freeze', () => {
      registry.freeze();
      expect(registry.isFrozen()).toBe(true);
    });
  });
});
