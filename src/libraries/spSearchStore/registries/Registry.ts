import { IRegistry } from '@interfaces/index';
import { spLog } from '@store/utils/spLog';

/**
 * Generic typed registry. Stores providers/definitions by ID.
 *
 * Rules:
 * - Duplicate IDs warn and first registration wins (no silent overwrite)
 * - force=true overrides an existing registration
 * - freeze() locks the registry — prevents further mutations
 * - Registries freeze after the first search execution
 */
export class Registry<T extends { id: string }> implements IRegistry<T> {
  private readonly _items: Map<string, T> = new Map();
  private _frozen: boolean = false;
  private readonly _name: string;

  public constructor(name: string) {
    this._name = name;
  }

  public register(provider: T, force?: boolean): void {
    if (this._frozen) {
      spLog.warn('Registry is frozen; provider cannot be registered', {
        registryName: this._name,
        providerId: provider.id,
      });
      return;
    }

    if (this._items.has(provider.id) && !force) {
      spLog.warn('Registry already contains provider; first registration wins', {
        registryName: this._name,
        providerId: provider.id,
      });
      return;
    }

    this._items.set(provider.id, provider);
  }

  public get(id: string): T | undefined {
    return this._items.get(id);
  }

  public getAll(): T[] {
    return Array.from(this._items.values());
  }

  public freeze(): void {
    this._frozen = true;
  }

  public isFrozen(): boolean {
    return this._frozen;
  }
}
