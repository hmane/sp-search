import { IRegistry } from '@interfaces/index';

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
      console.warn(
        `[SP Search] ${this._name} registry is frozen. Cannot register "${provider.id}". ` +
        `Registries lock after the first search execution.`
      );
      return;
    }

    if (this._items.has(provider.id) && !force) {
      console.warn(
        `[SP Search] ${this._name} registry already contains "${provider.id}". ` +
        `First registration wins. Use force=true to override.`
      );
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
