import { ISuggestionProvider, ISearchContext, ISuggestion, ISearchDataProvider } from '@interfaces/index';
import type { IRegistry } from '@interfaces/index';

/**
 * ManagedPropertyProvider — ISuggestionProvider that suggests property-scoped
 * queries like "Author: John Doe" or "FileType: docx" by matching the user's
 * input against available managed properties from the data provider schema.
 *
 * When the user types a partial property name (e.g., "auth"), the provider
 * suggests "Author:" as a KQL property prefix. When the user types
 * "Author: j", it suggests "Author: John Doe" by executing a quick search
 * to find matching values.
 *
 * Requirement: §4.4.2
 */
export class ManagedPropertyProvider implements ISuggestionProvider {
  public readonly id: string = 'managed-property';
  public readonly displayName: string = 'Properties';
  public readonly priority: number = 30;
  public readonly maxResults: number = 5;

  private readonly _dataProviderRegistry: IRegistry<ISearchDataProvider>;

  /** Cached schema properties */
  private _schemaCache: Array<{ name: string; alias: string }> = [];
  private _schemaCacheTimestamp: number = 0;
  private static readonly SCHEMA_CACHE_TTL: number = 10 * 60 * 1000; // 10 minutes

  public constructor(dataProviderRegistry: IRegistry<ISearchDataProvider>) {
    this._dataProviderRegistry = dataProviderRegistry;
  }

  public isEnabled(_context: ISearchContext): boolean {
    return true;
  }

  public async getSuggestions(query: string, _context: ISearchContext): Promise<ISuggestion[]> {
    try {
      const trimmed = query.trim();
      if (trimmed.length < 2) {
        return [];
      }

      // Check if the user is typing a property:value pattern
      const colonIndex = trimmed.indexOf(':');

      if (colonIndex > 0) {
        // User has typed "PropertyName: value" — suggest values
        return this._suggestPropertyValues(trimmed, colonIndex);
      }

      // No colon — suggest property names matching the input
      return await this._suggestPropertyNames(trimmed);
    } catch {
      return [];
    }
  }

  /**
   * Suggest managed property names that match the user's partial input.
   * e.g., typing "auth" suggests "Author:", "AuthorOWSUSER:"
   */
  private async _suggestPropertyNames(partialInput: string): Promise<ISuggestion[]> {
    const schema = await this._getSchema();
    if (schema.length === 0) {
      return [];
    }

    const normalized = partialInput.toLowerCase();
    const suggestions: ISuggestion[] = [];

    for (let i = 0; i < schema.length; i++) {
      const prop = schema[i];
      const matchesName = prop.name.toLowerCase().indexOf(normalized) >= 0;
      const matchesAlias = prop.alias.toLowerCase().indexOf(normalized) >= 0;

      if (matchesName || matchesAlias) {
        const display = prop.alias !== prop.name
          ? prop.alias + ' (' + prop.name + '):'
          : prop.name + ':';
        suggestions.push({
          displayText: display,
          groupName: 'Properties',
          iconName: 'Variable2',
        });

        if (suggestions.length >= this.maxResults) {
          break;
        }
      }
    }

    return suggestions;
  }

  /**
   * When the user has typed "PropertyName: partial", suggest common values.
   * Uses static well-known suggestions for common properties.
   */
  private _suggestPropertyValues(fullInput: string, colonIndex: number): ISuggestion[] {
    const propertyName = fullInput.substring(0, colonIndex).trim();
    const partialValue = fullInput.substring(colonIndex + 1).trim().toLowerCase();

    // Build suggestions for well-known property types
    const knownValues = ManagedPropertyProvider._getKnownValues(propertyName);
    if (knownValues.length === 0) {
      return [];
    }

    const suggestions: ISuggestion[] = [];
    for (let i = 0; i < knownValues.length; i++) {
      if (partialValue.length === 0 || knownValues[i].toLowerCase().indexOf(partialValue) >= 0) {
        suggestions.push({
          displayText: propertyName + ': ' + knownValues[i],
          groupName: 'Properties',
          iconName: 'Variable2',
        });
        if (suggestions.length >= this.maxResults) {
          break;
        }
      }
    }
    return suggestions;
  }

  /**
   * Return well-known values for common managed properties.
   * These are static suggestions that don't require a search call.
   */
  private static _getKnownValues(propertyName: string): string[] {
    const normalized = propertyName.toLowerCase();

    if (normalized === 'filetype' || normalized === 'fileextension') {
      return ['docx', 'xlsx', 'pptx', 'pdf', 'aspx', 'msg', 'txt', 'csv', 'jpg', 'png'];
    }
    if (normalized === 'contentclass') {
      return [
        'STS_ListItem_DocumentLibrary',
        'STS_Site',
        'STS_Web',
        'STS_ListItem',
        'STS_ListItem_GenericList',
        'STS_ListItem_Events',
        'STS_List_850',
      ];
    }
    if (normalized === 'iscontainer') {
      return ['true', 'false'];
    }
    if (normalized === 'isdocument') {
      return ['true', 'false'];
    }

    return [];
  }

  /**
   * Lazy-load and cache schema properties from the data provider.
   */
  private async _getSchema(): Promise<Array<{ name: string; alias: string }>> {
    const now = Date.now();
    if (this._schemaCache.length > 0 && now - this._schemaCacheTimestamp < ManagedPropertyProvider.SCHEMA_CACHE_TTL) {
      return this._schemaCache;
    }

    const providers = this._dataProviderRegistry.getAll();
    for (let i = 0; i < providers.length; i++) {
      const provider = providers[i];
      if (typeof provider.getSchema === 'function') {
        try {
          const props = await provider.getSchema();
          if (props && props.length > 0) {
            const schema: Array<{ name: string; alias: string }> = [];
            for (let j = 0; j < props.length; j++) {
              const p = props[j];
              if (p.queryable) {
                schema.push({
                  name: p.name,
                  alias: p.alias || p.name,
                });
              }
            }
            schema.sort(function (a, b): number { return a.name.localeCompare(b.name); });
            this._schemaCache = schema;
            this._schemaCacheTimestamp = now;
            return schema;
          }
        } catch {
          // Try next provider
        }
      }
    }

    return [];
  }
}
