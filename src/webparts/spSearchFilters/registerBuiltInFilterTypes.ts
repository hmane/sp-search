import type { IRegistry } from '@interfaces/index';
import type { IFilterTypeDefinition } from '@interfaces/index';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';

/**
 * Register all built-in filter types into the given FilterTypeRegistry.
 * Called once from SpSearchFiltersWebPart.onInit().
 *
 * Each filter component is lazy-loaded via createLazyComponent for code splitting.
 */
export function registerBuiltInFilterTypes(registry: IRegistry<IFilterTypeDefinition>): void {
  // ── Checkbox (default) ───────────────────────────────────
  registry.register({
    id: 'checkbox',
    displayName: 'Checkbox List',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    component: createLazyComponent(
      function () { return import('./components/CheckboxFilter') as any; },
      { errorMessage: 'Failed to load checkbox filter' }
    ),
    serializeValue: function (value: unknown): string {
      return String(value);
    },
    deserializeValue: function (raw: string): unknown {
      return raw;
    },
    buildRefinementToken: function (value: unknown, _managedProperty: string): string {
      const v: string = String(value);
      if (v.charAt(0) !== '"') {
        return '"' + v + '"';
      }
      return v;
    }
  });

  // ── Date Range ───────────────────────────────────────────
  registry.register({
    id: 'daterange',
    displayName: 'Date Range',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    component: createLazyComponent(
      function () { return import('./components/DateRangeFilter') as any; },
      { errorMessage: 'Failed to load date range filter' }
    ),
    serializeValue: function (value: unknown): string {
      return encodeURIComponent(String(value));
    },
    deserializeValue: function (raw: string): unknown {
      return decodeURIComponent(raw);
    },
    buildRefinementToken: function (value: unknown, _managedProperty: string): string {
      // Date range values are already FQL range() tokens
      return String(value);
    }
  });

  // ── Toggle (Boolean) ─────────────────────────────────────
  registry.register({
    id: 'toggle',
    displayName: 'Toggle',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    component: createLazyComponent(
      function () { return import('./components/ToggleFilter') as any; },
      { errorMessage: 'Failed to load toggle filter' }
    ),
    serializeValue: function (value: unknown): string {
      return String(value).replace(/"/g, '');
    },
    deserializeValue: function (raw: string): unknown {
      return '"' + raw + '"';
    },
    buildRefinementToken: function (value: unknown, _managedProperty: string): string {
      const v: string = String(value).replace(/"/g, '');
      return '"' + v + '"';
    }
  });

  // ── TagBox ───────────────────────────────────────────────
  registry.register({
    id: 'tagbox',
    displayName: 'Tag Box',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    component: createLazyComponent(
      function () { return import('./components/TagBoxFilter') as any; },
      { errorMessage: 'Failed to load tag box filter' }
    ),
    serializeValue: function (value: unknown): string {
      return encodeURIComponent(String(value));
    },
    deserializeValue: function (raw: string): unknown {
      return decodeURIComponent(raw);
    },
    buildRefinementToken: function (value: unknown, _managedProperty: string): string {
      const v: string = String(value);
      if (v.charAt(0) !== '"') {
        return '"' + v + '"';
      }
      return v;
    }
  });

  // ── Slider (Numeric Range) ───────────────────────────────
  registry.register({
    id: 'slider',
    displayName: 'Slider',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    component: createLazyComponent(
      function () { return import('./components/SliderFilter') as any; },
      { errorMessage: 'Failed to load slider filter' }
    ),
    serializeValue: function (value: unknown): string {
      // Convert range token to compact URL format
      const token: string = String(value);
      const regex: RegExp = /range\((?:decimal\()?([^),]+)\)?,\s*(?:decimal\()?([^),]+)\)?\)/;
      const match: RegExpMatchArray | null = token.match(regex);
      if (match) {
        return match[1].trim() + ':' + match[2].trim();
      }
      return encodeURIComponent(token);
    },
    deserializeValue: function (raw: string): unknown {
      const parts: string[] = raw.split(':');
      if (parts.length === 2) {
        const minPart: string = parts[0] && parts[0] !== 'min' ? 'decimal(' + parts[0] + ')' : 'min';
        const maxPart: string = parts[1] && parts[1] !== 'max' ? 'decimal(' + parts[1] + ')' : 'max';
        return 'range(' + minPart + ', ' + maxPart + ')';
      }
      return decodeURIComponent(raw);
    },
    buildRefinementToken: function (value: unknown, _managedProperty: string): string {
      // Slider emits FQL range() tokens directly
      return String(value);
    }
  });

  // ── Taxonomy Tree ────────────────────────────────────────
  registry.register({
    id: 'taxonomy',
    displayName: 'Taxonomy Tree',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    component: createLazyComponent(
      function () { return import('./components/TaxonomyTreeFilter') as any; },
      { errorMessage: 'Failed to load taxonomy filter' }
    ),
    serializeValue: function (value: unknown): string {
      // Extract GUID for compact URLs
      const token: string = String(value);
      const guidRegex: RegExp = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
      const match: RegExpMatchArray | null = token.match(guidRegex);
      return match ? match[0] : encodeURIComponent(token);
    },
    deserializeValue: function (raw: string): unknown {
      const guidRegex: RegExp = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
      if (guidRegex.test(raw)) {
        return 'GP0|#' + raw;
      }
      return decodeURIComponent(raw);
    },
    buildRefinementToken: function (value: unknown, _managedProperty: string): string {
      // Taxonomy tokens are used as-is (GP0|#GUID format)
      return String(value);
    }
  });

  // ── People Picker ────────────────────────────────────────
  registry.register({
    id: 'people',
    displayName: 'People Picker',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    component: createLazyComponent(
      function () { return import('./components/PeoplePickerFilter') as any; },
      { errorMessage: 'Failed to load people picker filter' }
    ),
    serializeValue: function (value: unknown): string {
      return encodeURIComponent(String(value));
    },
    deserializeValue: function (raw: string): unknown {
      return decodeURIComponent(raw);
    },
    buildRefinementToken: function (value: unknown, _managedProperty: string): string {
      // People claim strings are used as-is for refinement
      return String(value);
    }
  });
}
