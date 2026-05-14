import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { StoreApi } from 'zustand/vanilla';
import { spfxToolkitStylesLoaded } from '../../styles/loadSpfxToolkitStyles';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SpSearchFiltersWebPartStrings';
import SpSearchFilters from './components/SpSearchFilters';
import type { ISpSearchFiltersProps } from './components/ISpSearchFiltersProps';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { getStore, initializeSearchContext, incrementContextRef, decrementContextRef } from '@store/store';
import { recordWebPartInit } from '@store/utils/initOrderDiagnostic';
import type { ISearchStore, IFilterConfig } from '@interfaces/index';
import { registerBuiltInFilterTypes } from './registerBuiltInFilterTypes';
import { SharePointSearchProvider } from '@providers/index';
import { DisplayMode } from '@microsoft/sp-core-library';
import { AudienceGate, parseAudienceGroups } from '../../utilities/AudienceGate';
import { SearchContextIdBannerWrapper } from '../../utilities/SearchContextIdMismatchBanner';
import { SPDebugProvider } from 'spfx-toolkit/lib/components/debug';
import {
  propertyPaneSearchContextIdField,
  SEARCH_CONTEXT_ID_GROUP_NAME,
} from '../../propertyPaneControls/PropertyPaneSearchContextIdField';
// T4.D11 — context-sensitive help link helper.
import { propertyPaneGroupHelp } from '../../propertyPaneControls/propertyPaneGroupHelp';
import { sanitizeUrlAlias } from '@store/utils/filterUrlAliases';
import { DebugCollector } from '@store/debug';
import { ensurePnpPropertyControlStyles } from '../../styles/pnpPropertyControlsFix';

// Bundle DevExtreme CSS — injected via style-loader at runtime.
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.common.css');
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('devextreme/dist/css/dx.light.css');
// eslint-disable-next-line @typescript-eslint/no-unused-vars
const _ensureStyles = spfxToolkitStylesLoaded;

export interface ISpSearchFiltersWebPartProps {
  searchContextId: string;
  filtersCollection: IFilterCollectionItem[];
  applyMode: 'instant' | 'manual';
  operatorBetweenFilters: 'AND' | 'OR';
  showClearAll: boolean;
  enableVisualFilterBuilder: boolean;
  /** Stream D / #10 — comma/newline-separated Azure AD group IDs. Empty = visible to everyone. */
  audienceGroups: string;
}

interface IFilterCollectionItem {
  uniqueId: string;
  managedProperty: string;
  displayName: string;
  urlAlias?: string;
  filterType: string;
  operator: string;
  maxValues: number;
  defaultExpanded: boolean;
  showCount: boolean;
  sortBy: string;
  sortDirection: string;
  multiValues: boolean;
  dependsOn?: string;
  showWhenParentHasValue?: boolean;
  hideZeroCountValues?: boolean;
  resetWhenParentChanges?: boolean;
  trueLabel?: string;
  falseLabel?: string;
  invertBoolean?: boolean;
  /** Stream D / #5 — comma-separated Azure AD group object IDs (admin UX matches Verticals). */
  audience?: string;
}

function normalizeFiltersCollectionValue(
  raw: IFilterCollectionItem[] | string | undefined
): IFilterCollectionItem[] {
  if (Array.isArray(raw)) {
    return raw;
  }
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw) as IFilterCollectionItem[];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }
  return [];
}

function normalizeManagedPropertyForFilter(
  managedProperty: string,
  filterType: IFilterConfig['filterType']
): string {
  if (filterType === 'people' && managedProperty === 'Author') {
    return 'AuthorOWSUSER';
  }

  return managedProperty;
}

export default class SpSearchFiltersWebPart extends BaseClientSideWebPart<ISpSearchFiltersWebPartProps> {

  private _store: StoreApi<ISearchStore> | undefined;

  public render(): void {
    const innerElement: React.ReactElement<ISpSearchFiltersProps> = React.createElement(
      SpSearchFilters,
      {
        store: this._store,
        applyMode: this.properties.applyMode || 'instant',
        showClearAll: this.properties.showClearAll !== false,
        enableVisualFilterBuilder: !!this.properties.enableVisualFilterBuilder,
        isEditMode: this.displayMode === DisplayMode.Edit,
        // T4.D12 — cross-web-part preset propagation.
        searchContextId: this.properties.searchContextId || 'default',
        onApplyPresetFilters: (filterRows): void => {
          // Build new filter rows with sensible operator/option defaults
          // for each suggested filter type. Existing rows are kept and the
          // suggestions are appended only when the managedProperty isn't
          // already configured.
          const existing = normalizeFiltersCollectionValue(this.properties.filtersCollection);
          const haveProps = new Set(existing.map((r) => (r.managedProperty || '').toLowerCase()));
          const additions: IFilterCollectionItem[] = filterRows
            .filter((r) => !haveProps.has(r.managedProperty.toLowerCase()))
            .map((r, idx): IFilterCollectionItem => ({
              uniqueId: 'preset-filter-' + String(Date.now()) + '-' + String(idx),
              managedProperty: r.managedProperty,
              displayName: r.label,
              urlAlias: r.urlAlias,
              filterType: r.filterType,
              operator: '=',
              maxValues: 10,
              defaultExpanded: true,
              showCount: true,
              sortBy: 'count',
              sortDirection: 'desc',
              multiValues: true,
            }));
          if (additions.length === 0) { return; }
          this.properties.filtersCollection = existing.concat(additions);
          this.context.propertyPane.refresh();
          this.render();
        }
      }
    );

    // Stream D / #10 — wrap with AudienceGate so the web part hides itself
    // when the current user isn't in any of the configured groups.
    const audienceGroups = parseAudienceGroups(this.properties.audienceGroups);
    const gatedElement: React.ReactElement = React.createElement(
      AudienceGate,
      { audienceGroups, store: this._store },
      innerElement
    );

    // T3.D2 — edit-mode mismatch banner above the gated tree.
    const bannerWrapped: React.ReactElement = React.createElement(
      SearchContextIdBannerWrapper,
      {
        webPartId: this.instanceId,
        contextId: this.properties.searchContextId || 'default',
        webPartLabel: 'SP Search Filters',
        isEditMode: this.displayMode === DisplayMode.Edit,
      },
      gatedElement
    );

    // SPDebug — toolkit's debug runtime + lazy-loaded panel. See SpSearchBox
    // for the per-web-part-state-isolation note.
    const element: React.ReactElement = React.createElement(
      SPDebugProvider,
      { logger: SPContext.logger, allowInProduction: false },
      bannerWrapped
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // Inject PnP property controls CSS fix (SPFx 1.22 re-hashes .module.css class names)
    ensurePnpPropertyControlStyles();

    try {
      // Cast needed: spfx-toolkit uses SPFx 1.21.1 types; this project uses 1.22.2
      await SPContext.basic(this.context as unknown as Parameters<typeof SPContext.basic>[0], 'SPSearchFilters');
      const contextId: string = this.properties.searchContextId || 'default';
      this._store = getStore(contextId);
      // T3.D1 — refcount holder.
      incrementContextRef(contextId);
      // T3.D10 — record this web part's registration. If it arrives
      // after Results' first search ran with empty filterConfig,
      // the diagnostic flips `filtersLateRegistered` and the Results
      // edit-mode MessageBar surfaces.
      recordWebPartInit(contextId, 'SpSearchFiltersWebPart');

      // Register the SharePoint Search data provider (idempotent — skips if already registered by another web part)
      const provider = new SharePointSearchProvider();
      const dataProviders = this._store.getState().registries.dataProviders;
      if (!dataProviders.get(provider.id)) {
        dataProviders.register(provider);
      }

      // Register all built-in filter types (checkbox, daterange, toggle, tagbox, slider, taxonomy, people)
      const filterTypes = this._store.getState().registries.filterTypes;
      registerBuiltInFilterTypes(filterTypes);

      // Sync filter configuration and operator to the store BEFORE initializing context,
      // so the first search (triggered by initializeSearchContext) includes refiner properties
      this._syncFilterConfigToStore();
      this._syncOperatorToStore();

      // Initialize the shared search context (ensures library bundle's SPContext is ready)
      // Idempotent — if already initialized by another web part, this is a no-op
      await initializeSearchContext(contextId, this.context);
      DebugCollector.registerWebPart('SPSearchFiltersWebPart', this.properties as unknown as Record<string, unknown>);
    } catch (err) {
      console.error('[SPSearchFilters] onInit failed:', err);
      throw err;
    }
  }

  private _syncFilterConfigToStore(): void {
    if (!this._store) {
      return;
    }

    const rawFilters = normalizeFiltersCollectionValue(this.properties.filtersCollection);

    if (!rawFilters || rawFilters.length === 0) {
      const state: ISearchStore = this._store.getState();
      if (state.filterConfig.length > 0) {
        this._store.setState({ filterConfig: [], activeFilters: [] });
      }
      return;
    }

    const filterConfig: IFilterConfig[] = rawFilters.map((item: IFilterCollectionItem) => {
      const filterType = (item.filterType || 'checkbox') as IFilterConfig['filterType'];

      return {
        id: item.uniqueId,
        managedProperty: normalizeManagedPropertyForFilter(item.managedProperty, filterType),
        displayName: item.displayName,
        urlAlias: sanitizeUrlAlias(item.urlAlias),
        filterType,
        operator: (item.operator || 'OR') as IFilterConfig['operator'],
        maxValues: item.maxValues || 10,
        defaultExpanded: item.defaultExpanded !== false,
        showCount: item.showCount !== false,
        sortBy: (item.sortBy || 'count') as IFilterConfig['sortBy'],
        sortDirection: (item.sortDirection || 'desc') as IFilterConfig['sortDirection'],
        multiValues: item.multiValues !== false,
        dependsOn: item.dependsOn || undefined,
        showWhenParentHasValue: item.showWhenParentHasValue === true,
        hideZeroCountValues: item.hideZeroCountValues === true,
        resetWhenParentChanges: item.resetWhenParentChanges === true,
        trueLabel: item.trueLabel || undefined,
        falseLabel: item.falseLabel || undefined,
        invertBoolean: item.invertBoolean === true,
        audienceGroups: item.audience
          ? item.audience.split(',').map((s: string) => s.trim()).filter(Boolean)
          : undefined
      };
    });

    const state: ISearchStore = this._store.getState();
    if (JSON.stringify(state.filterConfig) !== JSON.stringify(filterConfig)) {
      this._store.setState({ filterConfig });
    }
  }

  private _syncOperatorToStore(): void {
    if (!this._store) {
      return;
    }
    const operator: 'AND' | 'OR' = (this.properties.operatorBetweenFilters === 'OR') ? 'OR' : 'AND';
    const state: ISearchStore = this._store.getState();
    if (state.operatorBetweenFilters !== operator) {
      state.setOperatorBetweenFilters(operator);
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
    }
  }

  protected onDispose(): void {
    // T3.D1 — drop refcount before unmounting React.
    const contextId: string = this.properties.searchContextId || 'default';
    decrementContextRef(contextId);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'searchContextId') {
      const contextId: string = this.properties.searchContextId || 'default';
      this._store = getStore(contextId);
    }

    if (propertyPath === 'filtersCollection' || propertyPath === 'searchContextId') {
      this._syncFilterConfigToStore();
    }

    if (propertyPath === 'operatorBetweenFilters' || propertyPath === 'searchContextId') {
      this._syncOperatorToStore();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.FiltersPageHeader
          },
          groups: [
            // T3.D4 — searchContextId is the first field every admin sees
            // on every search web part. Shared helper.
            {
              groupName: SEARCH_CONTEXT_ID_GROUP_NAME,
              groupFields: [
                propertyPaneSearchContextIdField()
              ]
            },
            {
              groupName: strings.FiltersGroupName,
              groupFields: [
                propertyPaneGroupHelp('filters-config', 'Help: Configure refiners and filter types'),
                PropertyFieldCollectionData('filtersCollection', {
                  key: 'filtersCollection',
                  label: strings.FiltersFieldLabel,
                  panelHeader: strings.FiltersPanelHeader,
                  manageBtnLabel: strings.FiltersManageBtn,
                  value: normalizeFiltersCollectionValue(this.properties.filtersCollection),
                  enableSorting: true,
                  fields: [
                    {
                      id: 'managedProperty',
                      title: strings.FilterPropertyColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'RefinableString00'
                    },
                    {
                      id: 'displayName',
                      title: strings.FilterDisplayNameColumn,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'File Type'
                    },
                    {
                      id: 'urlAlias',
                      title: strings.FilterUrlAliasColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'ft'
                    },
                    {
                      id: 'filterType',
                      title: strings.FilterTypeColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'checkbox', text: strings.FilterTypeCheckbox },
                        { key: 'dropdown', text: strings.FilterTypeDropdown },
                        { key: 'daterange', text: strings.FilterTypeDateRange },
                        { key: 'text', text: strings.FilterTypeText },
                        { key: 'people', text: strings.FilterTypePeople },
                        { key: 'taxonomy', text: strings.FilterTypeTaxonomy },
                        { key: 'slider', text: strings.FilterTypeSlider },
                        { key: 'tagbox', text: strings.FilterTypeTagBox },
                        { key: 'toggle', text: strings.FilterTypeToggle }
                      ]
                    },
                    {
                      id: 'operator',
                      title: strings.FilterOperatorColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      options: [
                        { key: 'OR', text: 'OR' },
                        { key: 'AND', text: 'AND' }
                      ]
                    },
                    {
                      id: 'maxValues',
                      title: strings.FilterMaxValuesColumn,
                      type: CustomCollectionFieldType.number,
                      required: false,
                      defaultValue: 10
                    },
                    {
                      id: 'showCount',
                      title: strings.FilterShowCountColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: true
                    },
                    {
                      id: 'defaultExpanded',
                      title: strings.FilterExpandedColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: true
                    },
                    {
                      id: 'sortBy',
                      title: strings.FilterSortByColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      options: [
                        { key: 'count', text: strings.FilterSortByCount },
                        { key: 'alphabetical', text: strings.FilterSortByName }
                      ]
                    },
                    {
                      id: 'sortDirection',
                      title: strings.FilterSortDirectionColumn,
                      type: CustomCollectionFieldType.dropdown,
                      required: false,
                      options: [
                        { key: 'desc', text: strings.FilterSortDescending },
                        { key: 'asc', text: strings.FilterSortAscending }
                      ]
                    },
                    {
                      id: 'multiValues',
                      title: strings.FilterMultiValuesColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: true
                    },
                    {
                      id: 'dependsOn',
                      title: strings.FilterDependsOnColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'ContentType'
                    },
                    {
                      id: 'showWhenParentHasValue',
                      title: strings.FilterShowWhenParentHasValueColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: false
                    },
                    {
                      id: 'hideZeroCountValues',
                      title: strings.FilterHideZeroCountValuesColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: false
                    },
                    {
                      id: 'resetWhenParentChanges',
                      title: strings.FilterResetWhenParentChangesColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: false
                    },
                    {
                      id: 'trueLabel',
                      title: strings.FilterTrueLabelColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'Yes'
                    },
                    {
                      id: 'falseLabel',
                      title: strings.FilterFalseLabelColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'No'
                    },
                    {
                      id: 'invertBoolean',
                      title: strings.FilterInvertBooleanColumn,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: false
                    },
                    // Stream D / #5 — audience targeting per refiner. Matches
                    // the convention used by `SpSearchVerticalsWebPart.ts:305`
                    // (comma-separated Azure AD group object IDs).
                    {
                      id: 'audience',
                      title: strings.FilterAudienceColumn,
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: strings.FilterAudiencePlaceholder
                    }
                  ]
                }),
              ]
            },
            {
              groupName: strings.BehaviorGroupName,
              groupFields: [
                propertyPaneGroupHelp('filters-behavior', 'Help: Apply mode and Clear All behaviour'),
                PropertyPaneChoiceGroup('applyMode', {
                  label: strings.ApplyModeLabel,
                  options: [
                    { key: 'instant', text: strings.ApplyModeInstant },
                    { key: 'manual', text: strings.ApplyModeManual }
                  ]
                }),
                PropertyPaneChoiceGroup('operatorBetweenFilters', {
                  label: strings.OperatorLabel,
                  options: [
                    { key: 'AND', text: 'AND' },
                    { key: 'OR', text: 'OR' }
                  ]
                }),
                PropertyPaneToggle('showClearAll', {
                  label: strings.ShowClearAllLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            }
          ]
        },
        // T3.D4 — Connections page retired: searchContextId is now on
        // page-1 group-1 via the shared helper, leaving no fields for a
        // separate page.
        {
          header: {
            description: strings.AdvancedPageHeader
          },
          groups: [
            {
              groupName: strings.AdvancedGroupName,
              groupFields: [
                PropertyPaneToggle('enableVisualFilterBuilder', {
                  label: strings.EnableVisualFilterBuilderLabel,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText
                })
              ]
            },
            // Stream D / #10 — per-web-part audience targeting.
            {
              groupName: strings.AudienceTargetingGroupName,
              groupFields: [
                PropertyPaneTextField('audienceGroups', {
                  label: strings.AudienceTargetingLabel,
                  description: strings.AudienceTargetingDescription,
                  multiline: true,
                  rows: 3,
                  resizable: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
