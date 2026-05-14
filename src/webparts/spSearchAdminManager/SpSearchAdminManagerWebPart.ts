import { DisplayMode } from '@microsoft/sp-core-library';
import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import SpSearchManagerWebPart from '../spSearchManager/SpSearchManagerWebPart';
import type { ISpSearchManagerProps } from '../spSearchManager/components/ISpSearchManagerProps';
import { DebugCollector } from '@store/debug';

/**
 * SpSearchAdminManagerWebPart — standalone admin-only web part for search
 * health, freshness, coverage monitoring, and insights dashboards.
 *
 * T4.D6 (Path B fork) — this subclass is the *only* surface that renders
 * the admin variant (`variant='admin'`, admin tabs enabled, user tabs
 * force-disabled). The base SpSearchManagerWebPart is the user-facing
 * Manager (saved/shared/collections/history); the two web parts each have
 * their own manifest, icon, and property pane shape.
 *
 * `BaseSPSearchManagerWebPart` enforces the ManageWeb permission check;
 * the base render path short-circuits to the access-denied panel when
 * `_hasAdminAccess` is false AND the variant is 'admin'. (The user
 * variant doesn't gate on ManageWeb.)
 */
export default class SpSearchAdminManagerWebPart extends SpSearchManagerWebPart {
  protected async onInit(): Promise<void> {
    await super.onInit();
    DebugCollector.registerWebPart('SPSearchAdminManagerWebPart', this.properties as unknown as Record<string, unknown>);
  }

  /** T3.D2 — surface name shown in the mismatch banner. */
  protected _getWebPartLabel(): string {
    return 'SP Search Admin Manager';
  }

  /** T4.D6 — admin variant of the rendered surface. */
  protected _getVariant(): 'user' | 'admin' {
    return 'admin';
  }

  /** T4.D6 — admin-only header copy in the property pane. */
  protected _getPropertyPaneHeaderDescription(): string {
    return 'Configure admin diagnostics: coverage profiles, expected sites, and admin-only tabs (coverage / health / insights). Requires ManageWeb (Owner/Admin) permission to render.';
  }

  /**
   * T4.D6 — admin property pane: coverage profiles, expected sites,
   * connection (coverageSourcePageUrl), defaultTab over admin tabs,
   * admin section toggles. NO user-facing toggles.
   */
  protected _buildPropertyPaneGroups(): ReturnType<SpSearchManagerWebPart['_buildPropertyPaneGroups']> {
    return this._buildAdminPropertyPaneGroups();
  }

  /**
   * T4.D6 — projects admin-variant props into `<SpSearchManager>`. User
   * tabs forced false; admin tabs read from `this.properties` (manifest
   * defaults). Matches the pre-fork render block at the original
   * SpSearchManagerWebPart.ts:107.
   */
  protected _buildManagerProps(): ISpSearchManagerProps {
    return {
      store: this._getStore(),
      service: this._getService(),
      theme: this._getTheme(),
      variant: 'admin',
      searchContextId: this.properties.searchContextId || 'default',
      mode: 'standalone',
      defaultTab: this.properties.defaultTab || 'coverage',
      headerTitle: 'Admin Search Manager',
      context: this.context,
      // User tabs — explicitly off for the admin variant.
      enableSavedSearches: false,
      enableSharedSearches: false,
      enableCollections: false,
      enableHistory: false,
      enableAnnotations: false,
      maxHistoryItems: 0,
      showResetAction: false,
      showSaveAction: false,
      // Admin tabs — projected from properties + manifest defaults.
      enableCoverage: this.properties.enableCoverage !== false,
      coverageSourcePageUrl: this.properties.coverageSourcePageUrl || '',
      coverageProfiles: this._normalizeCoverageProfiles(),
      enableHealth: this.properties.enableHealth !== false,
      enableInsights: this.properties.enableInsights !== false,
      enableDashboard: this.properties.enableDashboard,
      expectedSiteUrls: (this.properties.expectedSiteUrls || '').split('\n').map((s: string) => s.trim()).filter(Boolean),
      // T4.D5 — edit-mode validators fire for the admin variant since
      // both coverageProfiles + expectedSiteUrls are exposed here.
      isEditMode: this.displayMode === DisplayMode.Edit,
      tenantRoot: this.context.pageContext.web.absoluteUrl,
    };
  }

  // T4.D6 — narrow accessors so the subclass doesn't depend on the base
  // class's private fields. Same shape the original render block used.
  private _getStore(): StoreApi<ISearchStore> {
    return (this as unknown as { _store: StoreApi<ISearchStore> })._store;
  }

  private _getService(): SearchManagerService {
    return (this as unknown as { _service: SearchManagerService })._service;
  }

  private _getTheme(): ISpSearchManagerProps['theme'] {
    return (this as unknown as { _theme: ISpSearchManagerProps['theme'] })._theme;
  }
}
