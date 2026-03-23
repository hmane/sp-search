import SpSearchManagerWebPart from '../spSearchManager/SpSearchManagerWebPart';
import { DebugCollector } from '@store/debug';

/**
 * SpSearchAdminManagerWebPart — standalone admin-only web part for search
 * health, freshness, coverage monitoring, and insights dashboards.
 *
 * Extends the base SpSearchManagerWebPart which already enforces
 * ManageWeb permission checks and renders with variant='admin'.
 * This subclass exists so that the Admin Manager has its own web part
 * identity (manifest ID, preconfigured properties) while reusing all
 * base-class logic. The manifest defaults disable user-facing tabs
 * (saved, shared, collections, history) and enable admin tabs
 * (coverage, health, insights).
 */
export default class SpSearchAdminManagerWebPart extends SpSearchManagerWebPart {
  protected async onInit(): Promise<void> {
    await super.onInit();
    DebugCollector.registerWebPart('SPSearchAdminManagerWebPart', this.properties as unknown as Record<string, unknown>);
  }
}
