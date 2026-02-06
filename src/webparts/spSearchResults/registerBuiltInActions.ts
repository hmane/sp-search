import type { IRegistry, IActionProvider } from '@interfaces/index';
import {
  OpenAction,
  PreviewAction,
  ShareAction,
  PinAction,
  CopyLinkAction,
  DownloadAction,
  CompareAction,
  ExportCsvAction
} from '@providers/index';

/**
 * Register built-in action providers into the ActionProviderRegistry.
 * Called from SpSearchResultsWebPart.onInit().
 */
export function registerBuiltInActions(registry: IRegistry<IActionProvider>): void {
  registry.register(new OpenAction());
  registry.register(new PreviewAction());
  registry.register(new ShareAction());
  registry.register(new PinAction());
  registry.register(new CopyLinkAction());
  registry.register(new DownloadAction());
  registry.register(new CompareAction());
  registry.register(new ExportCsvAction());
}
