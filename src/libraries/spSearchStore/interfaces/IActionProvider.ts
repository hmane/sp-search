import { ISearchResult } from './ISearchResult';
import { ISearchContext } from './ISuggestionProvider';

/**
 * Action provider — registered actions appear as quick actions
 * on results and in the bulk actions toolbar.
 *
 * Built-in: OpenAction, PreviewAction, ShareAction, PinAction,
 *           CopyLinkAction, DownloadAction, CompareAction, ExportCsvAction
 */
export interface IActionProvider {
  id: string;
  label: string;
  /** Fluent UI icon name */
  iconName: string;
  position: 'toolbar' | 'contextMenu' | 'both';
  /** Conditional visibility — e.g. only show "Open in AutoCAD" for .dwg files */
  isApplicable: (item: ISearchResult) => boolean;
  /** Execute the action on one or more selected items */
  execute: (items: ISearchResult[], context: ISearchContext) => Promise<void>;
  /** Whether the action supports multi-select */
  isBulkEnabled: boolean;
}
