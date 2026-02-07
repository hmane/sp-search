declare interface ISpSearchResultsWebPartStrings {
  PropertyPaneDescription: string;
  DataGroupName: string;
  DisplayGroupName: string;
  SearchContextIdLabel: string;
  SearchContextIdDescription: string;
  QueryTemplateLabel: string;
  QueryTemplateDescription: string;
  SelectedPropertiesLabel: string;
  SelectedPropertiesDescription: string;
  PageSizeLabel: string;
  DefaultLayoutLabel: string;
  ListLayoutText: string;
  CompactLayoutText: string;
  ShowResultCountLabel: string;
  ShowSortDropdownLabel: string;
  SortFieldLabel: string;
  SortPanelHeader: string;
  SortManageBtn: string;
  SortPropertyColumn: string;
  SortLabelColumn: string;
  SortDirectionColumn: string;
  SortAscending: string;
  SortDescending: string;
  EnableSelectionLabel: string;
  ToggleOnText: string;
  ToggleOffText: string;
}

declare module 'SpSearchResultsWebPartStrings' {
  const strings: ISpSearchResultsWebPartStrings;
  export = strings;
}
