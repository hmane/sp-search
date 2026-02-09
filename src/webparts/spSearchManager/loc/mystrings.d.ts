declare interface ISpSearchManagerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ConnectionGroupName: string;
  FeaturesGroupName: string;
  SearchContextIdLabel: string;
  SearchContextIdDescription: string;
  DisplayModeLabel: string;
  ModeStandalone: string;
  ModePanel: string;
  EnableSavedSearchesLabel: string;
  EnableSharedSearchesLabel: string;
  EnableCollectionsLabel: string;
  EnableHistoryLabel: string;
  EnableAnnotationsLabel: string;
  MaxHistoryItemsLabel: string;
  MaxHistoryItemsDescription: string;
  ToggleOnText: string;
  ToggleOffText: string;
}

declare module 'SpSearchManagerWebPartStrings' {
  const strings: ISpSearchManagerWebPartStrings;
  export = strings;
}
