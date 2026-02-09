declare interface ISpSearchBoxWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ConnectionGroupName: string;
  InputGroupName: string;
  QueryGroupName: string;
  NavigationGroupName: string;
  FeaturesGroupName: string;
  SearchBoxPageHeader: string;
  FeaturesPageHeader: string;
  SearchContextIdFieldLabel: string;
  SearchContextIdFieldDescription: string;
  PlaceholderFieldLabel: string;
  DebounceMsFieldLabel: string;
  SearchBehaviorFieldLabel: string;
  SearchBehaviorOnEnter: string;
  SearchBehaviorOnButton: string;
  SearchBehaviorBoth: string;
  QueryInputTransformationLabel: string;
  QueryInputTransformationDescription: string;
  SearchInNewPageLabel: string;
  NewPageUrlLabel: string;
  NewPageUrlDescription: string;
  EnableScopeSelectorFieldLabel: string;
  EnableSuggestionsFieldLabel: string;
  EnableQueryBuilderFieldLabel: string;
  EnableSearchManagerFieldLabel: string;
  ToggleOnText: string;
  ToggleOffText: string;
}

declare module 'SpSearchBoxWebPartStrings' {
  const strings: ISpSearchBoxWebPartStrings;
  export = strings;
}
