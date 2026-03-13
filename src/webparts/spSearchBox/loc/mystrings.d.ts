declare interface ISpSearchBoxWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  SearchGroupName: string;
  SuggestionsGroupName: string;
  ConnectionGroupName: string;
  ScopeGroupName: string;
  AdvancedGroupName: string;
  QueryGroupName: string;
  NavigationGroupName: string;
  SearchPageHeader: string;
  SuggestionsPageHeader: string;
  ConnectionsPageHeader: string;
  AdvancedPageHeader: string;
  SearchContextIdFieldLabel: string;
  SearchContextIdFieldDescription: string;
  PlaceholderFieldLabel: string;
  DebounceMsFieldLabel: string;
  SearchBehaviorFieldLabel: string;
  SearchBehaviorOnEnter: string;
  SearchBehaviorOnButton: string;
  SearchBehaviorBoth: string;
  ResetSearchOnClearLabel: string;
  QueryInputTransformationLabel: string;
  QueryInputTransformationDescription: string;
  SearchInNewPageLabel: string;
  NewPageUrlLabel: string;
  NewPageUrlDescription: string;
  NewPageUrlRequiredMessage: string;
  NewPageOpenBehaviorLabel: string;
  NewPageOpenBehaviorSameTab: string;
  NewPageOpenBehaviorNewTab: string;
  NewPageParameterLocationLabel: string;
  NewPageParameterLocationQueryString: string;
  NewPageParameterLocationHash: string;
  NewPageQueryParameterLabel: string;
  NewPageQueryParameterDescription: string;
  EnableScopeSelectorFieldLabel: string;
  ScopeInfoLabel: string;
  EnableSuggestionsFieldLabel: string;
  SuggestionsPerGroupLabel: string;
  EnableSharePointSuggestionsLabel: string;
  EnableRecentSuggestionsLabel: string;
  EnablePopularSuggestionsLabel: string;
  EnableQuickResultsLabel: string;
  EnablePropertySuggestionsLabel: string;
  EnableQueryBuilderFieldLabel: string;
  EnableKqlModeFieldLabel: string;
  EnableSearchManagerFieldLabel: string;
  ToggleOnText: string;
  ToggleOffText: string;
}

declare module 'SpSearchBoxWebPartStrings' {
  const strings: ISpSearchBoxWebPartStrings;
  export = strings;
}
