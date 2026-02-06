declare interface ISpSearchBoxWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  SearchContextIdFieldLabel: string;
  SearchContextIdFieldDescription: string;
  PlaceholderFieldLabel: string;
  DebounceMsFieldLabel: string;
  SearchBehaviorFieldLabel: string;
  SearchBehaviorOnEnter: string;
  SearchBehaviorOnButton: string;
  SearchBehaviorBoth: string;
  EnableScopeSelectorFieldLabel: string;
  EnableSuggestionsFieldLabel: string;
}

declare module 'SpSearchBoxWebPartStrings' {
  const strings: ISpSearchBoxWebPartStrings;
  export = strings;
}
