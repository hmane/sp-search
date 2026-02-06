declare interface ISpSearchFiltersWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  SearchContextIdLabel: string;
  SearchContextIdDescription: string;
  ApplyModeLabel: string;
  ApplyModeInstant: string;
  ApplyModeManual: string;
  OperatorLabel: string;
  ShowClearAllLabel: string;
  EnableVisualFilterBuilderLabel: string;
  ToggleOnText: string;
  ToggleOffText: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'SpSearchFiltersWebPartStrings' {
  const strings: ISpSearchFiltersWebPartStrings;
  export = strings;
}
