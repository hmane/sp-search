declare interface ISpSearchManagerWebPartStrings {
  PropertyPaneDescription: string;
  ConnectionGroupName: string;
  DisplayGroupName: string;
  SectionsGroupName: string;
  SearchContextIdLabel: string;
  SearchContextIdDescription: string;
  CoverageSourcePageUrlLabel: string;
  CoverageSourcePageUrlDescription: string;
  DefaultTabLabel: string;
  DefaultTabCoverage: string;
  DefaultTabHealth: string;
  DefaultTabInsights: string;
  EnableCoverageLabel: string;
  EnableHealthLabel: string;
  EnableInsightsLabel: string;
  MonitoringGroupName: string;
  CoverageProfilesLabel: string;
  CoverageProfilesPanelHeader: string;
  CoverageProfilesManageButton: string;
  CoverageProfileTitleColumn: string;
  CoverageProfileDescriptionColumn: string;
  CoverageProfileSourceUrlsColumn: string;
  CoverageProfileContentTypesColumn: string;
  CoverageProfileExcludePathsColumn: string;
  CoverageProfileQueryTemplateColumn: string;
  CoverageProfileResultSourceColumn: string;
  CoverageProfileRefinementFiltersColumn: string;
  CoverageProfileIncludeFoldersColumn: string;
  CoverageProfileTrimDuplicatesColumn: string;
  AccessDeniedTitle: string;
  AccessDeniedDescription: string;
  ToggleOnText: string;
  ToggleOffText: string;
}

declare module 'SpSearchManagerWebPartStrings' {
  const strings: ISpSearchManagerWebPartStrings;
  export = strings;
}
