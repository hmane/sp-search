declare interface ISpSearchManagerWebPartStrings {
  PropertyPaneDescription: string;
  DisplayGroupName: string;
  SectionsGroupName: string;
  SearchContextIdLabel: string;
  SearchContextIdDescription: string;
  DefaultTabLabel: string;
  DefaultTabDashboard: string;
  DefaultTabHealth: string;
  DefaultTabInsights: string;
  EnableDashboardLabel: string;
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
