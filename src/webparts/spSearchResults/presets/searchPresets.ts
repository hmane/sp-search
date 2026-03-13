/**
 * Search scenario presets — centralized definitions for the Results web part
 * "Quick Start" property pane picker.
 *
 * Each preset configures every field the Results web part owns directly.
 * Fields owned by the Filters / Verticals web parts are listed in
 * `filterSuggestions` / `verticalSuggestions` as admin guidance only —
 * they are surfaced as hints in the property pane label, not auto-applied.
 */

// ─── Interfaces ───────────────────────────────────────────────────────────────

export interface IPresetProperty {
  property: string;
  alias: string;
}

export interface IPresetSortField {
  property: string;
  label: string;
  direction: 'Ascending' | 'Descending';
}

export interface IPresetFilterSuggestion {
  /** Managed property to offer as a filter */
  managedProperty: string;
  /** Suggested filter label */
  label: string;
  /** Short public URL alias for this filter */
  urlAlias?: string;
  /** Suggested filter type key */
  filterType: string;
}

export interface IScenarioPreset {
  id: string;
  label: string;
  description: string;
  /** Icon name from Fluent UI / Office Fabric icon set */
  iconName: string;
  // ── Results web part owned ────────────────────────────────────
  queryTemplate: string;
  defaultLayout: string;
  showListLayout: boolean;
  showCompactLayout: boolean;
  showGridLayout: boolean;
  showCardLayout: boolean;
  showPeopleLayout: boolean;
  showGalleryLayout: boolean;
  selectedProperties: IPresetProperty[];
  compactProperties: IPresetProperty[];
  sortableProperties: IPresetSortField[];
  // ── Admin hints (not auto-applied) ───────────────────────────
  /** Suggested data provider ID — set via the Verticals web part per-vertical dropdown */
  dataProviderHint: string;
  /** Suggested filters — configure in the Filters web part */
  filterSuggestions: IPresetFilterSuggestion[];
}

// ─── Preset definitions ───────────────────────────────────────────────────────

/** General-purpose search — sensible defaults, all content types, flexible layouts. */
const GENERAL: IScenarioPreset = {
  id: 'general',
  label: 'General',
  description: 'Broad search across all content. List layout with flexible refiners.',
  iconName: 'Search',
  queryTemplate: '{searchTerms}',
  defaultLayout: 'list',
  showListLayout: true, showCompactLayout: true, showGridLayout: true,
  showCardLayout: false, showPeopleLayout: false, showGalleryLayout: false,
  selectedProperties: [
    { property: 'Title',            alias: 'Title'    },
    { property: 'Author',           alias: 'Author'   },
    { property: 'LastModifiedTime', alias: 'Modified' },
    { property: 'FileType',         alias: 'Type'     },
    { property: 'Size',             alias: 'Size'     },
    { property: 'Path',             alias: 'URL'      },
    { property: 'SiteName',         alias: 'Site'     },
  ],
  compactProperties: [
    { property: 'Author',           alias: 'Author'   },
    { property: 'LastModifiedTime', alias: 'Modified' },
    { property: 'Size',             alias: 'Size'     },
    { property: 'FileType',         alias: 'Type'     },
  ],
  sortableProperties: [
    { property: 'LastModifiedTime', label: 'Date Modified', direction: 'Descending' },
    { property: 'Title',            label: 'Title',         direction: 'Ascending'  },
  ],
  dataProviderHint: 'sharepoint-search',
  filterSuggestions: [
    { managedProperty: 'FileType',         label: 'File type',      urlAlias: 'ft', filterType: 'checkbox' },
    { managedProperty: 'LastModifiedTime', label: 'Modified date',  urlAlias: 'md', filterType: 'daterange' },
    { managedProperty: 'AuthorOWSUSER',    label: 'Author',         urlAlias: 'au', filterType: 'people'   },
  ],
};

/** Document-centric search — scoped to IsDocument:1, compact + grid, file-type refiners. */
const DOCUMENTS: IScenarioPreset = {
  id: 'documents',
  label: 'Documents',
  description: 'Scoped to SharePoint documents. Compact and grid layouts, file-type and date refiners.',
  iconName: 'DocLibrary',
  queryTemplate: '{searchTerms} IsDocument:1',
  defaultLayout: 'list',
  showListLayout: true, showCompactLayout: true, showGridLayout: true,
  showCardLayout: false, showPeopleLayout: false, showGalleryLayout: false,
  selectedProperties: [
    { property: 'Title',            alias: 'Title'    },
    { property: 'Author',           alias: 'Author'   },
    { property: 'LastModifiedTime', alias: 'Modified' },
    { property: 'FileType',         alias: 'Type'     },
    { property: 'Size',             alias: 'Size'     },
    { property: 'Path',             alias: 'URL'      },
    { property: 'SiteName',         alias: 'Site'     },
    { property: 'HitHighlightedSummary', alias: 'Summary' },
  ],
  compactProperties: [
    { property: 'Author',           alias: 'Author'   },
    { property: 'LastModifiedTime', alias: 'Modified' },
    { property: 'Size',             alias: 'Size'     },
    { property: 'FileType',         alias: 'Type'     },
  ],
  sortableProperties: [
    { property: 'LastModifiedTime', label: 'Date Modified', direction: 'Descending' },
    { property: 'Title',            label: 'Title',         direction: 'Ascending'  },
    { property: 'Size',             label: 'File Size',     direction: 'Descending' },
    { property: 'Author',           label: 'Author',        direction: 'Ascending'  },
  ],
  dataProviderHint: 'sharepoint-search',
  filterSuggestions: [
    { managedProperty: 'FileType',         label: 'File type',     urlAlias: 'ft', filterType: 'checkbox' },
    { managedProperty: 'LastModifiedTime', label: 'Modified date', urlAlias: 'md', filterType: 'daterange' },
    { managedProperty: 'AuthorOWSUSER',    label: 'Author',        urlAlias: 'au', filterType: 'people'   },
    { managedProperty: 'SiteName',         label: 'Site',          urlAlias: 'si', filterType: 'checkbox' },
  ],
};

/** People directory search — Graph People provider, People layout, profile-rich properties. */
const PEOPLE: IScenarioPreset = {
  id: 'people',
  label: 'People',
  description: 'People directory via Microsoft Graph. Requires graph-people provider in a vertical.',
  iconName: 'Group',
  queryTemplate: '{searchTerms}',
  defaultLayout: 'people',
  showListLayout: false, showCompactLayout: false, showGridLayout: false,
  showCardLayout: false, showPeopleLayout: true, showGalleryLayout: false,
  selectedProperties: [
    { property: 'Title',            alias: 'Name'       },
    { property: 'WorkEmail',        alias: 'Email'      },
    { property: 'JobTitle',         alias: 'Job Title'  },
    { property: 'Department',       alias: 'Department' },
    { property: 'OfficeNumber',     alias: 'Office'     },
    { property: 'WorkPhone',        alias: 'Phone'      },
    { property: 'SPS-Skills',       alias: 'Skills'     },
    { property: 'AboutMe',          alias: 'About Me'   },
    { property: 'PictureURL',       alias: 'Photo'      },
  ],
  compactProperties: [
    { property: 'JobTitle',     alias: 'Job Title'  },
    { property: 'Department',   alias: 'Department' },
    { property: 'OfficeNumber', alias: 'Office'     },
    { property: 'WorkPhone',    alias: 'Phone'      },
  ],
  sortableProperties: [
    { property: 'Title',      label: 'Name',       direction: 'Ascending' },
    { property: 'Department', label: 'Department', direction: 'Ascending' },
  ],
  dataProviderHint: 'graph-people',
  filterSuggestions: [
    { managedProperty: 'Department',   label: 'Department', urlAlias: 'dp', filterType: 'checkbox' },
    { managedProperty: 'JobTitle',     label: 'Job title',  urlAlias: 'jt', filterType: 'checkbox' },
    { managedProperty: 'OfficeNumber', label: 'Office',     urlAlias: 'of', filterType: 'checkbox' },
  ],
};

/** News feed — scoped to modern pages (PromotedState:2), Card layout, date-sorted. */
const NEWS: IScenarioPreset = {
  id: 'news',
  label: 'News',
  description: 'SharePoint News pages (PromotedState:2). Card layout, sorted by published date.',
  iconName: 'News',
  queryTemplate: '{searchTerms} PromotedState:2',
  defaultLayout: 'card',
  showListLayout: true, showCompactLayout: false, showGridLayout: false,
  showCardLayout: true, showPeopleLayout: false, showGalleryLayout: false,
  selectedProperties: [
    { property: 'Title',                 alias: 'Title'       },
    { property: 'Author',                alias: 'Author'      },
    { property: 'Created',               alias: 'Published'   },
    { property: 'PictureThumbnailURL',   alias: 'Thumbnail'   },
    { property: 'HitHighlightedSummary', alias: 'Description' },
    { property: 'SiteName',              alias: 'Site'        },
    { property: 'Path',                  alias: 'URL'         },
  ],
  compactProperties: [
    { property: 'Author',  alias: 'Author'    },
    { property: 'Created', alias: 'Published' },
    { property: 'SiteName', alias: 'Site'     },
  ],
  sortableProperties: [
    { property: 'Created', label: 'Published', direction: 'Descending' },
    { property: 'Title',   label: 'Title',     direction: 'Ascending'  },
  ],
  dataProviderHint: 'sharepoint-search',
  filterSuggestions: [
    { managedProperty: 'Created',        label: 'Published date', urlAlias: 'pd', filterType: 'daterange' },
    { managedProperty: 'AuthorOWSUSER',  label: 'Author',         urlAlias: 'au', filterType: 'people'   },
    { managedProperty: 'SiteName',       label: 'Site',           urlAlias: 'si', filterType: 'checkbox' },
  ],
};

/** Hub intranet search — all content across associated sites, document/news/page verticals. */
const HUB_SEARCH: IScenarioPreset = {
  id: 'hub-search',
  label: 'Hub Search',
  description: 'Cross-site intranet search scoped to a hub. All content types with document, news, and page verticals.',
  iconName: 'Globe',
  queryTemplate: '{searchTerms}',
  defaultLayout: 'list',
  showListLayout: true, showCompactLayout: true, showGridLayout: true,
  showCardLayout: true, showPeopleLayout: false, showGalleryLayout: false,
  selectedProperties: [
    { property: 'Title',                 alias: 'Title'    },
    { property: 'Author',                alias: 'Author'   },
    { property: 'LastModifiedTime',      alias: 'Modified' },
    { property: 'FileType',              alias: 'Type'     },
    { property: 'SiteName',              alias: 'Site'     },
    { property: 'HitHighlightedSummary', alias: 'Summary'  },
    { property: 'Path',                  alias: 'URL'      },
  ],
  compactProperties: [
    { property: 'Author',           alias: 'Author'   },
    { property: 'LastModifiedTime', alias: 'Modified' },
    { property: 'SiteName',         alias: 'Site'     },
    { property: 'FileType',         alias: 'Type'     },
  ],
  sortableProperties: [
    { property: 'LastModifiedTime', label: 'Date Modified', direction: 'Descending' },
    { property: 'Title',            label: 'Title',         direction: 'Ascending'  },
  ],
  dataProviderHint: 'sharepoint-search',
  filterSuggestions: [
    { managedProperty: 'FileType',         label: 'Content type',  urlAlias: 'ft', filterType: 'checkbox' },
    { managedProperty: 'SiteName',         label: 'Site',          urlAlias: 'si', filterType: 'checkbox' },
    { managedProperty: 'AuthorOWSUSER',    label: 'Author',        urlAlias: 'au', filterType: 'people'   },
    { managedProperty: 'LastModifiedTime', label: 'Modified date', urlAlias: 'md', filterType: 'daterange' },
  ],
};

/** Knowledge base — articles, how-to guides, and reference docs. Card layout, rich preview. */
const KNOWLEDGE_BASE: IScenarioPreset = {
  id: 'knowledge-base',
  label: 'Knowledge Base',
  description: 'Knowledge articles, how-to guides, and reference documents. Card layout with rich preview.',
  iconName: 'BookAnswers',
  queryTemplate: '{searchTerms} (IsDocument:1 OR contentclass:STS_ListItem_Pages)',
  defaultLayout: 'card',
  showListLayout: true, showCompactLayout: false, showGridLayout: true,
  showCardLayout: true, showPeopleLayout: false, showGalleryLayout: false,
  selectedProperties: [
    { property: 'Title',                 alias: 'Title'     },
    { property: 'Author',                alias: 'Author'    },
    { property: 'Created',               alias: 'Published' },
    { property: 'HitHighlightedSummary', alias: 'Summary'   },
    { property: 'ContentType',           alias: 'Category'  },
    { property: 'SiteName',              alias: 'Site'      },
    { property: 'PictureThumbnailURL',   alias: 'Thumbnail' },
    { property: 'Path',                  alias: 'URL'       },
  ],
  compactProperties: [
    { property: 'ContentType', alias: 'Category'  },
    { property: 'Created',     alias: 'Published' },
    { property: 'SiteName',    alias: 'Site'      },
  ],
  sortableProperties: [
    { property: 'Created',          label: 'Published',    direction: 'Descending' },
    { property: 'LastModifiedTime', label: 'Last Updated', direction: 'Descending' },
    { property: 'Title',            label: 'Title',        direction: 'Ascending'  },
  ],
  dataProviderHint: 'sharepoint-search',
  filterSuggestions: [
    { managedProperty: 'ContentType',      label: 'Category',       urlAlias: 'ct', filterType: 'checkbox' },
    { managedProperty: 'SiteName',         label: 'Site',           urlAlias: 'si', filterType: 'checkbox' },
    { managedProperty: 'AuthorOWSUSER',    label: 'Author',         urlAlias: 'au', filterType: 'people'   },
    { managedProperty: 'Created',          label: 'Published date', urlAlias: 'pd', filterType: 'daterange' },
  ],
};

/** Policy search — corporate policies, procedures, and compliance documents. */
const POLICY_SEARCH: IScenarioPreset = {
  id: 'policy-search',
  label: 'Policy Search',
  description: 'Corporate policies, procedures, and compliance documents. Scoped to PDF and Office files.',
  iconName: 'Shield',
  queryTemplate: '{searchTerms} IsDocument:1 (FileType:pdf OR FileType:docx OR FileType:doc OR FileType:xlsx)',
  defaultLayout: 'list',
  showListLayout: true, showCompactLayout: true, showGridLayout: true,
  showCardLayout: false, showPeopleLayout: false, showGalleryLayout: false,
  selectedProperties: [
    { property: 'Title',                 alias: 'Title'    },
    { property: 'Author',                alias: 'Owner'    },
    { property: 'LastModifiedTime',      alias: 'Reviewed' },
    { property: 'FileType',              alias: 'Type'     },
    { property: 'Size',                  alias: 'Size'     },
    { property: 'SiteName',              alias: 'Source'   },
    { property: 'HitHighlightedSummary', alias: 'Summary'  },
    { property: 'Path',                  alias: 'URL'      },
  ],
  compactProperties: [
    { property: 'Author',           alias: 'Owner'    },
    { property: 'LastModifiedTime', alias: 'Reviewed' },
    { property: 'FileType',         alias: 'Type'     },
    { property: 'Size',             alias: 'Size'     },
  ],
  sortableProperties: [
    { property: 'Title',            label: 'Title',         direction: 'Ascending'  },
    { property: 'LastModifiedTime', label: 'Last Reviewed', direction: 'Descending' },
    { property: 'Author',           label: 'Owner',         direction: 'Ascending'  },
  ],
  dataProviderHint: 'sharepoint-search',
  filterSuggestions: [
    { managedProperty: 'FileType',         label: 'File type',     urlAlias: 'ft', filterType: 'checkbox' },
    { managedProperty: 'SiteName',         label: 'Source',        urlAlias: 'si', filterType: 'checkbox' },
    { managedProperty: 'AuthorOWSUSER',    label: 'Policy owner',  urlAlias: 'au', filterType: 'people'   },
    { managedProperty: 'LastModifiedTime', label: 'Last reviewed', urlAlias: 'md', filterType: 'daterange' },
  ],
};

/** Media gallery — image and video assets, Gallery + Card layouts, thumbnail-optimised. */
const MEDIA: IScenarioPreset = {
  id: 'media',
  label: 'Media',
  description: 'Images and video files. Gallery layout with thumbnail optimisation.',
  iconName: 'Photo2',
  queryTemplate: '{searchTerms} (FileType:jpg OR FileType:jpeg OR FileType:png OR FileType:gif OR FileType:mp4 OR FileType:mov)',
  defaultLayout: 'gallery',
  showListLayout: false, showCompactLayout: false, showGridLayout: false,
  showCardLayout: true, showPeopleLayout: false, showGalleryLayout: true,
  selectedProperties: [
    { property: 'Title',               alias: 'Title'    },
    { property: 'PictureThumbnailURL', alias: 'Thumbnail'},
    { property: 'FileType',            alias: 'Type'     },
    { property: 'Size',                alias: 'Size'     },
    { property: 'LastModifiedTime',    alias: 'Modified' },
    { property: 'Author',              alias: 'Author'   },
    { property: 'SiteName',            alias: 'Site'     },
    { property: 'Path',                alias: 'URL'      },
  ],
  compactProperties: [
    { property: 'FileType',         alias: 'Type'     },
    { property: 'LastModifiedTime', alias: 'Modified' },
    { property: 'Size',             alias: 'Size'     },
    { property: 'SiteName',         alias: 'Site'     },
  ],
  sortableProperties: [
    { property: 'LastModifiedTime', label: 'Date Modified', direction: 'Descending' },
    { property: 'Title',            label: 'Title',         direction: 'Ascending'  },
    { property: 'Size',             label: 'File Size',     direction: 'Descending' },
  ],
  dataProviderHint: 'sharepoint-search',
  filterSuggestions: [
    { managedProperty: 'FileType',         label: 'File type',     urlAlias: 'ft', filterType: 'checkbox' },
    { managedProperty: 'LastModifiedTime', label: 'Modified date', urlAlias: 'md', filterType: 'daterange' },
    { managedProperty: 'SiteName',         label: 'Site',          urlAlias: 'si', filterType: 'checkbox' },
  ],
};

// ─── Registry ─────────────────────────────────────────────────────────────────

/** All built-in scenario presets, keyed by preset ID. */
export const SCENARIO_PRESETS: Record<string, IScenarioPreset> = {
  general:        GENERAL,
  documents:      DOCUMENTS,
  people:         PEOPLE,
  news:           NEWS,
  media:          MEDIA,
  'hub-search':   HUB_SEARCH,
  'knowledge-base': KNOWLEDGE_BASE,
  'policy-search':  POLICY_SEARCH,
};
