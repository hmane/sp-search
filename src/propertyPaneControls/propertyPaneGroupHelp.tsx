/**
 * Context-sensitive help for property pane groups.
 *
 * SPFx does not expose a native per-group help surface, so this custom
 * property-pane field renders a lightweight Help button and opens a local
 * Fluent modal. Content is bundled with the solution instead of sending
 * admins to GitHub or another external documentation site.
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneCustomFieldProps,
  type IPropertyPaneField,
  PropertyPaneFieldType,
} from '@microsoft/sp-property-pane';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';

export interface IPropertyPaneHelpTopic {
  title: string;
  summary: string;
  bullets: string[];
  examples?: string[];
}

const HELP_TOPICS: Record<string, IPropertyPaneHelpTopic> = {
  'quick-start': {
    title: 'Quick Start Presets',
    summary: 'Presets configure a Results web part for a common search scenario in one step.',
    bullets: [
      'Use a preset when creating a new search page or when you want to reset the Results web part to a known scenario.',
      'A preset can change the default layout, enabled layouts, selected managed properties, grid columns, compact columns, sort options, query template, and filter suggestions.',
      'After applying a preset, you can still customize individual fields. Changing layout toggles or the default layout marks the configuration as custom.',
      'Preset filter suggestions are published for the Filters web part to consume; they do not automatically overwrite refiners without admin review.'
    ],
    examples: [
      'Documents: enables document-focused fields such as Author, Modified, File Type, Size, Path, and Site.',
      'People: expects a Graph-backed people vertical and enables People-oriented result rendering.',
      'Knowledge base or policy search: starts from a constrained query template and focused metadata columns.'
    ]
  },
  'results-data': {
    title: 'Search Scope And Managed Properties',
    summary: 'This group controls where Results searches and which managed properties are retrieved for layouts.',
    bullets: [
      'Search scope can target all indexed content, the current site, the current site collection, or a specific custom path.',
      'Query template controls the final KQL sent to the provider. Include {searchTerms} unless you intentionally want a browse-only fixed result set.',
      'Selected properties are the master list of managed properties available to result layouts, sorting, and column editors.',
      'Use retrievable managed properties. If a field is not retrievable in SharePoint Search, the renderer cannot display it even if the name is configured.',
      'The edit-mode validation messages flag common issues such as missing {searchTerms}, sparse grid columns, and invalid managed property names.'
    ],
    examples: [
      '{searchTerms} IsDocument:1 returns only documents matching the user query.',
      'Path:"https://contoso.sharepoint.com/sites/hr" {searchTerms} scopes results to one site.',
      'Title, Author, LastModifiedTime, FileType, Size, Path, and SiteName are safe starter display properties.'
    ]
  },
  'results-layouts': {
    title: 'Layouts And Presets',
    summary: 'Layouts decide how the same result set is presented to users.',
    bullets: [
      'List is the baseline layout and works well for general search.',
      'Compact is denser and uses a smaller set of metadata fields.',
      'Data Grid is for power users who need resizable columns, sorting, export, and a column chooser.',
      'Card, People, and Gallery layouts need the right selected properties to look complete.',
      'The Data Grid title column is fixed. Additional columns come from Data Grid Columns and can be always visible, on by default, or off by default.'
    ],
    examples: [
      'For legal or account documents, start with List plus Data Grid.',
      'For a people directory, enable People layout and a Graph people vertical.',
      'For images or media, use Gallery only when thumbnail properties are selected.'
    ]
  },
  'box-search': {
    title: 'Search Input Behaviour',
    summary: 'This group controls how the Search Box captures and submits user queries.',
    bullets: [
      'Placeholder text should describe what users can search, not how the control works.',
      'Search trigger can be Enter, button click, or both.',
      'Debounce affects live query updates and suggestions. Lower values feel faster but can create more activity.',
      'Query transformation lets admins wrap user input in a template before search execution.',
      'Use {searchTerms} in query transformation to preserve the user-entered query.'
    ],
    examples: [
      'Transformation: {searchTerms} IsDocument:1 limits box queries to documents.',
      'Placeholder: Search policies, procedures, and account documents.'
    ]
  },
  'box-navigation': {
    title: 'Same-Page Vs New-Page Navigation',
    summary: 'The Search Box can either update connected results on the current page or redirect to a search page.',
    bullets: [
      'Use same-page search when the Search Box and Results web part are on the same page and share a search context ID.',
      'Use new-page navigation when the Search Box is on a landing page or header and results live elsewhere.',
      'The target page URL can be server-relative or absolute.',
      'The query can be sent as a query-string parameter or a hash parameter.',
      'The target results page must have a Search Box or Results web part configured to read the same query parameter.'
    ],
    examples: [
      '/sites/search/SitePages/results.aspx?q=budget',
      '/sites/search/SitePages/results.aspx#q=budget'
    ]
  },
  'box-suggestions': {
    title: 'Search Suggestions And Quick Results',
    summary: 'Suggestions help users complete queries before submitting a search.',
    bullets: [
      'Recent searches are per-user and come from search history.',
      'Popular or frequent queries depend on stored history data.',
      'SharePoint suggestions call the search service for query suggestions.',
      'Property suggestions help power users discover KQL-style managed property filters.',
      'Quick results show a small preview of matching content before the full search is submitted.'
    ],
    examples: [
      'Enable recent searches for repeat work patterns.',
      'Enable property suggestions when users know fields such as FileType, Author, or Path.'
    ]
  },
  'filters-config': {
    title: 'Configure Refiners And Filter Types',
    summary: 'Refiners define which managed properties appear as user-facing filters.',
    bullets: [
      'Each refiner needs a managed property that SharePoint Search can return as a refinement bucket or queryable field.',
      'Checkbox and dropdown work well for small categorical fields.',
      'Tag Box is better for larger multi-select lists.',
      'People filters target individual user claims and are not intended for group values.',
      'Taxonomy filters use SharePoint taxonomy refinement tokens and can show term labels when a term set is configured.',
      'Dependencies let one refiner show or reset based on another refiner selection.',
      'Audience targeting on a refiner currently accepts only Entra ID group object IDs returned by Graph /me/memberOf.'
    ],
    examples: [
      'Document Type: Tag Box, OR, max values 50.',
      'AuthorOWSUSER: People, OR, users only.',
      'Account Name depends on Document Type and resets when the parent changes.'
    ]
  },
  'filters-behavior': {
    title: 'Apply Mode And Clear All Behaviour',
    summary: 'This group controls when refiner selections update the search results.',
    bullets: [
      'Instant mode applies every selection immediately and is best for simple pages.',
      'Manual mode lets users stage multiple selections and apply them together.',
      'Clear All removes every active filter in the connected search context.',
      'Logic between refiners controls how different refiner groups combine.',
      'AND between refiners narrows results. OR between refiners broadens across groups.',
      'The visual filter builder is an advanced admin/user surface for composing structured filter expressions.'
    ],
    examples: [
      'Use Instant mode for a compact left rail with a small number of refiners.',
      'Use Manual mode when filters are numerous or expensive and users commonly select several at once.'
    ]
  },
  'verticals-config': {
    title: 'Configure Vertical Tabs And Data Providers',
    summary: 'Verticals split one search experience into tabs such as All, Documents, Pages, People, or external links.',
    bullets: [
      'Each vertical has a stable key, display label, optional icon, and optional query template.',
      'A query template narrows only that vertical. Include {searchTerms} unless the tab is intentionally browse-only.',
      'A result source ID is SharePoint-search specific and ignored by Graph providers.',
      'Data provider ID can route a vertical to SharePoint Search, Graph Search, or Graph People.',
      'Link-only verticals navigate instead of searching and are excluded from count fan-out.',
      'Per-vertical audience targeting currently accepts only Entra ID group object IDs returned by Graph /me/memberOf.'
    ],
    examples: [
      'Documents: query template {searchTerms} IsDocument:1.',
      'People: dataProviderId graph-people and default layout people.',
      'External: mark as link and provide a target URL.'
    ]
  },
  'manager-user-tabs': {
    title: 'User-Facing Manager Tabs',
    summary: 'The Search Manager gives users a personal workspace for search activity.',
    bullets: [
      'Saved Searches stores reusable queries for the current user.',
      'Shared Searches surfaces searches shared with the user.',
      'Collections let users group useful results.',
      'History shows recent search activity when history logging is available.',
      'Only enable tabs that match the experience you want users to maintain.'
    ],
    examples: [
      'For a simple page, enable Saved Searches and History.',
      'For research workflows, also enable Collections and Shared Searches.'
    ]
  },
  'adminmgr-coverage': {
    title: 'Coverage Profiles And Monitoring',
    summary: 'Coverage profiles power Admin Manager diagnostics for content count and freshness.',
    bullets: [
      'Each profile describes one content area to monitor.',
      'Source URLs should point to SharePoint sites, lists, or libraries that matter to the search experience.',
      'Query templates let a profile measure only relevant content, such as documents or a content type family.',
      'Exclude paths remove archive or irrelevant locations from coverage calculations.',
      'Result source and refinement filters can mirror production search constraints.'
    ],
    examples: [
      'Policies profile: source URL /sites/hr/Policies, query template IsDocument:1.',
      'Account documents profile: source URLs for account libraries, exclude archive folders.'
    ]
  }
};

const FALLBACK_TOPIC: IPropertyPaneHelpTopic = {
  title: 'SP Search Help',
  summary: 'This section configures part of the SP Search experience.',
  bullets: [
    'Use the field descriptions in this property pane as the source of truth.',
    'Keep Search context ID values aligned across web parts that should share state.',
    'Validate changes in edit mode before publishing the page.'
  ]
};

export function getPropertyPaneHelpTopic(anchorId: string): IPropertyPaneHelpTopic {
  return HELP_TOPICS[anchorId] || FALLBACK_TOPIC;
}

function renderList(items: string[] | undefined, ordered: boolean): React.ReactElement | null {
  if (!items || items.length === 0) {
    return null;
  }
  const children = items.map((item: string, index: number): React.ReactElement => {
    return React.createElement('li', { key: String(index), style: { marginBottom: 6 } }, item);
  });
  return React.createElement(
    ordered ? 'ol' : 'ul',
    { style: { margin: '8px 0 0 20px', padding: 0 } },
    children
  );
}

const PropertyPaneHelpControl: React.FC<{
  anchorId: string;
  buttonText: string;
}> = (props) => {
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const topic = getPropertyPaneHelpTopic(props.anchorId);

  return React.createElement(
    'div',
    { style: { margin: '0 0 8px 0' } },
    React.createElement(
      Link,
      {
        href: '#',
        onClick: (event: React.MouseEvent<HTMLElement | HTMLAnchorElement | HTMLButtonElement>): void => {
          event.preventDefault();
          setIsOpen(true);
        },
        styles: {
          root: {
            display: 'inline-flex',
            alignItems: 'center',
            gap: 6,
            fontSize: 12,
            lineHeight: 18,
            fontWeight: 400
          }
        }
      },
      React.createElement(Icon, {
        iconName: 'Info',
        styles: { root: { fontSize: 13, lineHeight: 18 } }
      }),
      React.createElement('span', null, props.buttonText)
    ),
    React.createElement(
      Modal,
      {
        isOpen,
        onDismiss: (): void => setIsOpen(false),
        isBlocking: false,
        containerClassName: 'sp-search-property-pane-help-modal'
      },
      React.createElement(
        'div',
        {
          style: {
            width: 560,
            maxWidth: 'calc(100vw - 48px)',
            maxHeight: 'calc(100vh - 96px)',
            display: 'flex',
            flexDirection: 'column',
            background: '#fff'
          }
        },
        React.createElement(
          'div',
          {
            style: {
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              padding: '16px 20px',
              borderBottom: '1px solid #edebe9'
            }
          },
          React.createElement('h2', { style: { margin: 0, fontSize: 20, fontWeight: 600 } }, topic.title),
          React.createElement(IconButton, {
            iconProps: { iconName: 'Cancel' },
            ariaLabel: 'Close help',
            onClick: (): void => setIsOpen(false)
          })
        ),
        React.createElement(
          'div',
          { style: { padding: '16px 20px', overflowY: 'auto', lineHeight: 1.45 } },
          React.createElement('p', { style: { marginTop: 0 } }, topic.summary),
          React.createElement('h3', { style: { fontSize: 15, margin: '16px 0 0' } }, 'What to know'),
          renderList(topic.bullets, false),
          topic.examples && topic.examples.length > 0
            ? React.createElement(
              React.Fragment,
              null,
              React.createElement('h3', { style: { fontSize: 15, margin: '18px 0 0' } }, 'Examples'),
              renderList(topic.examples, false)
            )
            : null
        ),
        React.createElement(
          'div',
          {
            style: {
              display: 'flex',
              justifyContent: 'flex-end',
              gap: 8,
              padding: '12px 20px 16px',
              borderTop: '1px solid #edebe9'
            }
          },
          React.createElement(PrimaryButton, {
            text: 'Close',
            onClick: (): void => setIsOpen(false)
          })
        )
      )
    )
  );
};

export function propertyPaneGroupHelp(
  anchorId: string,
  linkText: string
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: 'help-' + anchorId,
    properties: {
      key: 'help-' + anchorId,
      onRender: function (domElement: HTMLElement): void {
        ReactDom.render(
          React.createElement(PropertyPaneHelpControl, {
            anchorId,
            buttonText: linkText
          }),
          domElement
        );
      },
      onDispose: function (domElement: HTMLElement): void {
        ReactDom.unmountComponentAtNode(domElement);
      }
    }
  };
}
