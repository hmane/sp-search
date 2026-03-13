import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import styles from './SpSearchBox.module.scss';
import type { ISuggestion } from '@interfaces/index';

export interface ISuggestionDropdownProps {
  suggestions: ISuggestion[];
  queryText: string;
  onSelect: (suggestion: ISuggestion) => void;
  onRemove: (suggestion: ISuggestion) => void;
  onDismiss: () => void;
}

function getGroupDisplay(groupName: string): { title: string; description: string } {
  switch (groupName) {
    case 'Suggestions':
      return {
        title: 'SharePoint Suggestions',
        description: 'Matches from the SharePoint search index'
      };
    case 'People':
      return {
        title: 'People Suggestions',
        description: 'Suggested names from SharePoint search'
      };
    case 'Quick Results':
      return {
        title: 'Quick Results',
        description: 'Top matching items from SharePoint search'
      };
    case 'Recent':
      return {
        title: 'Recent Searches',
        description: 'Your recent search history'
      };
    case 'Trending':
      return {
        title: 'Popular Searches',
        description: 'Frequently used terms from your history'
      };
    case 'Properties':
      return {
        title: 'Search Properties',
        description: 'Managed property shortcuts for advanced search'
      };
    default:
      return {
        title: groupName,
        description: ''
      };
  }
}

/**
 * Group suggestions by groupName, preserving insertion order
 * (providers already ordered by priority in the parent).
 */
function groupSuggestions(items: ISuggestion[]): Array<{ groupName: string; items: ISuggestion[] }> {
  const map: Map<string, ISuggestion[]> = new Map();
  for (let i = 0; i < items.length; i++) {
    const s = items[i];
    const existing = map.get(s.groupName);
    if (existing) {
      existing.push(s);
    } else {
      map.set(s.groupName, [s]);
    }
  }

  const groups: Array<{ groupName: string; items: ISuggestion[] }> = [];
  map.forEach(function (groupItems, groupName): void {
    groups.push({ groupName, items: groupItems });
  });
  return groups;
}

/**
 * SuggestionDropdown — displays grouped suggestions from all providers
 * below the search box. Supports keyboard navigation (arrow keys, Enter, Escape).
 */
const SuggestionDropdown: React.FC<ISuggestionDropdownProps> = (props: ISuggestionDropdownProps): React.ReactElement => {
  const { suggestions, queryText, onSelect, onRemove, onDismiss } = props;
  const [activeIndex, setActiveIndex] = React.useState<number>(-1);
  const listRef = React.useRef<HTMLDivElement>(undefined as unknown as HTMLDivElement);

  const normalizedQueryText = queryText.trim().toLowerCase();

  // Build a flat list of all items for keyboard navigation indexing
  const flatItems = React.useMemo(function (): ISuggestion[] {
    return suggestions;
  }, [suggestions]);

  // Reset active index when suggestions change
  React.useEffect(function (): void {
    setActiveIndex(-1);
  }, [suggestions]);

  // Keyboard handler
  React.useEffect(function (): (() => void) {
    function handleKeyDown(e: KeyboardEvent): void {
      if (flatItems.length === 0) {
        return;
      }

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setActiveIndex(function (prev): number {
          return prev < flatItems.length - 1 ? prev + 1 : 0;
        });
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setActiveIndex(function (prev): number {
          return prev > 0 ? prev - 1 : flatItems.length - 1;
        });
      } else if (e.key === 'Enter' && activeIndex >= 0 && activeIndex < flatItems.length) {
        e.preventDefault();
        onSelect(flatItems[activeIndex]);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        onDismiss();
      }
    }

    document.addEventListener('keydown', handleKeyDown);
    return function cleanup(): void {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [flatItems, activeIndex, onSelect, onDismiss]);

  // Scroll active item into view
  React.useEffect(function (): void {
    if (activeIndex >= 0 && listRef.current) {
      const activeEl = listRef.current.querySelector('[data-active="true"]');
      if (activeEl) {
        activeEl.scrollIntoView({ block: 'nearest' });
      }
    }
  }, [activeIndex]);

  if (suggestions.length === 0) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null as unknown as React.ReactElement;
  }

  const groups = groupSuggestions(suggestions);

  function renderSuggestionText(text: string): React.ReactNode {
    if (!normalizedQueryText || normalizedQueryText.length < 2) {
      return text;
    }

    const lowerText = text.toLowerCase();
    const matchIndex = lowerText.indexOf(normalizedQueryText);

    if (matchIndex < 0) {
      return text;
    }

    const before = text.substring(0, matchIndex);
    const match = text.substring(matchIndex, matchIndex + normalizedQueryText.length);
    const after = text.substring(matchIndex + normalizedQueryText.length);

    return (
      <>
        {before}
        <mark className={styles.suggestionMatch}>{match}</mark>
        {after}
      </>
    );
  }

  // Track flat index for keyboard navigation
  let flatIndex = 0;

  return (
    <div
      className={styles.suggestionsContainer}
      ref={listRef}
      role="listbox"
      aria-label="Search suggestions"
    >
      {groups.map(function (group): React.ReactElement {
        const groupDisplay = getGroupDisplay(group.groupName);

        return (
          <div key={group.groupName} className={styles.suggestionGroup}>
            <div className={styles.suggestionGroupHeader} role="presentation">
              <div className={styles.suggestionGroupMeta}>
                <span>{groupDisplay.title}</span>
                {groupDisplay.description && (
                  <span className={styles.suggestionGroupDescription}>{groupDisplay.description}</span>
                )}
              </div>
              <span className={styles.suggestionGroupCount}>{group.items.length}</span>
            </div>
            {group.items.map(function (suggestion): React.ReactElement {
              const currentIndex = flatIndex;
              flatIndex++;
              const isActive = currentIndex === activeIndex;

              return (
                <div
                  key={group.groupName + '-' + String(currentIndex)}
                  className={
                    isActive
                      ? styles.suggestionItem + ' ' + styles.suggestionItemActive
                      : styles.suggestionItem
                  }
                  onClick={function (): void { onSelect(suggestion); }}
                  role="option"
                  aria-selected={isActive}
                  data-active={isActive ? 'true' : undefined}
                  onMouseEnter={function (): void { setActiveIndex(currentIndex); }}
                >
                  {suggestion.filePath ? (
                    <span className={styles.suggestionFileIcon}>
                      <FileTypeIcon type={IconType.image} path={suggestion.filePath} size={ImageSize.small} />
                    </span>
                  ) : suggestion.iconName ? (
                    <Icon
                      iconName={suggestion.iconName}
                      className={styles.suggestionIcon}
                    />
                  ) : null}
                  <div className={styles.suggestionTextBlock}>
                    <span className={styles.suggestionText}>
                      {renderSuggestionText(suggestion.displayText)}
                    </span>
                    {suggestion.secondaryText && (
                      <span className={styles.suggestionSecondaryText} title={suggestion.secondaryText}>
                        {suggestion.secondaryText}
                      </span>
                    )}
                  </div>
                  {suggestion.removeAction && (
                    <IconButton
                      className={styles.suggestionRemoveButton}
                      iconProps={{ iconName: 'Cancel' }}
                      title={suggestion.removeLabel || 'Remove'}
                      ariaLabel={suggestion.removeLabel || 'Remove'}
                      onClick={function (event): void {
                        event.stopPropagation();
                        onRemove(suggestion);
                      }}
                    />
                  )}
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
};

export default SuggestionDropdown;
