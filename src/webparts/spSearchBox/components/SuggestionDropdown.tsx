import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SpSearchBox.module.scss';
import type { ISuggestion } from '@interfaces/index';

export interface ISuggestionDropdownProps {
  suggestions: ISuggestion[];
  onSelect: (suggestion: ISuggestion) => void;
  onDismiss: () => void;
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
  const { suggestions, onSelect, onDismiss } = props;
  const [activeIndex, setActiveIndex] = React.useState<number>(-1);
  const listRef = React.useRef<HTMLDivElement>(undefined as unknown as HTMLDivElement);

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
        return (
          <div key={group.groupName} className={styles.suggestionGroup}>
            <div className={styles.suggestionGroupHeader} role="presentation">
              {group.groupName}
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
                >
                  {suggestion.iconName && (
                    <Icon
                      iconName={suggestion.iconName}
                      className={styles.suggestionIcon}
                    />
                  )}
                  <span className={styles.suggestionText}>
                    {suggestion.displayText}
                  </span>
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
