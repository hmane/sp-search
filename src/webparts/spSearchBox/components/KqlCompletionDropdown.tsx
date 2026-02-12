import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IKqlCompletion, IKqlCompletionContext } from '../kql';
import styles from './SpSearchBox.module.scss';

export interface IKqlCompletionDropdownProps {
  completions: IKqlCompletion[];
  context: IKqlCompletionContext | undefined;
  onSelect: (completion: IKqlCompletion) => void;
  onDismiss: () => void;
  /** Called by parent when Enter/Tab should accept current selection. Returns true if handled. */
  onKeyAction?: (key: string) => boolean;
}

/** Static KQL syntax help items shown as a fallback group. */
const SYNTAX_HELP: Array<{ syntax: string; description: string }> = [
  { syntax: 'Property:value', description: 'Contains' },
  { syntax: 'Property="exact phrase"', description: 'Exact match' },
  { syntax: 'AND / OR / NOT', description: 'Boolean operators' },
  { syntax: '( ... )', description: 'Grouping' },
  { syntax: 'Title:annual*', description: 'Wildcard' },
  { syntax: '"annual report"', description: 'Exact phrase' },
];

/**
 * Badge class based on completion type.
 */
function getBadgeClass(type: string): string {
  switch (type) {
    case 'property': return styles.kqlBadgeProperty;
    case 'value': return styles.kqlBadgeValue;
    case 'operator': return styles.kqlBadgeOperator;
    case 'keyword': return styles.kqlBadgeKeyword;
    default: return styles.kqlBadgeValue;
  }
}

/**
 * Badge label based on completion type.
 */
function getBadgeLabel(type: string): string {
  switch (type) {
    case 'property': return 'Property';
    case 'value': return 'Value';
    case 'operator': return 'Operator';
    case 'keyword': return 'Keyword';
    default: return '';
  }
}

/**
 * Format a count number for display (e.g., 1245 → "1,245").
 */
function formatCount(count: number): string {
  return count.toLocaleString();
}

/**
 * KqlCompletionDropdown — context-aware completion dropdown for KQL mode.
 * Supports keyboard navigation (ArrowUp/Down, Enter/Tab to accept, Escape to dismiss).
 */
const KqlCompletionDropdown: React.FC<IKqlCompletionDropdownProps> = (props) => {
  const { completions, onSelect, onDismiss } = props;
  const [activeIndex, setActiveIndex] = React.useState<number>(0);
  const listRef = React.useRef<HTMLDivElement>(undefined as unknown as HTMLDivElement);

  // Reset active index when completions change
  React.useEffect(() => {
    setActiveIndex(0);
  }, [completions]);

  // Scroll active item into view
  React.useEffect(() => {
    if (activeIndex >= 0 && listRef.current) {
      const activeEl = listRef.current.querySelector('[data-active="true"]');
      if (activeEl) {
        activeEl.scrollIntoView({ block: 'nearest' });
      }
    }
  }, [activeIndex]);

  // Keyboard handler
  React.useEffect(() => {
    function handleKeyDown(e: KeyboardEvent): void {
      if (completions.length === 0) {
        if (e.key === 'Escape') {
          e.preventDefault();
          onDismiss();
        }
        return;
      }

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setActiveIndex((prev) => (prev < completions.length - 1 ? prev + 1 : 0));
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setActiveIndex((prev) => (prev > 0 ? prev - 1 : completions.length - 1));
      } else if ((e.key === 'Enter' || e.key === 'Tab') && activeIndex >= 0 && activeIndex < completions.length) {
        e.preventDefault();
        e.stopPropagation();
        onSelect(completions[activeIndex]);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        onDismiss();
      }
    }

    document.addEventListener('keydown', handleKeyDown, true);
    return (): void => {
      document.removeEventListener('keydown', handleKeyDown, true);
    };
  }, [completions, activeIndex, onSelect, onDismiss]);

  const showSyntaxHelp: boolean = completions.length === 0;

  return (
    <div
      className={styles.kqlCompletionContainer}
      ref={listRef}
      role="listbox"
      aria-label="KQL completions"
    >
      {/* Completion items */}
      {completions.map((completion: IKqlCompletion, index: number): React.ReactElement => {
        const isActive: boolean = index === activeIndex;
        const itemClass: string = isActive
          ? styles.kqlCompletionItem + ' ' + styles.kqlCompletionItemActive
          : styles.kqlCompletionItem;

        return (
          <div
            key={completion.completionType + '-' + completion.insertText + '-' + String(index)}
            className={itemClass}
            onClick={(): void => onSelect(completion)}
            role="option"
            aria-selected={isActive}
            data-active={isActive ? 'true' : undefined}
          >
            {completion.iconName && (
              <span className={styles.kqlCompletionIcon}>
                <Icon iconName={completion.iconName} />
              </span>
            )}
            <span className={styles.kqlCompletionText}>
              <span className={styles.kqlCompletionName}>{completion.displayText}</span>
              {completion.propertyType && (
                <span className={styles.kqlCompletionType}>{completion.propertyType}</span>
              )}
              {completion.description && !completion.propertyType && (
                <span className={styles.kqlCompletionType}>{completion.description}</span>
              )}
            </span>
            <span className={styles.kqlCompletionMeta}>
              {completion.count !== undefined && (
                <span className={styles.kqlCompletionCount}>{formatCount(completion.count)}</span>
              )}
              <span className={styles.kqlCompletionBadge + ' ' + getBadgeClass(completion.completionType)}>
                {getBadgeLabel(completion.completionType)}
              </span>
            </span>
          </div>
        );
      })}

      {/* Syntax help when no completions */}
      {showSyntaxHelp && (
        <div className={styles.kqlSyntaxGroup}>
          <div className={styles.kqlSyntaxHeader}>KQL Syntax</div>
          {SYNTAX_HELP.map((item): React.ReactElement => (
            <div key={item.syntax} className={styles.kqlSyntaxRow}>
              <code className={styles.kqlSyntaxCode}>{item.syntax}</code>
              <span className={styles.kqlSyntaxDesc}>{item.description}</span>
            </div>
          ))}
        </div>
      )}

      {/* Footer hint */}
      <div className={styles.kqlCompletionFooter}>
        {completions.length > 0
          ? 'Tab to accept \u00b7 Ctrl+Space for all'
          : 'Ctrl+Space for property suggestions'}
      </div>
    </div>
  );
};

export default KqlCompletionDropdown;
