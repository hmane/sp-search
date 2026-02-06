import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SpSearchManager.module.scss';

/** Pre-defined tag suggestions */
const PRESET_TAGS: string[] = [
  'Reviewed',
  'Important',
  'Follow-up',
  'Archived',
  'Action Required',
];

export interface IResultAnnotationsProps {
  itemId: number;
  tags: string[];
  onTagsChanged: (itemId: number, tags: string[]) => void;
}

/**
 * ResultAnnotations — inline tag editor for a single collection item.
 * Renders existing tags as removable pill badges and provides an input
 * field to add new tags with preset suggestions.
 */
const ResultAnnotations: React.FC<IResultAnnotationsProps> = (props) => {
  const { itemId, tags, onTagsChanged } = props;

  const [isEditing, setIsEditing] = React.useState<boolean>(false);
  const [inputValue, setInputValue] = React.useState<string>('');
  const [showSuggestions, setShowSuggestions] = React.useState<boolean>(false);
  const calloutTargetRef = React.useRef<HTMLDivElement | null>(null);

  // Compute available suggestions (exclude already-applied tags)
  const availableSuggestions: string[] = React.useMemo(function (): string[] {
    const result: string[] = [];
    for (let i = 0; i < PRESET_TAGS.length; i++) {
      if (tags.indexOf(PRESET_TAGS[i]) < 0) {
        result.push(PRESET_TAGS[i]);
      }
    }
    return result;
  }, [tags]);

  // Filter suggestions by input
  const filteredSuggestions: string[] = React.useMemo(function (): string[] {
    if (!inputValue.trim()) {
      return availableSuggestions;
    }
    const query = inputValue.toLowerCase().trim();
    const result: string[] = [];
    for (let i = 0; i < availableSuggestions.length; i++) {
      if (availableSuggestions[i].toLowerCase().indexOf(query) >= 0) {
        result.push(availableSuggestions[i]);
      }
    }
    return result;
  }, [availableSuggestions, inputValue]);

  function handleAddTag(tagText: string): void {
    const trimmed = tagText.trim();
    if (trimmed.length === 0 || trimmed.length > 50) {
      return;
    }
    // Check duplicate (case-insensitive)
    for (let i = 0; i < tags.length; i++) {
      if (tags[i].toLowerCase() === trimmed.toLowerCase()) {
        setInputValue('');
        return;
      }
    }
    const updated = tags.slice();
    updated.push(trimmed);
    onTagsChanged(itemId, updated);
    setInputValue('');
  }

  function handleRemoveTag(tagText: string): void {
    const updated: string[] = [];
    for (let i = 0; i < tags.length; i++) {
      if (tags[i] !== tagText) {
        updated.push(tags[i]);
      }
    }
    onTagsChanged(itemId, updated);
  }

  function handleInputChange(
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void {
    const val = newValue !== undefined ? newValue : '';
    // Comma triggers tag addition
    if (val.indexOf(',') >= 0) {
      const parts = val.split(',');
      for (let i = 0; i < parts.length - 1; i++) {
        if (parts[i].trim()) {
          handleAddTag(parts[i]);
        }
      }
      setInputValue(parts[parts.length - 1]);
    } else {
      setInputValue(val);
    }
  }

  function handleInputKeyDown(event: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>): void {
    if (event.key === 'Enter') {
      event.preventDefault();
      if (inputValue.trim()) {
        handleAddTag(inputValue);
      }
      setShowSuggestions(false);
    } else if (event.key === 'Escape') {
      setIsEditing(false);
      setInputValue('');
      setShowSuggestions(false);
    }
  }

  function handleInputFocus(): void {
    setShowSuggestions(true);
  }

  function handleInputBlur(): void {
    // Delay hiding to allow click on suggestion
    setTimeout(function (): void {
      setShowSuggestions(false);
    }, 200);
  }

  function handleSuggestionClick(suggestion: string): void {
    handleAddTag(suggestion);
    setShowSuggestions(false);
  }

  function handleToggleEdit(): void {
    setIsEditing(function (prev): boolean { return !prev; });
    setInputValue('');
  }

  return (
    <div className={styles.tagAnnotationRow}>
      {/* Existing tag badges (removable when editing) */}
      {tags.map(function (tag): React.ReactElement {
        return (
          <span key={tag} className={styles.tagBadge} title={tag}>
            {tag}
            {isEditing && (
              <button
                className={styles.tagBadgeRemove}
                onClick={function (): void { handleRemoveTag(tag); }}
                title={'Remove tag ' + tag}
                aria-label={'Remove tag ' + tag}
                type="button"
              >
                <Icon iconName="Cancel" style={{ fontSize: 8 }} />
              </button>
            )}
          </span>
        );
      })}

      {/* Add tag button / input */}
      {isEditing ? (
        <div ref={calloutTargetRef} className={styles.tagInput}>
          <TextField
            value={inputValue}
            onChange={handleInputChange}
            onKeyDown={handleInputKeyDown}
            onFocus={handleInputFocus}
            onBlur={handleInputBlur}
            placeholder="Add tag..."
            autoFocus={true}
            borderless={false}
          />
          {showSuggestions && filteredSuggestions.length > 0 && calloutTargetRef.current && (
            <Callout
              target={calloutTargetRef.current}
              directionalHint={DirectionalHint.bottomLeftEdge}
              isBeakVisible={false}
              gapSpace={2}
              onDismiss={function (): void { setShowSuggestions(false); }}
            >
              <ul className={styles.tagSuggestions}>
                {filteredSuggestions.map(function (suggestion): React.ReactElement {
                  return (
                    <li
                      key={suggestion}
                      className={styles.tagSuggestionItem}
                      onMouseDown={function (e: React.MouseEvent): void {
                        e.preventDefault();
                        handleSuggestionClick(suggestion);
                      }}
                    >
                      {suggestion}
                    </li>
                  );
                })}
              </ul>
            </Callout>
          )}
        </div>
      ) : (
        <IconButton
          iconProps={{ iconName: 'Tag', style: { fontSize: 12 } }}
          title="Add tags"
          ariaLabel="Add tags"
          onClick={handleToggleEdit}
          styles={{ root: { width: 24, height: 24 } }}
        />
      )}
    </div>
  );
};

export default ResultAnnotations;
