import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { getCompletionContext, getCompletions, validate } from '../kql';
import type { IKqlCompletionContext, IKqlCompletion, IKqlValidation } from '../kql';
import type { IManagedProperty, IRefiner } from '@interfaces/index';
import styles from './SpSearchBox.module.scss';

export interface IKqlInputProps {
  value: string;
  onChange: (newValue: string, cursorPosition: number) => void;
  onSearch: (value: string) => void;
  onClear: () => void;
  onCompletionsChange: (completions: IKqlCompletion[], context: IKqlCompletionContext | undefined) => void;
  onValidationChange: (validation: IKqlValidation) => void;
  onFocus: () => void;
  onBlur: () => void;
  onForceOpenCompletions: () => void;
  placeholder?: string;
  schema: IManagedProperty[];
  refiners: IRefiner[];
  disabled?: boolean;
}

/**
 * KqlInput — custom input component for KQL mode.
 * Renders a monospace <input> with cursor tracking, completion triggering,
 * validation indicator, and clear button.
 */
const KqlInput: React.FC<IKqlInputProps> = (props) => {
  const {
    value,
    onChange,
    onSearch,
    onClear,
    onCompletionsChange,
    onValidationChange,
    onFocus,
    onBlur,
    schema,
    refiners,
    placeholder,
    disabled,
  } = props;

  const inputRef = React.useRef<HTMLInputElement>(undefined as unknown as HTMLInputElement);
  const completionTimerRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const validationTimerRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const [validation, setValidation] = React.useState<IKqlValidation>({ isValid: true, severity: 'valid', message: '' });

  // Cleanup timers
  React.useEffect(() => {
    return (): void => {
      if (completionTimerRef.current !== undefined) {
        clearTimeout(completionTimerRef.current);
      }
      if (validationTimerRef.current !== undefined) {
        clearTimeout(validationTimerRef.current);
      }
    };
  }, []);

  // Build known properties set for validation
  const knownProperties = React.useMemo((): Set<string> => {
    const set: Set<string> = new Set();
    for (let i: number = 0; i < schema.length; i++) {
      if (schema[i].queryable) {
        set.add(schema[i].name);
        if (schema[i].alias) {
          set.add(schema[i].alias as string);
        }
      }
    }
    return set;
  }, [schema]);

  /**
   * Triggers completion computation with debounce.
   */
  function triggerCompletions(text: string, cursor: number): void {
    if (completionTimerRef.current !== undefined) {
      clearTimeout(completionTimerRef.current);
    }

    completionTimerRef.current = setTimeout((): void => {
      completionTimerRef.current = undefined;

      if (text.length === 0) {
        onCompletionsChange([], undefined);
        return;
      }

      const context: IKqlCompletionContext = getCompletionContext(text, cursor);
      const completions: IKqlCompletion[] = getCompletions(context, schema, refiners);
      onCompletionsChange(completions, context);
    }, 150);
  }

  /**
   * Triggers validation with debounce.
   */
  function triggerValidation(text: string): void {
    if (validationTimerRef.current !== undefined) {
      clearTimeout(validationTimerRef.current);
    }

    validationTimerRef.current = setTimeout((): void => {
      validationTimerRef.current = undefined;
      const result: IKqlValidation = validate(text, knownProperties);
      setValidation(result);
      onValidationChange(result);
    }, 250);
  }

  /**
   * Handle input change.
   */
  function handleChange(e: React.ChangeEvent<HTMLInputElement>): void {
    const newValue: string = e.target.value;
    const cursor: number = e.target.selectionStart || newValue.length;

    onChange(newValue, cursor);
    triggerCompletions(newValue, cursor);
    triggerValidation(newValue);
  }

  /**
   * Handle cursor movement (click or arrow keys).
   */
  function handleSelect(): void {
    if (inputRef.current && value.length > 0) {
      const cursor: number = inputRef.current.selectionStart || 0;
      triggerCompletions(value, cursor);
    }
  }

  /**
   * Handle keydown — special keys.
   */
  function handleKeyDown(e: React.KeyboardEvent<HTMLInputElement>): void {
    // Ctrl+Space → force open completions
    if (e.ctrlKey && e.key === ' ') {
      e.preventDefault();
      const cursor: number = inputRef.current?.selectionStart || value.length;
      // Immediately compute completions (no debounce)
      if (completionTimerRef.current !== undefined) {
        clearTimeout(completionTimerRef.current);
        completionTimerRef.current = undefined;
      }
      const context: IKqlCompletionContext = getCompletionContext(value, cursor);
      const completions: IKqlCompletion[] = getCompletions(context, schema, refiners);
      onCompletionsChange(completions, context);
      props.onForceOpenCompletions();
      return;
    }

    // Enter → execute search directly
    if (e.key === 'Enter' && !e.ctrlKey && !e.shiftKey) {
      e.preventDefault();
      onSearch(value);
      return;
    }
  }

  /**
   * Handle clear button click.
   */
  function handleClear(e: React.MouseEvent<HTMLButtonElement>): void {
    e.preventDefault();
    e.stopPropagation();
    onClear();
    if (inputRef.current) {
      inputRef.current.focus();
    }
  }

  // Determine validation icon
  let validationIconName: string | undefined;
  let validationClass: string = '';
  if (value.length > 0) {
    if (validation.severity === 'error') {
      validationIconName = 'ErrorBadge';
      validationClass = styles.kqlValidationError;
    } else if (validation.severity === 'warning') {
      validationIconName = 'Warning';
      validationClass = styles.kqlValidationWarning;
    }
  }

  return (
    <div className={styles.kqlInputContainer}>
      <span className={styles.kqlInputIcon}>
        <Icon iconName="Code" />
      </span>
      <input
        ref={inputRef}
        type="text"
        className={styles.kqlInputField}
        value={value}
        onChange={handleChange}
        onSelect={handleSelect}
        onKeyDown={handleKeyDown}
        onFocus={onFocus}
        onBlur={onBlur}
        placeholder={placeholder || 'Enter KQL... (e.g., Author:John AND FileType:docx)'}
        disabled={disabled}
        autoComplete="off"
        spellCheck={false}
        role="combobox"
        aria-expanded={false}
        aria-autocomplete="list"
        aria-label="KQL query input"
      />
      {validationIconName && (
        <span className={styles.kqlValidationIcon + ' ' + validationClass} title={validation.message}>
          <Icon iconName={validationIconName} />
        </span>
      )}
      {value.length > 0 && (
        <button
          className={styles.kqlClearButton}
          onClick={handleClear}
          title="Clear"
          aria-label="Clear query"
          type="button"
        >
          <Icon iconName="Clear" />
        </button>
      )}
    </div>
  );
};

export default KqlInput;
