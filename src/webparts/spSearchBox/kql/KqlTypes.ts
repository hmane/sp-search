/**
 * KQL auto-completion and validation type definitions.
 */

/** Completion context detected by KqlParser based on cursor position. */
export interface IKqlCompletionContext {
  /** What kind of completion is expected. */
  type: 'PropertyName' | 'PropertyValue' | 'BooleanConnective' | 'FreeText';
  /** The partial text being typed in the current token (for filtering). */
  partialText: string;
  /** If type is PropertyValue, the property name being queried. */
  propertyName?: string;
  /** Character offset where the current token starts (for replacement). */
  tokenStart: number;
  /** Character offset where the current token ends (for replacement). */
  tokenEnd: number;
}

/** A single auto-completion item shown in the dropdown. */
export interface IKqlCompletion {
  /** The text to insert at cursor position. */
  insertText: string;
  /** What the user sees in the dropdown. */
  displayText: string;
  /** Category badge type. */
  completionType: 'property' | 'value' | 'operator' | 'keyword';
  /** Property type indicator for property completions (Text, DateTime, Integer, etc.). */
  propertyType?: string;
  /** Refiner count for value completions. */
  count?: number;
  /** Secondary description text. */
  description?: string;
  /** Fluent UI icon name. */
  iconName?: string;
}

/** Validation result from KqlValidator. */
export interface IKqlValidation {
  isValid: boolean;
  severity: 'valid' | 'warning' | 'error';
  message: string;
}

/** Internal token produced by the parser tokenizer. */
export interface IKqlToken {
  text: string;
  start: number;
  end: number;
}
