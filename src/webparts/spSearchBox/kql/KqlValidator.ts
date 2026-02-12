import type { IKqlValidation } from './KqlTypes';

/**
 * Lightweight KQL syntax validator.
 * Checks for common syntax issues — balanced parens/quotes, dangling operators.
 * Optionally warns about unknown property names.
 */
export function validate(input: string, knownProperties?: Set<string>): IKqlValidation {
  if (!input || input.trim().length === 0) {
    return { isValid: true, severity: 'valid', message: '' };
  }

  const trimmed: string = input.trim();

  // Check balanced parentheses
  let parenDepth: number = 0;
  let inQuote: boolean = false;
  for (let i: number = 0; i < trimmed.length; i++) {
    const ch: string = trimmed[i];
    if (ch === '"') {
      inQuote = !inQuote;
    } else if (!inQuote) {
      if (ch === '(') {
        parenDepth++;
      } else if (ch === ')') {
        parenDepth--;
        if (parenDepth < 0) {
          return {
            isValid: false,
            severity: 'error',
            message: 'Unexpected closing parenthesis at position ' + String(i + 1),
          };
        }
      }
    }
  }

  if (parenDepth > 0) {
    return {
      isValid: false,
      severity: 'error',
      message: 'Missing ' + String(parenDepth) + ' closing parenthes' + (parenDepth === 1 ? 'is' : 'es'),
    };
  }

  // Check balanced quotes
  if (inQuote) {
    return {
      isValid: false,
      severity: 'error',
      message: 'Unclosed quotation mark',
    };
  }

  // Check dangling operators at start or end
  const words: string[] = trimmed.split(/\s+/);
  const connectives: string[] = ['AND', 'OR'];
  const firstWord: string = words[0].toUpperCase();
  const lastWord: string = words[words.length - 1].toUpperCase();

  if (connectives.indexOf(firstWord) >= 0) {
    return {
      isValid: false,
      severity: 'error',
      message: firstWord + ' cannot be at the start of the query',
    };
  }

  if (words.length > 1 && connectives.indexOf(lastWord) >= 0) {
    return {
      isValid: false,
      severity: 'warning',
      message: lastWord + ' at end of query — expecting another term',
    };
  }

  // Check for unknown property names (warning, not error)
  if (knownProperties && knownProperties.size > 0) {
    for (let w: number = 0; w < words.length; w++) {
      const word: string = words[w];
      const colonIdx: number = word.indexOf(':');
      if (colonIdx > 0) {
        const propName: string = word.substring(0, colonIdx);
        // Skip if it starts with NOT (e.g., NOTFileType:)
        if (propName.toUpperCase() !== 'NOT' && !knownProperties.has(propName)) {
          return {
            isValid: true,
            severity: 'warning',
            message: 'Unknown property: ' + propName,
          };
        }
      }
    }
  }

  return { isValid: true, severity: 'valid', message: '' };
}
