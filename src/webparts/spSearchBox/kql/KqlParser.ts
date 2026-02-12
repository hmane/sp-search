import type { IKqlCompletionContext, IKqlToken } from './KqlTypes';

/** KQL boolean connectives (case-insensitive match). */
const CONNECTIVES: Set<string> = new Set(['AND', 'OR', 'NOT']);

/**
 * Tokenizes a KQL input string into tokens delimited by whitespace and parentheses.
 * Respects quoted strings as single tokens.
 */
export function tokenize(input: string): IKqlToken[] {
  const tokens: IKqlToken[] = [];
  let i: number = 0;
  const len: number = input.length;

  while (i < len) {
    // Skip whitespace
    if (input[i] === ' ' || input[i] === '\t') {
      i++;
      continue;
    }

    // Parentheses are standalone tokens
    if (input[i] === '(' || input[i] === ')') {
      tokens.push({ text: input[i], start: i, end: i + 1 });
      i++;
      continue;
    }

    // Quoted string
    if (input[i] === '"') {
      const start: number = i;
      i++; // skip opening quote
      while (i < len && input[i] !== '"') {
        i++;
      }
      if (i < len) {
        i++; // skip closing quote
      }
      tokens.push({ text: input.substring(start, i), start, end: i });
      continue;
    }

    // Regular token — read until whitespace, paren, or end
    const start: number = i;
    while (i < len && input[i] !== ' ' && input[i] !== '\t' && input[i] !== '(' && input[i] !== ')') {
      // Handle quoted segments within tokens (e.g., Property:"value with spaces")
      if (input[i] === '"') {
        i++;
        while (i < len && input[i] !== '"') {
          i++;
        }
        if (i < len) {
          i++; // skip closing quote
        }
      } else {
        i++;
      }
    }
    tokens.push({ text: input.substring(start, i), start, end: i });
  }

  return tokens;
}

/**
 * Finds the token at or immediately before the cursor position.
 * Returns the token index, or -1 if cursor is in whitespace at start.
 */
function findTokenAtCursor(tokens: IKqlToken[], cursorPosition: number): number {
  // If cursor is inside or at the end of a token, return that token
  for (let i: number = 0; i < tokens.length; i++) {
    if (cursorPosition >= tokens[i].start && cursorPosition <= tokens[i].end) {
      return i;
    }
  }
  // Cursor is in whitespace — return -1 to indicate "after last token"
  return -1;
}

/**
 * Checks whether a token text contains a property:value delimiter.
 * Returns the index of the delimiter character within the token, or -1.
 */
function findPropertyDelimiter(text: string): number {
  // Check for multi-char operators first: <>, >=, <=
  const twoCharOps: string[] = ['<>', '>=', '<='];
  for (let j: number = 0; j < twoCharOps.length; j++) {
    const idx: number = text.indexOf(twoCharOps[j]);
    if (idx > 0) {
      return idx;
    }
  }
  // Single-char operators: :, =, >, <
  for (let i: number = 0; i < text.length; i++) {
    const ch: string = text[i];
    if (ch === ':' || ch === '=' || ch === '>' || ch === '<') {
      // Make sure it's not at position 0 (would be an operator token, not property:value)
      if (i > 0) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * Determines the operator length at a given position in a token.
 */
function getOperatorLength(text: string, pos: number): number {
  const twoChar: string = text.substring(pos, pos + 2);
  if (twoChar === '<>' || twoChar === '>=' || twoChar === '<=') {
    return 2;
  }
  return 1;
}

/**
 * Analyzes the KQL input at the given cursor position and returns the
 * completion context — what kind of completions should be shown.
 *
 * This is a rule-based parser, not a full AST parser. It examines the
 * text left of the cursor to determine what the user is typing.
 */
export function getCompletionContext(input: string, cursorPosition: number): IKqlCompletionContext {
  if (!input || cursorPosition === 0) {
    return {
      type: 'PropertyName',
      partialText: '',
      tokenStart: 0,
      tokenEnd: 0,
    };
  }

  const tokens: IKqlToken[] = tokenize(input);

  if (tokens.length === 0) {
    return {
      type: 'PropertyName',
      partialText: '',
      tokenStart: cursorPosition,
      tokenEnd: cursorPosition,
    };
  }

  const tokenIdx: number = findTokenAtCursor(tokens, cursorPosition);

  // Case 1: Cursor is in whitespace after all tokens
  if (tokenIdx === -1) {
    const lastToken: IKqlToken = tokens[tokens.length - 1];
    const lastText: string = lastToken.text.toUpperCase();

    // After a connective → suggest property names
    if (CONNECTIVES.has(lastText)) {
      return {
        type: 'PropertyName',
        partialText: '',
        tokenStart: cursorPosition,
        tokenEnd: cursorPosition,
      };
    }

    // After an opening paren → suggest property names
    if (lastText === '(') {
      return {
        type: 'PropertyName',
        partialText: '',
        tokenStart: cursorPosition,
        tokenEnd: cursorPosition,
      };
    }

    // After a complete clause → suggest boolean connectives
    return {
      type: 'BooleanConnective',
      partialText: '',
      tokenStart: cursorPosition,
      tokenEnd: cursorPosition,
    };
  }

  const currentToken: IKqlToken = tokens[tokenIdx];
  const currentText: string = currentToken.text;
  // The partial text is only the part up to the cursor, not the full token
  const partialFromToken: string = currentText.substring(0, cursorPosition - currentToken.start);

  // Case 2: Current token contains a property delimiter (e.g., "Author:John" or "FileType:do")
  const delimIdx: number = findPropertyDelimiter(partialFromToken);
  if (delimIdx > 0) {
    const propertyName: string = partialFromToken.substring(0, delimIdx);
    const opLen: number = getOperatorLength(partialFromToken, delimIdx);
    const partialValue: string = partialFromToken.substring(delimIdx + opLen);

    return {
      type: 'PropertyValue',
      partialText: partialValue,
      propertyName: propertyName,
      tokenStart: currentToken.start + delimIdx + opLen,
      tokenEnd: currentToken.end,
    };
  }

  // Case 3: Current token is a standalone paren
  if (currentText === '(') {
    return {
      type: 'PropertyName',
      partialText: '',
      tokenStart: cursorPosition,
      tokenEnd: cursorPosition,
    };
  }

  // Case 4: Look at previous token to determine context
  const prevToken: IKqlToken | undefined = tokenIdx > 0 ? tokens[tokenIdx - 1] : undefined;

  if (prevToken) {
    const prevUpper: string = prevToken.text.toUpperCase();

    // After a connective → user is typing a property name
    if (CONNECTIVES.has(prevUpper)) {
      return {
        type: 'PropertyName',
        partialText: partialFromToken,
        tokenStart: currentToken.start,
        tokenEnd: currentToken.end,
      };
    }

    // After an opening paren → user is typing a property name
    if (prevUpper === '(') {
      return {
        type: 'PropertyName',
        partialText: partialFromToken,
        tokenStart: currentToken.start,
        tokenEnd: currentToken.end,
      };
    }
  }

  // Case 5: Check if user might be typing a connective (e.g., "AN" → could be AND)
  const partialUpper: string = partialFromToken.toUpperCase();
  if (partialFromToken.length >= 1) {
    let couldBeConnective: boolean = false;
    const connectiveList: string[] = ['AND', 'OR', 'NOT'];
    for (let c: number = 0; c < connectiveList.length; c++) {
      if (connectiveList[c].indexOf(partialUpper) === 0) {
        couldBeConnective = true;
        break;
      }
    }

    // Only suggest connective if there's a preceding clause (not at start)
    if (couldBeConnective && tokenIdx > 0) {
      // Check if previous token looks like a completed clause
      const prev: IKqlToken = tokens[tokenIdx - 1];
      if (prev.text !== '(' && !CONNECTIVES.has(prev.text.toUpperCase())) {
        return {
          type: 'BooleanConnective',
          partialText: partialFromToken,
          tokenStart: currentToken.start,
          tokenEnd: currentToken.end,
        };
      }
    }
  }

  // Default: user is typing a property name or free text
  return {
    type: 'PropertyName',
    partialText: partialFromToken,
    tokenStart: currentToken.start,
    tokenEnd: currentToken.end,
  };
}
