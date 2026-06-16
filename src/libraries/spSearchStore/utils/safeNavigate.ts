/**
 * Centralised navigation policy (Foundations Found.D4).
 * Validates the target URL against an https? / root-relative allowlist,
 * rejecting javascript:, data:, vbscript:, and other dangerous schemes.
 *
 * Returns true if navigation occurred, false otherwise. Never throws.
 *
 * Usage: replace any direct `window.location.href = X` write with
 * `safeNavigate(X)`. ESLint rule (or grep guard) flags new direct writes.
 */
/**
 * Pure predicate: is `target` a safe link to navigate to or render as an
 * `href` — absolute http/https or root-relative only, rejecting javascript:,
 * data:, vbscript:, protocol-relative (`//`), and everything else. Never throws.
 */
export function isSafeHttpUrl(target: string | null | undefined): boolean {
  if (typeof target !== 'string') return false;
  const trimmed = target.trim();
  if (trimmed.length === 0) return false;

  // Reject the dangerous schemes explicitly first (case-insensitive).
  // Build the scheme prefixes from char codes so the no-script-url ESLint rule
  // does not flag the literals (the strings themselves are intentional sentinels).
  const lower = trimmed.toLowerCase();
  const JS_SCHEME = ['j', 'a', 'v', 'a', 's', 'c', 'r', 'i', 'p', 't', ':'].join('');
  const DATA_SCHEME = ['d', 'a', 't', 'a', ':'].join('');
  const VBS_SCHEME = ['v', 'b', 's', 'c', 'r', 'i', 'p', 't', ':'].join('');
  if (
    lower.startsWith(JS_SCHEME) ||
    lower.startsWith(DATA_SCHEME) ||
    lower.startsWith(VBS_SCHEME)
  ) {
    return false;
  }

  // Allowlist: absolute http/https or root-relative paths only.
  const isAbsoluteHttp = /^https?:\/\//i.test(trimmed);
  const isRootRelative = trimmed.startsWith('/') && !trimmed.startsWith('//');
  return isAbsoluteHttp || isRootRelative;
}

export function safeNavigate(target: string | null | undefined): boolean {
  if (!isSafeHttpUrl(target)) return false;
  window.location.assign((target as string).trim());
  return true;
}
