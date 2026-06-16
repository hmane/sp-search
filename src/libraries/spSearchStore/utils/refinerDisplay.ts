/**
 * Format a raw refiner token value for human-readable display. Decodes the
 * SharePoint `ǂǂ`-prefixed hex tokens (e.g. "ǂǂ31393534" → "1954"), strips FQL
 * `string("…")` wrappers, `GP0|#GUID` taxonomy prefixes, and surrounding quotes.
 *
 * Shared by the active-filter pills and the Search Manager history so a stored
 * refinement token always renders as its label, not its internal id.
 */
// SharePoint FQL hex-encoded refinement marker — two U+01C2 (ǂ) chars. Built
// from char codes (not a string literal) so the match never depends on how a
// given toolchain decodes the source file's encoding.
const HEX_PREFIX = String.fromCharCode(0x01C2, 0x01C2);

export function formatRefinerValueForDisplay(rawValue: string): string {
  let value = String(rawValue || '').trim();
  if (!value) {
    return '';
  }

  // Strip a single layer of surrounding quotes.
  if (value.length >= 2 && value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
    value = value.substring(1, value.length - 1);
  }

  // ǂǂ hex token → decode the hex pairs back to text. Treat each hex pair as a
  // percent-encoded byte and let decodeURIComponent handle UTF-8 (works in the
  // browser AND jsdom, unlike TextDecoder which jsdom doesn't provide).
  if (value.indexOf(HEX_PREFIX) === 0) {
    const hex = value.substring(2);
    if (hex.length > 0 && hex.length % 2 === 0 && /^[0-9a-fA-F]+$/.test(hex)) {
      try {
        let encoded = '';
        for (let i = 0; i < hex.length; i += 2) {
          encoded += '%' + hex.substring(i, i + 2);
        }
        return decodeURIComponent(encoded);
      } catch {
        // fall through to the raw value
      }
    }
  }

  // FQL string("…") wrapper.
  if (value.indexOf('string("') === 0 && value.lastIndexOf('")') === value.length - 2) {
    return value.substring(8, value.length - 2);
  }

  // GP0|#GUID taxonomy token — the label is the last pipe segment.
  if (value.indexOf('GP0|#') >= 0) {
    const parts = value.split('|');
    return parts[parts.length - 1] || value;
  }

  return value;
}
