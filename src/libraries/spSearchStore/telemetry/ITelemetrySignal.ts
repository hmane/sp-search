/**
 * Minimal signal interface so Foundations Found.D9 transport can compile
 * before T5.D8 lands the canonical schema. T5.D8 replaces this interface
 * with the full type discriminated union (kind: 'queryTiming' | 'errorRate' | ...).
 *
 * Foundations enforces: NEVER capture queryText, userId, resultTitle, urls,
 * tenantName, or list item content. T5.D8 enforces this at compile time
 * via the ITelemetrySignal discriminated union.
 */
export interface ITelemetrySignal {
  kind: string;
  timestamp: string;
  // Type-erased payload until T5.D8 lands the schema. Transport never
  // inspects payload contents — purely the wire.
  [key: string]: unknown;
}
