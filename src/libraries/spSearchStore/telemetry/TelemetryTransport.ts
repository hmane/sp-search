import { ITelemetryConfig } from './ITelemetryConfig';
import { ITelemetrySignal } from './ITelemetrySignal';

export interface TelemetryTransportOptions {
  /** Base backoff in ms; doubles per attempt. Default 2000. */
  backoffMs?: number;
  /** Maximum number of attempts (initial + retries). Default 3. */
  maxAttempts?: number;
  /** How long the resolved config is cached before re-loading. Default 60s. */
  configRefreshSeconds?: number;
}

/**
 * Foundations Found.D9 — HTTPS POST transport for opt-in telemetry.
 * Never inspects payload contents. T5.D8's ITelemetrySignal discriminated
 * union enforces the never-captured field list at compile time.
 *
 * Usage: instantiated by sp-search-store; consumes the SearchTelemetryConfig
 * SP list via the configLoader callback. Returns immediately when
 * config.isEnabled === false (the no-op default).
 */
export class TelemetryTransport {
  private cachedConfig: ITelemetryConfig | null = null;
  private lastConfigLoad = 0;
  private readonly backoffMs: number;
  private readonly maxAttempts: number;
  private readonly configRefreshMs: number;

  public constructor(
    private readonly configLoader: () => Promise<ITelemetryConfig>,
    options: TelemetryTransportOptions = {},
  ) {
    this.backoffMs = options.backoffMs !== undefined ? options.backoffMs : 2000;
    this.maxAttempts = options.maxAttempts !== undefined ? options.maxAttempts : 3;
    this.configRefreshMs = (options.configRefreshSeconds !== undefined ? options.configRefreshSeconds : 60) * 1000;
  }

  public async flush(batch: ITelemetrySignal[]): Promise<void> {
    if (batch.length === 0) return;
    let config: ITelemetryConfig;
    try {
      config = await this.loadConfig();
    } catch {
      // Config load failure - swallow silently and skip this flush cycle.
      return;
    }
    if (!config.isEnabled || !config.destinationEndpoint) return;

    const body = JSON.stringify({ signals: batch });
    let attempt = 0;
    while (attempt < this.maxAttempts) {
      attempt++;
      try {
        const res = await fetch(config.destinationEndpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body,
        });
        if (res.ok) return;
        if (res.status >= 400 && res.status < 500 && res.status !== 408 && res.status !== 429) {
          // Permanent client error — drop and stop retrying.
          return;
        }
      } catch {
        // Network error — retry with backoff.
      }
      if (attempt < this.maxAttempts) {
        const delay = this.backoffMs * Math.pow(2, attempt - 1);
        if (delay > 0) {
          await new Promise<void>((resolve) => setTimeout(resolve, delay));
        }
      }
    }
  }

  private async loadConfig(): Promise<ITelemetryConfig> {
    const now = Date.now();
    if (this.cachedConfig && now - this.lastConfigLoad < this.configRefreshMs) {
      return this.cachedConfig;
    }
    this.cachedConfig = await this.configLoader();
    this.lastConfigLoad = now;
    return this.cachedConfig;
  }
}
