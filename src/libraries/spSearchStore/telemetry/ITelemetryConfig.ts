export interface ITelemetryConfig {
  isEnabled: boolean;
  destinationEndpoint: string;
  batchIntervalSeconds: number;
  batchSizeMax: number;
  privacyAcknowledgedBy?: string;
  privacyAcknowledgedAt?: string;
}

export const TELEMETRY_DEFAULTS: ITelemetryConfig = {
  isEnabled: false,
  destinationEndpoint: '',
  batchIntervalSeconds: 300,
  batchSizeMax: 50,
};
