import { TelemetryTransport } from '../../src/libraries/spSearchStore/telemetry/TelemetryTransport';
import { ITelemetryConfig } from '../../src/libraries/spSearchStore/telemetry/ITelemetryConfig';

describe('TelemetryTransport (Found.D9)', () => {
  const baseSignal = { kind: 'queryTiming', timestamp: '2026-05-02T00:00:00Z' };

  afterEach(() => {
    delete (globalThis as unknown as { fetch?: unknown }).fetch;
  });

  it('flush is a no-op when isEnabled=false', async () => {
    const fetchMock = jest.fn();
    (globalThis as unknown as { fetch: jest.Mock }).fetch = fetchMock;
    const config: ITelemetryConfig = {
      isEnabled: false,
      destinationEndpoint: 'https://example.com/telemetry',
      batchIntervalSeconds: 300,
      batchSizeMax: 50,
    };
    const transport = new TelemetryTransport(() => Promise.resolve(config));
    await transport.flush([baseSignal]);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('flush POSTs to destinationEndpoint when isEnabled=true', async () => {
    const fetchMock = jest.fn().mockResolvedValue({ ok: true, status: 200 });
    (globalThis as unknown as { fetch: jest.Mock }).fetch = fetchMock;
    const config: ITelemetryConfig = {
      isEnabled: true,
      destinationEndpoint: 'https://example.com/telemetry',
      batchIntervalSeconds: 300,
      batchSizeMax: 50,
    };
    const transport = new TelemetryTransport(() => Promise.resolve(config));
    await transport.flush([baseSignal]);
    expect(fetchMock).toHaveBeenCalledWith(
      'https://example.com/telemetry',
      expect.objectContaining({
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
      }),
    );
  });

  it('flush retries with exponential backoff on 5xx, max 3 attempts', async () => {
    let attempt = 0;
    const fetchMock = jest.fn().mockImplementation(() => {
      attempt++;
      if (attempt < 3) return Promise.resolve({ ok: false, status: 503 });
      return Promise.resolve({ ok: true, status: 200 });
    });
    (globalThis as unknown as { fetch: jest.Mock }).fetch = fetchMock;
    const config: ITelemetryConfig = {
      isEnabled: true,
      destinationEndpoint: 'https://example.com/telemetry',
      batchIntervalSeconds: 300,
      batchSizeMax: 50,
    };
    const transport = new TelemetryTransport(() => Promise.resolve(config), { backoffMs: 0 });
    await transport.flush([baseSignal]);
    expect(fetchMock).toHaveBeenCalledTimes(3);
  });

  it('flush does not throw when fetch rejects', async () => {
    (globalThis as unknown as { fetch: jest.Mock }).fetch = jest.fn().mockRejectedValue(new Error('network'));
    const config: ITelemetryConfig = {
      isEnabled: true,
      destinationEndpoint: 'https://example.com/telemetry',
      batchIntervalSeconds: 300,
      batchSizeMax: 50,
    };
    const transport = new TelemetryTransport(() => Promise.resolve(config), { backoffMs: 0, maxAttempts: 1 });
    await expect(transport.flush([baseSignal])).resolves.toBeUndefined();
  });
});
