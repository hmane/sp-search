# SP Search — Privacy Notice (Telemetry)

> Owned by Foundations track (Found.D9). Current runtime does not instantiate telemetry; the lists and transport are provisioned for future opt-in use and remain disabled by default.

## What we collect (when telemetry is enabled)

If telemetry is wired in a future release, an admin would enable it via the `SearchTelemetryConfig` list and a user opt-in flow before any data is sent. The intended signal set is:

- Query timing — milliseconds end-to-end per search request
- Error rates — count of failed search requests, grouped by error class (no message bodies)
- Refiner usage — count of filter applications, grouped by filter type (Checkbox, DateRange, etc.)
- Layout switches — count of layout changes, grouped by layout key (DataGrid, Card, etc.)
- Vertical switches — count of vertical changes, grouped by vertical key
- Feature adoption — flags indicating whether a user opens the detail panel, opens the Search Manager, exports CSV/XLSX
- Anonymized session ID — SHA-256 hash of `tenantId + userPrincipalName + 'sp-search-telemetry-v1'`, truncated to first 8 hex chars

All counts are aggregated client-side per `BatchIntervalSeconds` (default 300s) before transmission.

## What we NEVER collect

- Query text (the literal string typed by users)
- User identity (email, login name, display name, UPN, or any reversible token)
- Result titles, URLs, or summaries
- Tenant name, site collection name, list name, or item content
- Page URL or referrer
- IP address (relies on transport infrastructure to redact at the destination)
- Browser fingerprint, geolocation, or device identifier

The transport (`TelemetryTransport.ts`) never inspects payload contents — it is purely the wire. The active runtime currently does not instantiate it.

## Where the data goes

The destination is admin-configured in the `SearchTelemetryConfig` list (`DestinationEndpoint` field). Same plumbing supports:
- Application Insights ingestion endpoint
- Azure Monitor custom logs
- Tenant-internal log collector
- A custom HTTPS POST endpoint of the admin's choice

The `.sppkg` ships with telemetry **disabled by default** (`IsEnabled: false`). No data leaves the tenant in the current runtime because telemetry is not wired. A future implementation must require both admin enablement and user opt-in before transmission.

## Opt-in / opt-out

- **Opt in** — Future runtime wiring must require `SearchTelemetryConfig.IsEnabled = true`, a destination URL, and per-user opt-in before sending data.
- **Opt out** — Future runtime wiring must stop telemetry immediately when the user opts out or when an admin sets `SearchTelemetryConfig.IsEnabled = false`.

## Data retention

The transport does not retain anything client-side beyond the in-flight batch. Retention at the destination is the admin's policy.

## Compliance

This plumbing is compliant by design with the spec section 4.3 T5 "never captured" list. Tenant-specific compliance posture (GDPR, CCPA, HIPAA, FedRAMP) depends on the destination endpoint and the admin's data processing agreement with that endpoint.

## Reporting a privacy concern

File a bug via `.github/ISSUE_TEMPLATE/bug_report.md` tagged `privacy`. Include the SearchTelemetryConfig.DestinationEndpoint value and the specific signal that surfaced the concern.
