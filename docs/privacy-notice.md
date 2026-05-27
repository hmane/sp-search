# SP Search — Privacy Notice (Telemetry)

> Owned by Foundations track (Found.D9). Read this before enabling telemetry. T5.D8 ships the schema; T5.D9 ships the aggregate dashboard view.

## What we collect (when telemetry is enabled)

When an admin enables telemetry via the `SearchTelemetryConfig` list and a user opts in via the Admin Manager Telemetry property pane group:

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

The `ITelemetrySignal` interface (T5.D8) enforces the never-captured field list at compile time. The transport (`TelemetryTransport.ts`) never inspects payload contents — it is purely the wire.

## Where the data goes

The destination is admin-configured in the `SearchTelemetryConfig` list (`DestinationEndpoint` field). Same plumbing supports:
- Application Insights ingestion endpoint
- Azure Monitor custom logs
- Tenant-internal log collector
- A custom HTTPS POST endpoint of the admin's choice

The `.sppkg` ships with telemetry **disabled by default** (`IsEnabled: false`). No data leaves the tenant unless an admin both (a) sets `IsEnabled: true` + a destination URL, and (b) at least one user opts in via the Admin Manager property pane.

## Opt-in / opt-out

- **Opt in** — Admin sets `SearchTelemetryConfig.IsEnabled = true` + a destination URL. End users see the "View what we send" Panel (T5.D8) on the property pane and can opt in per user. Opt-in events recorded in the `SearchTelemetryOptIn` list (per-user, anonymized hash only).
- **Opt out** — End users can clear their opt-in by toggling the Admin Manager property pane setting back to off; immediately stops telemetry transmission for that user. Admin can disable tenant-wide by setting `SearchTelemetryConfig.IsEnabled = false`.

## Data retention

The transport does not retain anything client-side beyond the in-flight batch. Retention at the destination is the admin's policy.

## Compliance

This plumbing is compliant by design with the spec section 4.3 T5 "never captured" list. Tenant-specific compliance posture (GDPR, CCPA, HIPAA, FedRAMP) depends on the destination endpoint and the admin's data processing agreement with that endpoint.

## Reporting a privacy concern

File a bug via `.github/ISSUE_TEMPLATE/bug_report.md` tagged `privacy`. Include the SearchTelemetryConfig.DestinationEndpoint value and the specific signal that surfaced the concern.
