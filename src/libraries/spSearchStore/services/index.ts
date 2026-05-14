export type { ITokenContext } from './TokenService';
export { TokenService } from './TokenService';
export { SearchService } from './SearchService';
export { SearchManagerService } from './SearchManagerService';
export { evaluatePromotedResults } from './PromotedResultsService';
export { resolveUserGroupIds, isInAudience } from './AudienceService';
export { fetchManagedProperties, getCachedSchema } from './SchemaService';
export type { ISchemaResult } from './SchemaService';
export { CoverageStatsService } from './CoverageStatsService';
export type { ICoverageConfig, ICoverageStatsResult } from './CoverageStatsService';
// T4.D9 — tenant-readiness pre-flight scan.
export { runTenantReadinessScan } from './TenantReadinessService';
export type {
  IReadinessReport,
  IReadinessCheck,
  IReadinessFixLink,
  ReadinessStatus,
} from './TenantReadinessService';
