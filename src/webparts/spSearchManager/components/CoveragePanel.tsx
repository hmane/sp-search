import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import {
  ICoverageProfile,
  ICoverageResult,
  IDiscoveredCoverageProfile,
  SearchCoverageService
} from '@services/SearchCoverageService';
import styles from './SpSearchManager.module.scss';

export interface ICoveragePanelProps {
  profiles: ICoverageProfile[];
  searchContextId?: string;
  sourcePageUrl?: string;
}

function formatNumber(value: number): string {
  return value.toLocaleString();
}

function formatDate(value: Date | undefined): string {
  if (!value || isNaN(value.getTime())) {
    return 'Unknown';
  }

  return value.toLocaleString();
}

function deriveWebUrlFromPageUrl(pageUrl: string): string | undefined {
  try {
    const parsed = new URL(pageUrl, window.location.origin);
    const marker = '/SitePages/';
    const markerIndex = parsed.pathname.toLowerCase().lastIndexOf(marker.toLowerCase());
    if (markerIndex >= 0) {
      return parsed.origin + parsed.pathname.substring(0, markerIndex);
    }
    return parsed.origin + parsed.pathname.replace(/\/[^/]+$/, '');
  } catch {
    return undefined;
  }
}

const CoveragePanel: React.FC<ICoveragePanelProps> = (props) => {
  const { profiles, searchContextId, sourcePageUrl } = props;
  const coverageService = React.useMemo(function (): SearchCoverageService {
    return new SearchCoverageService();
  }, []);
  const [discoveredProfile, setDiscoveredProfile] = React.useState<IDiscoveredCoverageProfile | undefined>(undefined);
  const [discoveryError, setDiscoveryError] = React.useState<string | undefined>(undefined);
  const [isDiscovering, setIsDiscovering] = React.useState<boolean>(false);

  React.useEffect(function (): (() => void) | void {
    const effectivePageUrl = sourcePageUrl || window.location.href;
    if (!effectivePageUrl) {
      setDiscoveredProfile(undefined);
      setDiscoveryError(undefined);
      setIsDiscovering(false);
      return;
    }

    const abortController = new AbortController();

    setIsDiscovering(true);
    setDiscoveryError(undefined);

    coverageService.discoverCoverageProfileFromPage(effectivePageUrl, searchContextId)
      .then(function (profile): void {
        if (abortController.signal.aborted) {
          return;
        }
        setDiscoveredProfile(profile);
        setIsDiscovering(false);
      })
      .catch(function (error): void {
        if (abortController.signal.aborted) {
          return;
        }
        setDiscoveredProfile(undefined);
        setDiscoveryError(error instanceof Error ? error.message : 'Failed to inspect the target search page');
        setIsDiscovering(false);
      });

    return function cleanup(): void {
      abortController.abort();
    };
  }, [coverageService, searchContextId, sourcePageUrl]);

  const usableDiscoveredProfile = React.useMemo(function (): IDiscoveredCoverageProfile | undefined {
    if (!discoveredProfile || discoveredProfile.profile.sourceUrls.length === 0) {
      return undefined;
    }
    return discoveredProfile;
  }, [discoveredProfile]);

  const resolvedProfiles = React.useMemo(function (): ICoverageProfile[] {
    if (!usableDiscoveredProfile) {
      return profiles;
    }

    const manualProfiles = profiles.filter(function (profile): boolean {
      return profile.id !== usableDiscoveredProfile.profile.id;
    });

    return [usableDiscoveredProfile.profile].concat(manualProfiles);
  }, [profiles, usableDiscoveredProfile]);

  const profileOptions = React.useMemo(function (): IDropdownOption[] {
    return resolvedProfiles.map(function (profile): IDropdownOption {
      return {
        key: profile.id,
        text: profile.title
      };
    });
  }, [resolvedProfiles]);

  const [selectedProfileId, setSelectedProfileId] = React.useState<string | undefined>(
    resolvedProfiles[0]?.id
  );
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [result, setResult] = React.useState<ICoverageResult | undefined>(undefined);
  const effectivePageUrl = sourcePageUrl || window.location.href;
  const reindexSettingsUrl = React.useMemo(function (): string | undefined {
    const webUrl = deriveWebUrlFromPageUrl(effectivePageUrl);
    return webUrl ? webUrl.replace(/\/$/, '') + '/_layouts/srchvis.aspx' : undefined;
  }, [effectivePageUrl]);

  const selectedProfile = React.useMemo(function (): ICoverageProfile | undefined {
    return resolvedProfiles.find(function (profile): boolean {
      return profile.id === selectedProfileId;
    }) || resolvedProfiles[0];
  }, [resolvedProfiles, selectedProfileId]);

  const loadResult = React.useCallback(function (): () => void {
    if (!selectedProfile) {
      setResult(undefined);
      setIsLoading(false);
      return function noop(): void { /* noop */ };
    }

    const abortController = new AbortController();

    setIsLoading(true);
    setError(undefined);

    coverageService.evaluateProfile(selectedProfile, abortController.signal)
      .then(function (nextResult): void {
        if (abortController.signal.aborted) {
          return;
        }

        setResult(nextResult);
        setIsLoading(false);
      })
      .catch(function (loadError): void {
        if (abortController.signal.aborted) {
          return;
        }

        setError(loadError instanceof Error ? loadError.message : 'Failed to load coverage diagnostics');
        setIsLoading(false);
      });

    return function cleanup(): void {
      abortController.abort();
    };
  }, [coverageService, selectedProfile]);

  React.useEffect(function (): void {
    if (!selectedProfileId && resolvedProfiles[0]) {
      setSelectedProfileId(resolvedProfiles[0].id);
    }
  }, [resolvedProfiles, selectedProfileId]);

  React.useEffect(function (): void {
    if (selectedProfileId && resolvedProfiles.some(function (profile): boolean {
      return profile.id === selectedProfileId;
    })) {
      return;
    }
    setSelectedProfileId(resolvedProfiles[0]?.id);
  }, [resolvedProfiles, selectedProfileId]);

  React.useEffect(function (): (() => void) | void {
    if (!selectedProfile) {
      setResult(undefined);
      setIsLoading(false);
      return;
    }

    return loadResult();
  }, [loadResult, selectedProfile]);

  function handleProfileChange(
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    if (option) {
      setSelectedProfileId(option.key as string);
    }
  }

  if (resolvedProfiles.length === 0 && !isDiscovering) {
    return (
      <div className={styles.healthPanel}>
        <div className={styles.emptyState}>
          <div className={styles.emptyIcon}>
            <Icon iconName="DatabaseSync" />
          </div>
          <h3 className={styles.emptyTitle}>No coverage profiles configured</h3>
          <p className={styles.emptyDescription}>
            Configure a target search page URL for auto-detection or add one or more manual coverage profiles with the exact source paths you want to monitor.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.healthPanel}>
      <div className={styles.coverageToolbar}>
        <div className={styles.coverageProfilePicker}>
          <Dropdown
            label="Coverage profile"
            selectedKey={selectedProfile?.id}
            options={profileOptions}
            onChange={handleProfileChange}
          />
        </div>
        <DefaultButton
          iconProps={{ iconName: 'Refresh' }}
          text="Refresh"
          onClick={function (): void { loadResult(); }}
          disabled={!selectedProfile || isLoading}
        />
        <DefaultButton
          iconProps={{ iconName: 'Settings' }}
          text="Open Site Reindex Settings"
          href={reindexSettingsUrl}
          target="_blank"
          disabled={!reindexSettingsUrl}
        />
      </div>

      {selectedProfile?.description && (
        <p className={styles.coverageProfileDescription}>{selectedProfile.description}</p>
      )}

      {isDiscovering && (
        <div className={styles.errorContainer}>
          <MessageBar messageBarType={MessageBarType.info}>
            Inspecting the target search page to auto-detect query scope.
          </MessageBar>
        </div>
      )}

      {discoveryError && (
        <div className={styles.errorContainer}>
          <MessageBar messageBarType={MessageBarType.warning}>
            Could not auto-detect Search Results configuration from the target page: {discoveryError}
          </MessageBar>
        </div>
      )}

      {!discoveryError && discoveredProfile && !usableDiscoveredProfile && (
        <div className={styles.errorContainer}>
          <MessageBar messageBarType={MessageBarType.warning}>
            Search Results configuration was detected, but the source scope could not be inferred precisely. Add manual source profiles or target a page with current-site or custom path scope.
          </MessageBar>
        </div>
      )}

      {error && (
        <div className={styles.errorContainer} role="alert">
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
          >
            {error}
          </MessageBar>
        </div>
      )}

      {isLoading && (
        <div className={styles.loadingContainer}>
          <Spinner size={SpinnerSize.large} label="Loading coverage diagnostics..." />
        </div>
      )}

      {!isLoading && result && (
        <>
          {result.warnings.map(function (warning, index): React.ReactElement {
            return (
              <div key={result.profile.id + '-warning-' + String(index)} className={styles.errorContainer}>
                <MessageBar messageBarType={MessageBarType.warning}>
                  {warning}
                </MessageBar>
              </div>
            );
          })}
          {usableDiscoveredProfile && selectedProfile?.id === usableDiscoveredProfile.profile.id && (
            <div className={styles.errorContainer}>
              <MessageBar messageBarType={MessageBarType.info}>
                Auto-detected from {usableDiscoveredProfile.sourcePageUrl}. Add manual profiles if you need narrower or list-specific comparisons.
              </MessageBar>
            </div>
          )}

          <div className={styles.coverageSummaryNarrative}>
            <span className={styles.coverageSummaryNarrativeLabel}>What this means</span>
            <p className={styles.coverageSummaryNarrativeText}>
              Coverage is currently using the
              {' '}
              <strong>{result.profile.trimDuplicates ? 'trimmed' : 'non-trimmed'}</strong>
              {' '}
              search count of
              {' '}
              <strong>{formatNumber(result.searchCount)}</strong>
              {' '}
              against
              {' '}
              <strong>{formatNumber(result.sourceCount)}</strong>
              {' '}
              source items.
              {result.duplicateDelta > 0 && (
                <>
                  {' '}
                  If duplicate collapsing is applied, the visible result count drops to
                  {' '}
                  <strong>{formatNumber(result.searchCountTrimmed)}</strong>
                  , which hides
                  {' '}
                  <strong>{formatNumber(result.duplicateDelta)}</strong>
                  {' '}
                  duplicate result
                  {result.duplicateDelta === 1 ? '' : 's'}.
                </>
              )}
            </p>
            {reindexSettingsUrl && (
              <div className={styles.coverageNarrativeActions}>
                <DefaultButton
                  iconProps={{ iconName: 'OpenInNewWindow' }}
                  text="Open site reindex settings"
                  href={reindexSettingsUrl}
                  target="_blank"
                />
              </div>
            )}
          </div>

          <div className={styles.coverageSummarySections}>
            <div className={styles.coverageSummarySection}>
              <div className={styles.coverageSectionHeader}>
                <div className={styles.coverageSectionHeaderTop}>
                  <div>
                    <span className={styles.coverageSectionEyebrow}>Coverage</span>
                    <h3 className={styles.coverageSectionTitle}>Coverage gap</h3>
                  </div>
                  <div className={styles.coverageSectionMetric}>
                    <span className={styles.coverageSectionMetricLabel}>Gap</span>
                    <strong className={styles.coverageSectionMetricValue}>{String(result.deltaPercent)}%</strong>
                  </div>
                </div>
                <p className={styles.coverageSectionDescription}>
                  This section shows the count currently used for the coverage comparison.
                </p>
              </div>
              <div className={styles.coverageSummaryCards}>
                <div className={styles.coverageSummaryCard}>
                  <span className={styles.coverageSummaryLabel}>Source items</span>
                  <strong className={styles.coverageSummaryValue}>{formatNumber(result.sourceCount)}</strong>
                  <span className={styles.coverageSummaryHint}>Items found directly in the configured lists and libraries.</span>
                </div>
                <div className={styles.coverageSummaryCard}>
                  <span className={styles.coverageSummaryLabel}>Indexed for comparison</span>
                  <strong className={styles.coverageSummaryValue}>{formatNumber(result.searchCount)}</strong>
                  <span className={styles.coverageSummaryHint}>
                    {result.profile.trimDuplicates ? 'Uses trimmed search results.' : 'Uses non-trimmed search results.'}
                  </span>
                </div>
                <div className={styles.coverageSummaryCard}>
                  <span className={styles.coverageSummaryLabel}>Missing from search</span>
                  <strong className={styles.coverageSummaryValue}>{formatNumber(result.delta)}</strong>
                  <span className={styles.coverageSummaryHint}>Source items that are not represented in the active search count.</span>
                </div>
              </div>
            </div>

            <div className={styles.coverageSummarySection}>
              <div className={styles.coverageSectionHeader}>
                <span className={styles.coverageSectionEyebrow}>Duplicate Analysis</span>
                <h3 className={styles.coverageSectionTitle}>Duplicate impact</h3>
                <p className={styles.coverageSectionDescription}>
                  Compare raw search results with duplicate-collapsed results to see how much count is lost to duplicate trimming.
                </p>
              </div>
              <div className={styles.coverageSummaryCards}>
                <div className={styles.coverageSummaryCard}>
                  <span className={styles.coverageSummaryLabel}>Raw search count</span>
                  <strong className={styles.coverageSummaryValue}>
                    {formatNumber(result.searchCountUntrimmed)}
                  </strong>
                  <span className={styles.coverageSummaryHint}>Search count before duplicate collapsing.</span>
                </div>
                <div className={styles.coverageSummaryCard}>
                  <span className={styles.coverageSummaryLabel}>Trimmed search count</span>
                  <strong className={styles.coverageSummaryValue}>{formatNumber(result.searchCountTrimmed)}</strong>
                  <span className={styles.coverageSummaryHint}>Search count after duplicate collapsing.</span>
                </div>
                <div className={styles.coverageSummaryCard}>
                  <span className={styles.coverageSummaryLabel}>Hidden by duplicate trim</span>
                  <strong className={styles.coverageSummaryValue}>{formatNumber(result.duplicateDelta)}</strong>
                  <span className={styles.coverageSummaryHint}>Difference between raw and trimmed result counts.</span>
                </div>
              </div>
            </div>
          </div>

          <div className={styles.coverageQueryBox}>
            <div className={styles.coverageQueryRow}>
              <span className={styles.coverageQueryLabel}>Query template</span>
              <code className={styles.coverageQueryValue}>{result.executedQueryTemplate}</code>
            </div>
            <div className={styles.coverageQueryRow}>
              <span className={styles.coverageQueryLabel}>Coverage query</span>
              <code className={styles.coverageQueryValue}>{result.executedQueryText}</code>
            </div>
            <div className={styles.coverageQueryMeta}>
              Checked {formatDate(result.checkedAt)} across {String(result.sourceResults.length)} configured source
              {result.sourceResults.length === 1 ? '' : 's'}.
            </div>
          </div>

          <div className={styles.coverageTable} role="table" aria-label="Coverage diagnostics by source">
            <div className={styles.coverageHeader} role="row">
              <div role="columnheader">Source</div>
              <div role="columnheader">Source items</div>
              <div role="columnheader">Indexed</div>
              <div role="columnheader">Trimmed</div>
              <div role="columnheader">With duplicates</div>
              <div role="columnheader">Duplicates</div>
              <div role="columnheader">Delta</div>
              <div role="columnheader">Notes</div>
            </div>
            {result.sourceResults.map(function (source): React.ReactElement {
              const notes: string[] = [];
              if (source.noCrawl) {
                notes.push('NoCrawl');
              }
              if (source.hidden) {
                notes.push('Hidden list');
              }
              if (notes.length === 0) {
                notes.push('In scope');
              }

              return (
                <div key={source.sourceUrl} className={styles.coverageRow} role="row">
                  <div role="cell" className={styles.coverageSourceCell}>
                    <div className={styles.coverageSourceTitle}>{source.title}</div>
                    <div className={styles.coverageSourceUrl}>{source.sourceUrl}</div>
                  </div>
                  <div role="cell">{formatNumber(source.sourceCount)}</div>
                  <div role="cell">{formatNumber(source.searchCount)}</div>
                  <div role="cell">{formatNumber(source.searchCountTrimmed)}</div>
                  <div role="cell">{formatNumber(source.searchCountUntrimmed)}</div>
                  <div role="cell">{formatNumber(source.duplicateDelta)}</div>
                  <div role="cell" className={source.delta > 0 ? styles.coverageDeltaWarning : undefined}>
                    {formatNumber(source.delta)}
                  </div>
                  <div role="cell" className={styles.coverageNotesCell}>
                    {notes.join(', ')}
                  </div>
                </div>
              );
            })}
          </div>

          <div className={styles.coverageMissingSection}>
            <h3 className={styles.insightSectionTitle}>
              <Icon iconName="SearchIssue" className={styles.insightSectionIcon} />
              Sample source items missing from search
            </h3>
            {result.missingSamples.length === 0 ? (
              <p className={styles.insightNoData}>
                No missing items were found in the recent sample that was checked.
              </p>
            ) : (
              <div className={styles.coverageMissingList}>
                {result.missingSamples.map(function (item, index): React.ReactElement {
                  return (
                    <div key={item.path + '-' + String(index)} className={styles.coverageMissingItem}>
                      <div className={styles.coverageMissingTitle}>{item.title}</div>
                      <div className={styles.coverageMissingMeta}>
                        {item.sourceTitle} | {formatDate(item.modified)}
                      </div>
                      <div className={styles.coverageMissingPath}>{item.path}</div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
};

export default CoveragePanel;
