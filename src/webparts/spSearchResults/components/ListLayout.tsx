import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import { UserPersona as _UserPersona } from 'spfx-toolkit/lib/components/UserPersona';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const UserPersona: any = _UserPersona;
import { ISearchResult, ISearchScope } from '@interfaces/index';
import { sanitizeHtml } from 'spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml';
import { formatFileSize, formatRelativeDate, formatUrlBreadcrumb, formatDateTime, formatTitleText, isImageType, TitleDisplayMode } from './documentTitleUtils';
import { resolveResultLink, type IResultLinkConfig } from './resultLink';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import AddToCollectionButton from './AddToCollectionButton';
import { buildRowActionMenu } from './buildRowActionMenu';
import styles from './SpSearchResults.module.scss';

export interface IListLayoutProps {
  items: ISearchResult[];
  searchContextId: string;
  scope: ISearchScope;
  titleDisplayMode: TitleDisplayMode;
  onItemClick?: (item: ISearchResult, position: number) => void;
  // Stream C / #7
  linkConfig: IResultLinkConfig;
  onOpenInSidePanel?: (item: ISearchResult) => void;
}

const ListLayout: React.FC<IListLayoutProps> = (props) => {
  const { items, searchContextId, scope, titleDisplayMode, onItemClick, linkConfig, onOpenInSidePanel } = props;

  return (
    <ul className={styles.resultList} role="list">
      {items.map((item: ISearchResult, index: number) => {
        const sizeDisplay: string = formatFileSize(item.fileSize);
        const linkProps = resolveResultLink(item, linkConfig);
        const breadcrumbBaseUrl = getCurrentSiteBreadcrumbBaseUrl(scope, item);
        const breadcrumbText = formatUrlBreadcrumb(item.url, { baseUrl: breadcrumbBaseUrl });

        return (
          <li
            key={item.key}
            className={styles.resultCard}
            role="listitem"
          >
            <div className={styles.resultIcon}>
              {isImageType(item) && item.thumbnailUrl ? (
                // Stream C / #8 — show the image itself for image results.
                <img className={styles.resultIconImage} src={item.thumbnailUrl} alt="" loading="lazy" />
              ) : (
                <Icon {...getFileTypeIconProps({ extension: item.fileType || '', size: 32 })} />
              )}
            </div>

            <div className={styles.resultBody}>
              <h3 className={styles.resultTitle}>
                <div className={styles.resultTitleRow}>
                  <DocumentTitleHoverCard
                    item={item}
                    position={index + 1}
                    onItemClick={onItemClick}
                    clickTarget={linkConfig.clickTarget}
                    onOpenInSidePanel={onOpenInSidePanel}
                  >
                    {(handleClick): React.ReactNode => (
                      <a
                        href={linkProps.href}
                        target={linkProps.target}
                        rel={linkProps.rel}
                        data-interception="off"
                        className={titleDisplayMode === 'wrap' ? styles.resultTitleLinkWrap : styles.resultTitleLink}
                        onClick={handleClick}
                      >
                        {formatTitleText(item.title, titleDisplayMode)}
                      </a>
                    )}
                  </DocumentTitleHoverCard>
                </div>
              </h3>
              <p className={styles.resultUrl} title={item.url}>{breadcrumbText}</p>
              {item.summary && (
                <div
                  className={styles.resultSummary}
                  dangerouslySetInnerHTML={{ __html: sanitizeHtml(item.summary) }}
                />
              )}
              <div className={styles.resultMeta}>
                {item.author && item.author.displayText && (
                  <span className={styles.metaItem}>
                    <UserPersona
                      userIdentifier={item.author.email || item.author.displayText}
                      displayName={item.author.displayText}
                      size={24}
                      displayMode="avatarAndName"
                    />
                  </span>
                )}
                {item.modified && (
                  <>
                    <span className={styles.metaSeparator} />
                    <span className={styles.metaItem} title={formatDateTime(item.modified)}>
                      <Icon iconName="Clock" style={{ fontSize: 12 }} />
                      {formatRelativeDate(item.modified)}
                    </span>
                  </>
                )}
                {item.siteName && (
                  <>
                    <span className={styles.metaSeparator} />
                    <span className={styles.metaItem}>
                      <Icon iconName="SharePointLogo" style={{ fontSize: 12 }} />
                      {item.siteName}
                    </span>
                  </>
                )}
                {sizeDisplay && (
                  <>
                    <span className={styles.metaSeparator} />
                    <span className={styles.metaItem}>
                      {sizeDisplay}
                    </span>
                  </>
                )}
              </div>
            </div>
            <div className={styles.resultRowEcb}>
              <AddToCollectionButton
                item={item}
                searchContextId={searchContextId}
                triggerRenderer={(openAddToCollection): React.ReactNode => (
                  <IconButton
                    iconProps={{ iconName: 'MoreVertical' }}
                    ariaLabel={'More actions for ' + item.title}
                    title="More actions"
                    menuProps={{
                      items: buildRowActionMenu(item, {
                        position: index + 1,
                        onItemClick,
                        onAddToCollection: openAddToCollection,
                      }),
                    }}
                  />
                )}
              />
            </div>
          </li>
        );
      })}
    </ul>
  );
};

function getCurrentSiteBreadcrumbBaseUrl(scope: ISearchScope, item: ISearchResult): string | undefined {
  if (!scope || scope.id !== 'currentsite') {
    return undefined;
  }

  const kqlPath = scope.kqlPath || '';
  const quotedMatch = /Path:"([^"]+)"/i.exec(kqlPath);
  if (quotedMatch && quotedMatch[1]) {
    return quotedMatch[1];
  }

  const unquotedMatch = /Path:([^\s]+)/i.exec(kqlPath);
  if (unquotedMatch && unquotedMatch[1]) {
    return unquotedMatch[1];
  }

  return item.siteUrl || undefined;
}

export default ListLayout;
