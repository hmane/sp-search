import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import { UserPersona as _UserPersona } from 'spfx-toolkit/lib/components/UserPersona';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const UserPersona: any = _UserPersona;
import { ISearchResult } from '@interfaces/index';
import { formatFileSize, formatRelativeDate, formatUrlBreadcrumb, sanitizeSummaryHtml, formatDateTime, getResultAnchorProps, formatTitleText, TitleDisplayMode } from './documentTitleUtils';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import AddToCollectionButton from './AddToCollectionButton';
import styles from './SpSearchResults.module.scss';

export interface IListLayoutProps {
  items: ISearchResult[];
  searchContextId: string;
  titleDisplayMode: TitleDisplayMode;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

const ListLayout: React.FC<IListLayoutProps> = (props) => {
  const { items, searchContextId, titleDisplayMode, onItemClick } = props;

  return (
    <ul className={styles.resultList} role="list">
      {items.map((item: ISearchResult, index: number) => {
        const sizeDisplay: string = formatFileSize(item.fileSize);
        const linkProps = getResultAnchorProps(item);

        return (
          <li key={item.key} className={styles.resultCard} role="listitem">

            <div className={styles.resultIcon}>
              <FileTypeIcon type={IconType.image} path={item.url} size={ImageSize.medium} />
            </div>

            <div className={styles.resultBody}>
              <h3 className={styles.resultTitle}>
                <div className={styles.resultTitleRow}>
                  <DocumentTitleHoverCard item={item} position={index + 1} onItemClick={onItemClick}>
                    {(handleClick): React.ReactNode => (
                      <a
                        href={linkProps.href}
                        target={linkProps.target}
                        rel={linkProps.rel}
                        className={titleDisplayMode === 'wrap' ? styles.resultTitleLinkWrap : styles.resultTitleLink}
                        onClick={handleClick}
                      >
                        {formatTitleText(item.title, titleDisplayMode)}
                      </a>
                    )}
                  </DocumentTitleHoverCard>
                  <div className={styles.resultTitleActions}>
                    <AddToCollectionButton
                      item={item}
                      searchContextId={searchContextId}
                    />
                    {item.fileType && (
                      <span className={styles.resultFileTypeBadge}>
                        {item.fileType.toUpperCase()}
                      </span>
                    )}
                  </div>
                </div>
              </h3>
              <p className={styles.resultUrl}>{formatUrlBreadcrumb(item.url)}</p>
              {item.summary && (
                <div
                  className={styles.resultSummary}
                  dangerouslySetInnerHTML={{ __html: sanitizeSummaryHtml(item.summary) }}
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
          </li>
        );
      })}
    </ul>
  );
};

export default ListLayout;
