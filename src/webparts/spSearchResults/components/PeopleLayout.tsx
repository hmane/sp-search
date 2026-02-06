import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { UserPersona } from 'spfx-toolkit/lib/components/UserPersona';
import { ISearchResult } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

export interface IPeopleLayoutProps {
  items: ISearchResult[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

/**
 * Extracts a string value from the item's dynamic properties bag.
 * Returns an empty string if the property is not found or not a string.
 */
function getProperty(item: ISearchResult, key: string): string {
  const value: unknown = item.properties[key];
  if (typeof value === 'string') {
    return value;
  }
  return '';
}

/**
 * Attempts to extract a job title from common managed property names.
 */
function getJobTitle(item: ISearchResult): string {
  return (
    getProperty(item, 'JobTitle') ||
    getProperty(item, 'SPS-JobTitle') ||
    getProperty(item, 'Title') ||
    ''
  );
}

/**
 * Attempts to extract a department from common managed property names.
 */
function getDepartment(item: ISearchResult): string {
  return (
    getProperty(item, 'Department') ||
    getProperty(item, 'SPS-Department') ||
    ''
  );
}

/**
 * Attempts to extract a work location from common managed property names.
 */
function getLocation(item: ISearchResult): string {
  return (
    getProperty(item, 'OfficeNumber') ||
    getProperty(item, 'BaseOfficeLocation') ||
    getProperty(item, 'SPS-Location') ||
    getProperty(item, 'Office') ||
    ''
  );
}

/**
 * Attempts to extract a work phone number from common managed property names.
 */
function getWorkPhone(item: ISearchResult): string {
  return (
    getProperty(item, 'WorkPhone') ||
    getProperty(item, 'SPS-WorkPhone') ||
    getProperty(item, 'Phone') ||
    ''
  );
}

/**
 * Gets the person's display name — prefers author.displayText, fallback to title.
 */
function getDisplayName(item: ISearchResult): string {
  if (item.author && item.author.displayText) {
    return item.author.displayText;
  }
  return item.title || '';
}

/**
 * Gets the person's email — prefers author.email, fallback to properties.
 */
function getEmail(item: ISearchResult): string {
  if (item.author && item.author.email) {
    return item.author.email;
  }
  return (
    getProperty(item, 'WorkEmail') ||
    getProperty(item, 'SPS-SipAddress') ||
    ''
  );
}

/**
 * Single persona card rendered inside the people grid.
 */
const PersonaCard: React.FC<{
  item: ISearchResult;
  position: number;
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}> = (cardProps) => {
  const { item, position, onPreviewItem, onItemClick } = cardProps;

  const displayName: string = getDisplayName(item);
  const email: string = getEmail(item);
  const jobTitle: string = getJobTitle(item);
  const department: string = getDepartment(item);
  const location: string = getLocation(item);
  const workPhone: string = getWorkPhone(item);

  const handleCardClick = React.useCallback((): void => {
    if (onPreviewItem) {
      onPreviewItem(item);
    }
  }, [item, onPreviewItem]);

  const handleLinkClick = React.useCallback((): void => {
    if (onItemClick) {
      onItemClick(item, position);
    }
  }, [item, position, onItemClick]);

  const handleKeyDown = React.useCallback(
    (ev: React.KeyboardEvent<HTMLDivElement>): void => {
      if (ev.key === 'Enter' || ev.key === ' ') {
        ev.preventDefault();
        if (onPreviewItem) {
          onPreviewItem(item);
        }
      }
    },
    [item, onPreviewItem]
  );

  return (
    <div
      className={styles.personaCard}
      role="listitem"
      tabIndex={0}
      onClick={handleCardClick}
      onKeyDown={handleKeyDown}
    >
      <div className={styles.personaHeader}>
        <UserPersona
          userIdentifier={email}
          displayName={displayName}
          email={email}
          size={72}
          displayMode="avatarAndName"
          showLivePersona={true}
          showSecondaryText={true}
        />
      </div>
      <div className={styles.personaDetails}>
        {/* Name as link to profile */}
        <h3 className={styles.personaName}>
          <a href={item.url} target="_blank" rel="noopener noreferrer" onClick={handleLinkClick}>
            {displayName}
          </a>
        </h3>

        {/* Job title */}
        {jobTitle && (
          <p className={styles.personaJobTitle}>{jobTitle}</p>
        )}

        {/* Department */}
        {department && (
          <p className={styles.personaDepartment}>{department}</p>
        )}

        {/* Contact information */}
        <div className={styles.personaContactInfo}>
          {email && (
            <div className={styles.personaContactItem}>
              <Icon iconName="Mail" style={{ fontSize: 13 }} />
              <a href={'mailto:' + email} className={styles.personaContactLink}>
                {email}
              </a>
            </div>
          )}
          {workPhone && (
            <div className={styles.personaContactItem}>
              <Icon iconName="Phone" style={{ fontSize: 13 }} />
              <a href={'tel:' + workPhone} className={styles.personaContactLink}>
                {workPhone}
              </a>
            </div>
          )}
          {location && (
            <div className={styles.personaContactItem}>
              <Icon iconName="POI" style={{ fontSize: 13 }} />
              <span>{location}</span>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

/**
 * PeopleLayout — renders search results as person cards optimized for people searches.
 * Uses spfx-toolkit UserPersona for consistent avatar display with automatic
 * profile fetching, photo loading, and LivePersona hover card support.
 *
 * Grid columns:
 *  - Desktop (>= 1024px): 2 columns
 *  - Mobile (< 640px): 1 column
 */
const PeopleLayout: React.FC<IPeopleLayoutProps> = (props) => {
  const { items, onPreviewItem, onItemClick } = props;

  return (
    <div className={styles.peopleGrid} role="list">
      {items.map((item: ISearchResult, index: number) => (
        <PersonaCard
          key={item.key}
          item={item}
          position={index + 1}
          onPreviewItem={onPreviewItem}
          onItemClick={onItemClick}
        />
      ))}
    </div>
  );
};

export default PeopleLayout;
