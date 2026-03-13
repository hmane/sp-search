import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { PersonaCoin, PersonaSize } from '@fluentui/react/lib/Persona';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ISearchResult } from '@interfaces/index';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import styles from './SpSearchResults.module.scss';
import { GraphOrgService, IOrgPerson } from './GraphOrgService';

export interface IPeopleLayoutProps {
  items: ISearchResult[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
  graphOrgService?: GraphOrgService;
}

// ─── Property extraction helpers ──────────────────────────────────────────────
// Each helper checks Graph-native camelCase properties first (stored in
// item.properties by GraphSearchProvider), then falls back to SharePoint
// PascalCase managed property names (from SharePointSearchProvider).

/**
 * Extracts a string value from the item's properties bag.
 */
function getProperty(item: ISearchResult, key: string): string {
  const value: unknown = item.properties[key];
  return typeof value === 'string' ? value : '';
}

function getDisplayName(item: ISearchResult): string {
  return item.author?.displayText || item.title || '';
}

function getJobTitle(item: ISearchResult): string {
  const p = item.properties;
  return (
    (typeof p.jobTitle === 'string' ? p.jobTitle : '') ||
    getProperty(item, 'JobTitle') ||
    getProperty(item, 'SPS-JobTitle') ||
    ''
  );
}

function getDepartment(item: ISearchResult): string {
  const p = item.properties;
  return (
    (typeof p.department === 'string' ? p.department : '') ||
    item.siteName ||  // GraphSearchProvider maps department → siteName
    getProperty(item, 'Department') ||
    getProperty(item, 'SPS-Department') ||
    ''
  );
}

function getLocation(item: ISearchResult): string {
  const p = item.properties;
  return (
    (typeof p.officeLocation === 'string' ? p.officeLocation : '') ||
    getProperty(item, 'OfficeNumber') ||
    getProperty(item, 'BaseOfficeLocation') ||
    getProperty(item, 'SPS-Location') ||
    getProperty(item, 'Office') ||
    ''
  );
}

/**
 * Strips SharePoint claim string prefixes (e.g. "i:0#.f|membership|user@domain.com")
 * to extract the raw email address.
 */
function extractEmail(raw: string): string {
  if (!raw) {
    return '';
  }
  const pipeIdx: number = raw.lastIndexOf('|');
  return pipeIdx >= 0 ? raw.substring(pipeIdx + 1) : raw;
}

function getEmail(item: ISearchResult): string {
  if (item.author?.email) {
    return extractEmail(item.author.email);
  }
  return extractEmail(
    getProperty(item, 'WorkEmail') ||
    getProperty(item, 'SPS-SipAddress') ||
    ''
  );
}

/**
 * Work phone — Graph gives `phones: [{ type, number }]`; SP uses managed props.
 */
function getWorkPhone(item: ISearchResult): string {
  const p = item.properties;
  if (Array.isArray(p.phones)) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const biz = (p.phones as any[]).find((ph) => ph.type === 'business');
    if (biz && typeof biz.number === 'string') {
      return biz.number;
    }
  }
  return (
    getProperty(item, 'WorkPhone') ||
    getProperty(item, 'SPS-WorkPhone') ||
    getProperty(item, 'Phone') ||
    ''
  );
}

/**
 * Skills — Graph gives `skills: string[]`; SP uses pipe-delimited `SPS-Skills`.
 * Capped at 5 tags to keep the card compact.
 */
function getSkills(item: ISearchResult): string[] {
  const p = item.properties;
  if (Array.isArray(p.skills)) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (p.skills as any[])
      .filter((s) => typeof s === 'string' && s.trim())
      .slice(0, 5);
  }
  const raw = getProperty(item, 'SPS-Skills');
  if (raw) {
    return raw.split('|').map((s: string) => s.trim()).filter(Boolean).slice(0, 5);
  }
  return [];
}

/**
 * About me — truncated to 120 chars to keep cards a consistent height.
 */
function getAboutMe(item: ISearchResult): string {
  const p = item.properties;
  const raw: string =
    (typeof p.aboutMe === 'string' ? p.aboutMe : '') ||
    getProperty(item, 'AboutMe') ||
    getProperty(item, 'SPS-AboutMe') ||
    '';

  if (!raw) {
    return '';
  }
  return raw.length > 120 ? raw.substring(0, 117) + '…' : raw;
}

// ─── Action bar ────────────────────────────────────────────────────────────────

interface IPersonaActionBarProps {
  email: string;
  profileUrl: string;
  displayName: string;
}

/**
 * Horizontal quick-action row at the bottom of each persona card.
 * All links use stopPropagation so the card-click preview handler is not triggered.
 */
const PersonaActionBar: React.FC<IPersonaActionBarProps> = (barProps) => {
  const { email, profileUrl, displayName } = barProps;

  const stopBubble = React.useCallback((e: React.MouseEvent): void => {
    e.stopPropagation();
  }, []);

  const hasAnyAction = email || profileUrl;
  if (!hasAnyAction) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null;
  }

  return (
    <div className={styles.personaActionBar}>
      {email && (
        <a
          href={'https://teams.microsoft.com/l/chat/0/0?users=' + encodeURIComponent(email)}
          target="_blank"
          rel="noopener noreferrer"
          className={styles.personaActionBtn}
          title={'Chat with ' + displayName + ' in Teams'}
          onClick={stopBubble}
        >
          <Icon iconName="TeamsLogo" className={styles.personaActionIcon} />
          <span>Chat</span>
        </a>
      )}
      {email && (
        <a
          href={'mailto:' + email}
          className={styles.personaActionBtn}
          title={'Email ' + displayName}
          onClick={stopBubble}
        >
          <Icon iconName="Mail" className={styles.personaActionIcon} />
          <span>Email</span>
        </a>
      )}
      {profileUrl && (
        <a
          href={profileUrl}
          target="_blank"
          rel="noopener noreferrer"
          className={styles.personaActionBtn}
          title={'View profile for ' + displayName}
          onClick={stopBubble}
        >
          <Icon iconName="ContactInfo" className={styles.personaActionIcon} />
          <span>Profile</span>
        </a>
      )}
    </div>
  );
};

// ─── Skills tag list ───────────────────────────────────────────────────────────

const PersonaSkills: React.FC<{ skills: string[] }> = ({ skills }) => {
  if (skills.length === 0) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null;
  }
  return (
    <div className={styles.personaSkillsList}>
      {skills.map((skill, i) => (
        <span key={skill + '-' + String(i)} className={styles.personaSkillTag}>
          {skill}
        </span>
      ))}
    </div>
  );
};

// ─── Org chart section ─────────────────────────────────────────────────────────

const OrgPersonRow: React.FC<{ person: IOrgPerson }> = ({ person }) => {
  const photoUrl = person.mail
    ? '/_layouts/15/userphoto.aspx?size=S&accountname=' + encodeURIComponent(person.mail)
    : '';

  return (
    <div className={styles.orgPersonRow}>
      <PersonaCoin
        text={person.displayName}
        size={PersonaSize.size32}
        imageUrl={photoUrl || undefined}
      />
      <div className={styles.orgPersonDetails}>
        {person.mail ? (
          <a
            href={'https://teams.microsoft.com/_#/profile/' + encodeURIComponent(person.userPrincipalName || person.mail)}
            target="_blank"
            rel="noopener noreferrer"
            className={styles.orgPersonName}
          >
            {person.displayName}
          </a>
        ) : (
          <span className={styles.orgPersonName}>{person.displayName}</span>
        )}
        {person.jobTitle && (
          <span className={styles.orgPersonTitle}>{person.jobTitle}</span>
        )}
      </div>
    </div>
  );
};

type OrgLoadState = 'idle' | 'loading' | 'loaded' | 'error' | 'unavailable';

interface IOrgSectionProps {
  userId: string;
  graphOrgService: GraphOrgService;
}

/**
 * Expandable org chart panel shown at the bottom of each persona card.
 * Fetches manager and direct reports lazily on first expand.
 * Shows a graceful notice when Graph permissions are unavailable.
 */
const OrgSection: React.FC<IOrgSectionProps> = ({ userId, graphOrgService }) => {
  const [expanded, setExpanded] = React.useState(false);
  const [loadState, setLoadState] = React.useState<OrgLoadState>('idle');
  const [manager, setManager] = React.useState<IOrgPerson | null | undefined>(undefined);
  const [reports, setReports] = React.useState<IOrgPerson[]>([]);

  const toggle = React.useCallback((e: React.MouseEvent): void => {
    e.stopPropagation();
    setExpanded((prev) => !prev);
  }, []);

  React.useEffect((): void => {
    if (!expanded || loadState !== 'idle') {
      return;
    }
    setLoadState('loading');

    Promise.all([
      graphOrgService.fetchManager(userId),
      graphOrgService.fetchDirectReports(userId),
    ]).then(([mgr, rpts]): void => {
      // undefined from either call = permissions error
      if (mgr === undefined || rpts === undefined) {
        setLoadState('unavailable');
        return;
      }
      setManager(mgr);
      setReports(rpts || []);
      setLoadState('loaded');
    }).catch((): void => {
      setLoadState('error');
    });
  }, [expanded, loadState, userId, graphOrgService]);

  const stopBubble = React.useCallback((e: React.MouseEvent): void => {
    e.stopPropagation();
  }, []);

  const hasContent = loadState === 'loaded' && (manager !== null || reports.length > 0);
  const isEmpty = loadState === 'loaded' && manager === null && reports.length === 0;

  return (
    <div className={styles.orgSection} onClick={stopBubble}>
      <button
        className={styles.orgToggleBtn}
        onClick={toggle}
        type="button"
        aria-expanded={expanded}
      >
        <Icon iconName="Org" className={styles.orgToggleIcon} />
        <span>Org chart</span>
        <Icon
          iconName={expanded ? 'ChevronUp' : 'ChevronDown'}
          className={styles.orgChevron}
        />
      </button>

      {expanded && (
        <div className={styles.orgPanel}>
          {loadState === 'loading' && (
            <div className={styles.orgLoading}>
              <Spinner size={SpinnerSize.xSmall} label="Loading..." labelPosition="right" />
            </div>
          )}

          {loadState === 'error' && (
            <p className={styles.orgNotice}>Could not load org chart.</p>
          )}

          {loadState === 'unavailable' && (
            <p className={styles.orgNotice}>
              Org chart requires the <em>User.Read.All</em> Graph permission.
            </p>
          )}

          {isEmpty && (
            <p className={styles.orgNotice}>No org relationships configured.</p>
          )}

          {hasContent && (
            <>
              {manager && (
                <div className={styles.orgGroup}>
                  <span className={styles.orgGroupLabel}>Manager</span>
                  <OrgPersonRow person={manager} />
                </div>
              )}

              {reports.length > 0 && (
                <div className={styles.orgGroup}>
                  <span className={styles.orgGroupLabel}>
                    Direct reports ({reports.length})
                  </span>
                  <ul className={styles.orgReportsList}>
                    {reports.map((r) => (
                      <li key={r.id}>
                        <OrgPersonRow person={r} />
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
};

// ─── Persona card ──────────────────────────────────────────────────────────────

const PersonaCard: React.FC<{
  item: ISearchResult;
  position: number;
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
  graphOrgService?: GraphOrgService;
}> = (cardProps) => {
  const { item, position, onPreviewItem, onItemClick, graphOrgService } = cardProps;

  const displayName: string = getDisplayName(item);
  const email: string = getEmail(item);
  const jobTitle: string = getJobTitle(item);
  const department: string = getDepartment(item);
  const location: string = getLocation(item);
  const workPhone: string = getWorkPhone(item);
  const skills: string[] = getSkills(item);
  const aboutMe: string = getAboutMe(item);

  // AAD object ID or UPN — used to call Graph /users/{id}/manager and directReports
  const orgUserId: string =
    (typeof item.properties.id === 'string' ? item.properties.id : '') ||
    (typeof item.properties.userPrincipalName === 'string' ? item.properties.userPrincipalName : '') ||
    email;

  const handleCardClick = React.useCallback((): void => {
    if (onPreviewItem) {
      onPreviewItem(item);
    }
  }, [item, onPreviewItem]);

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

  const photoUrl: string = email
    ? '/_layouts/15/userphoto.aspx?size=L&accountname=' + encodeURIComponent(email)
    : '';

  // Meta line: "Department · Location"
  const metaParts: string[] = [department, location].filter(Boolean);

  return (
    <div
      className={styles.personaCard}
      role="listitem"
      tabIndex={0}
      onClick={handleCardClick}
      onKeyDown={handleKeyDown}
    >
      {/* ── Header: avatar + name/title/department ── */}
      <div className={styles.personaHeader}>
        <PersonaCoin
          text={displayName}
          size={PersonaSize.size56}
          imageUrl={photoUrl || undefined}
        />
        <div className={styles.personaDetails}>
          <h3 className={styles.personaName}>
            <DocumentTitleHoverCard item={item} position={position} onItemClick={onItemClick} disabled>
              {(handleClick): React.ReactNode => (
                <a
                  href={item.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  onClick={(e: React.MouseEvent): void => {
                    e.stopPropagation();
                    handleClick(e);
                  }}
                >
                  {displayName}
                </a>
              )}
            </DocumentTitleHoverCard>
          </h3>
          {jobTitle && <p className={styles.personaJobTitle}>{jobTitle}</p>}
          {metaParts.length > 0 && (
            <p className={styles.personaDepartment}>{metaParts.join(' · ')}</p>
          )}
        </div>
      </div>

      {/* ── About me excerpt ── */}
      {aboutMe && (
        <p className={styles.personaAboutMe}>{aboutMe}</p>
      )}

      {/* ── Contact info ── */}
      <div className={styles.personaContactInfo}>
        {email && (
          <div className={styles.personaContactItem}>
            <Icon iconName="Mail" style={{ fontSize: 13 }} />
            <a
              href={'mailto:' + email}
              className={styles.personaContactLink}
              onClick={(e: React.MouseEvent): void => { e.stopPropagation(); }}
            >
              {email}
            </a>
          </div>
        )}
        {workPhone && (
          <div className={styles.personaContactItem}>
            <Icon iconName="Phone" style={{ fontSize: 13 }} />
            <a
              href={'tel:' + workPhone}
              className={styles.personaContactLink}
              onClick={(e: React.MouseEvent): void => { e.stopPropagation(); }}
            >
              {workPhone}
            </a>
          </div>
        )}
      </div>

      {/* ── Skills/interests tags ── */}
      <PersonaSkills skills={skills} />

      {/* ── Quick-action buttons ── */}
      <PersonaActionBar
        email={email}
        profileUrl={item.url}
        displayName={displayName}
      />

      {/* ── Org chart (Graph only — hidden when service unavailable) ── */}
      {graphOrgService && orgUserId && (
        <OrgSection userId={orgUserId} graphOrgService={graphOrgService} />
      )}
    </div>
  );
};

// ─── Layout ────────────────────────────────────────────────────────────────────

/**
 * PeopleLayout — renders search results as person cards.
 *
 * Works with both SharePoint Search (managed property bags) and the Graph
 * People provider (entityTypes: ['person']). Graph-sourced results get
 * richer data (skills, about me, Teams actions); SP-sourced results fall
 * back gracefully to available managed properties.
 *
 * Grid columns:
 *  - Desktop (>= 1024px): 2 columns
 *  - Mobile (<  640px):   1 column
 */
const PeopleLayout: React.FC<IPeopleLayoutProps> = (props) => {
  const { items, onPreviewItem, onItemClick, graphOrgService } = props;

  return (
    <div className={styles.peopleGrid} role="list">
      {items.map((item: ISearchResult, index: number) => (
        <PersonaCard
          key={item.key}
          item={item}
          position={index + 1}
          onPreviewItem={onPreviewItem}
          onItemClick={onItemClick}
          graphOrgService={graphOrgService}
        />
      ))}
    </div>
  );
};

export default PeopleLayout;
