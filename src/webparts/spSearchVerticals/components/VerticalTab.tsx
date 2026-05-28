import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { TooltipHost, ITooltipHostStyles } from '@fluentui/react/lib/Tooltip';
import styles from './SpSearchVerticals.module.scss';

const TOOLTIP_HOST_STYLES: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };

export interface IVerticalTabProps {
  verticalKey: string;
  label: string;
  iconName: string | undefined;
  count: number | undefined;
  isActive: boolean;
  showCounts: boolean;
  isDimmed: boolean;
  isLink: boolean;
  linkUrl: string | undefined;
  openBehavior: 'currentTab' | 'newTab' | undefined;
  onClick: (key: string) => void;
}

const VerticalTab: React.FC<IVerticalTabProps> = (props: IVerticalTabProps): React.ReactElement => {
  const {
    verticalKey,
    label,
    iconName,
    count,
    isActive,
    showCounts,
    isDimmed,
    isLink,
    linkUrl,
    openBehavior,
    onClick
  } = props;

  const handleClick = React.useCallback(function (): void {
    if (!isDimmed) {
      onClick(verticalKey);
    }
  }, [isDimmed, onClick, verticalKey]);

  const handleKeyDown = React.useCallback(function (ev: React.KeyboardEvent<HTMLButtonElement>): void {
    if (ev.key === 'Enter' || ev.key === ' ') {
      ev.preventDefault();
      if (!isDimmed) {
        onClick(verticalKey);
      }
    }
  }, [isDimmed, onClick, verticalKey]);

  const classNames: string[] = [styles.verticalTab];
  if (isActive) {
    classNames.push(styles.active);
  }
  if (isDimmed) {
    classNames.push(styles.dimmed);
  }

  const hasCount: boolean = showCounts && count !== undefined;

  const tabContent = (
    <React.Fragment>
      {iconName && (
        <span className={styles.tabIcon}>
          <Icon iconName={iconName} />
        </span>
      )}
      <span className={styles.tabLabel}>{label}</span>
      {hasCount && (
        <span className={styles.countBadge}>{String(count)}</span>
      )}
    </React.Fragment>
  );

  // Render as <a> for link tabs. When dimmed, suppress navigation —
  // `aria-disabled` keeps the tab focusable (so the tooltip can announce
  // why), and the onClick guard stops the browser from navigating.
  if (isLink && linkUrl) {
    const linkEl = (
      <a
        className={classNames.join(' ')}
        href={isDimmed ? undefined : linkUrl}
        target={openBehavior === 'newTab' ? '_blank' : '_self'}
        rel={openBehavior === 'newTab' ? 'noopener noreferrer' : undefined}
        aria-label={label + (isDimmed ? ' (no results)' : '')}
        aria-disabled={isDimmed || undefined}
        role={isDimmed ? 'link' : undefined}
        tabIndex={isDimmed ? 0 : undefined}
        title={isDimmed ? 'No results in this vertical for the current query.' : undefined}
        onClick={isDimmed ? ((ev: React.MouseEvent<HTMLAnchorElement>): void => { ev.preventDefault(); }) : undefined}
        data-vertical-key={verticalKey}
      >
        {tabContent}
      </a>
    );

    if (isDimmed) {
      return (
        <TooltipHost
          content="No results in this vertical for the current query."
          styles={TOOLTIP_HOST_STYLES}
        >
          {linkEl}
        </TooltipHost>
      );
    }
    return linkEl;
  }

  // `aria-disabled` instead of `disabled` so the dimmed tab still receives
  // keyboard focus — disabled buttons can't be focused in most browsers,
  // and a tooltip on an unfocusable element fails WCAG 2.4.3. The onClick
  // guard at the top of the file already blocks activation when dimmed.
  const buttonEl = (
    <button
      className={classNames.join(' ')}
      role="tab"
      aria-selected={isActive}
      aria-label={label + (hasCount ? ' (' + String(count) + ')' : '') + (isDimmed ? ' (no results)' : '')}
      aria-disabled={isDimmed || undefined}
      tabIndex={isActive ? 0 : -1}
      title={isDimmed ? 'No results in this vertical for the current query.' : undefined}
      onClick={handleClick}
      onKeyDown={handleKeyDown}
      data-vertical-key={verticalKey}
    >
      {tabContent}
    </button>
  );

  // T1.D8 — explain why dimmed tabs aren't clickable. TooltipHost shows on
  // hover AND focus, so keyboard users now hear the message too.
  if (isDimmed) {
    return (
      <TooltipHost
        content="No results in this vertical for the current query."
        styles={TOOLTIP_HOST_STYLES}
      >
        {buttonEl}
      </TooltipHost>
    );
  }

  return buttonEl;
};

export default VerticalTab;
