import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SpSearchVerticals.module.scss';

export interface IVerticalTabProps {
  verticalKey: string;
  label: string;
  iconName: string | undefined;
  count: number | undefined;
  isActive: boolean;
  showCounts: boolean;
  isDimmed: boolean;
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

  return (
    <button
      className={classNames.join(' ')}
      role="tab"
      aria-selected={isActive}
      aria-label={label + (hasCount ? ' (' + String(count) + ')' : '')}
      tabIndex={isActive ? 0 : -1}
      onClick={handleClick}
      onKeyDown={handleKeyDown}
      data-vertical-key={verticalKey}
    >
      {iconName && (
        <span className={styles.tabIcon}>
          <Icon iconName={iconName} />
        </span>
      )}
      <span className={styles.tabLabel}>{label}</span>
      {hasCount && (
        <span className={styles.countBadge}>{String(count)}</span>
      )}
    </button>
  );
};

export default VerticalTab;
