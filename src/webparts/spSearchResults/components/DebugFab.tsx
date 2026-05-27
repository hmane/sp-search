import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DebugCollector } from '@store/debug';
import styles from './DebugPanel.module.scss';

export interface IDebugFabProps {
  onClick: () => void;
}

const DebugFab: React.FC<IDebugFabProps> = ({ onClick }) => {
  const [hasError, setHasError] = React.useState(false);

  React.useEffect(() => {
    return DebugCollector.subscribe(() => {
      const events = DebugCollector.getEvents();
      const recentError = events.length > 0 &&
        events[0].type === 'ERROR' &&
        Date.now() - events[0].timestamp < 5000;
      setHasError(!!recentError);
    });
  }, []);

  return (
    <button
      className={`${styles.debugFab}${hasError ? ' ' + styles.hasError : ''}`}
      onClick={onClick}
      title="SP Search Debug Panel"
      type="button"
    >
      <Icon iconName="Bug" />
    </button>
  );
};

export default DebugFab;
