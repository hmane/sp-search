import * as React from 'react';
import styles from './SpSearchExperience.module.scss';
import type { FiltersPlacement, ISpSearchExperienceProps } from './ISpSearchExperienceProps';

function getPlacementClassName(filtersPlacement: FiltersPlacement): string {
  if (filtersPlacement === 'left') {
    return styles.placementLeft;
  }
  if (filtersPlacement === 'top') {
    return styles.placementTop;
  }
  return '';
}

const SpSearchExperience: React.FC<ISpSearchExperienceProps> = (props): React.ReactElement => {
  const filterWidth = Math.max(260, Math.min(520, props.filtersWidth || 360));
  const rootStyle = {
    '--spSearchExperienceFilterWidth': filterWidth + 'px',
  } as React.CSSProperties;

  return (
    <div
      className={[styles.spSearchExperience, getPlacementClassName(props.filtersPlacement)].filter(Boolean).join(' ')}
      style={rootStyle}
    >
      <main className={styles.resultsPane}>
        {props.resultsElement}
      </main>
      <aside className={styles.filtersPane}>
        {props.filtersElement}
      </aside>
    </div>
  );
};

export default SpSearchExperience;
