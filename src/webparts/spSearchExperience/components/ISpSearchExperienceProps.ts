import * as React from 'react';

export type FiltersPlacement = 'right' | 'left' | 'top';

export interface ISpSearchExperienceProps {
  resultsElement: React.ReactElement;
  filtersElement: React.ReactElement;
  filtersPlacement: FiltersPlacement;
  filtersWidth: number;
}
