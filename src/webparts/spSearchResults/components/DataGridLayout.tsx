import * as React from 'react';
import { ISearchResult } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const LazyDataGridContent: any = React.lazy(() => import('./DataGridContent') as any);

export interface IDataGridLayoutProps {
  items: ISearchResult[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

const DataGridLayout: React.FC<IDataGridLayoutProps> = (props) => {
  return (
    <div className={styles.dataGridContainer}>
      <React.Suspense fallback={<div className={styles.dataGridLoading}>Loading data grid...</div>}>
        <LazyDataGridContent
          items={props.items}
          onPreviewItem={props.onPreviewItem}
          onItemClick={props.onItemClick}
        />
      </React.Suspense>
    </div>
  );
};

export default DataGridLayout;
