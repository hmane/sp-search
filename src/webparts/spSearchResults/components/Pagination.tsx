import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SpSearchResults.module.scss';

export interface IPaginationProps {
  currentPage: number;
  totalCount: number;
  pageSize: number;
  onPageChange: (page: number) => void;
}

/**
 * Builds an array of page numbers to display, with ellipsis markers.
 * Uses -1 as a sentinel value for ellipsis positions.
 */
function buildPageNumbers(currentPage: number, totalPages: number): number[] {
  if (totalPages <= 7) {
    const pages: number[] = [];
    for (let i: number = 1; i <= totalPages; i++) {
      pages.push(i);
    }
    return pages;
  }

  const pages: number[] = [];

  // Always show first page
  pages.push(1);

  if (currentPage > 4) {
    pages.push(-1); // ellipsis
  }

  // Pages around current
  const start: number = Math.max(2, currentPage - 1);
  const end: number = Math.min(totalPages - 1, currentPage + 1);

  for (let i: number = start; i <= end; i++) {
    pages.push(i);
  }

  if (currentPage < totalPages - 3) {
    pages.push(-1); // ellipsis
  }

  // Always show last page
  if (pages[pages.length - 1] !== totalPages) {
    pages.push(totalPages);
  }

  return pages;
}

const Pagination: React.FC<IPaginationProps> = (props) => {
  const { currentPage, totalCount, pageSize, onPageChange } = props;
  const totalPages: number = Math.max(1, Math.ceil(totalCount / pageSize));

  if (totalPages <= 1) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null;
  }

  const pageNumbers: number[] = buildPageNumbers(currentPage, totalPages);

  const handlePrevious = React.useCallback((): void => {
    if (currentPage > 1) {
      onPageChange(currentPage - 1);
    }
  }, [currentPage, onPageChange]);

  const handleNext = React.useCallback((): void => {
    if (currentPage < totalPages) {
      onPageChange(currentPage + 1);
    }
  }, [currentPage, totalPages, onPageChange]);

  return (
    <nav className={styles.pagination} aria-label="Search results pagination">
      <button
        className={styles.pageButton}
        disabled={currentPage <= 1}
        onClick={handlePrevious}
        aria-label="Previous page"
      >
        <Icon iconName="ChevronLeft" />
      </button>

      {pageNumbers.map((page: number, index: number) => {
        if (page === -1) {
          return (
            <span key={'ellipsis-' + String(index)} className={styles.pageEllipsis} aria-hidden="true">
              &hellip;
            </span>
          );
        }

        if (page === currentPage) {
          return (
            <button
              key={page}
              className={styles.pageButtonActive}
              aria-label={'Page ' + String(page) + ', current page'}
              aria-current="page"
            >
              {page}
            </button>
          );
        }

        return (
          <button
            key={page}
            className={styles.pageButton}
            onClick={(): void => { onPageChange(page); }}
            aria-label={'Go to page ' + String(page)}
          >
            {page}
          </button>
        );
      })}

      <button
        className={styles.pageButton}
        disabled={currentPage >= totalPages}
        onClick={handleNext}
        aria-label="Next page"
      >
        <Icon iconName="ChevronRight" />
      </button>
    </nav>
  );
};

export default Pagination;
