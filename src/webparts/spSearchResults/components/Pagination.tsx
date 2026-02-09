import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SpSearchResults.module.scss';

export interface IPaginationProps {
  currentPage: number;
  totalCount: number;
  pageSize: number;
  showPaging: boolean;
  pageRange: number;
  onPageChange: (page: number) => void;
}

/**
 * Builds an array of page numbers to display, with ellipsis markers.
 * Uses -1 as a sentinel value for ellipsis positions.
 */
function buildPageNumbers(currentPage: number, totalPages: number, pageRange: number): number[] {
  const maxButtons = Math.max(3, pageRange);

  if (totalPages <= maxButtons) {
    const pages: number[] = [];
    for (let i: number = 1; i <= totalPages; i++) {
      pages.push(i);
    }
    return pages;
  }

  const pages: number[] = [];
  const sideCount: number = Math.floor((maxButtons - 3) / 2); // pages on each side of current (excluding first, last, current)

  // Always show first page
  pages.push(1);

  const rangeStart: number = Math.max(2, currentPage - sideCount);
  const rangeEnd: number = Math.min(totalPages - 1, currentPage + sideCount);

  if (rangeStart > 2) {
    pages.push(-1); // ellipsis
  }

  for (let i: number = rangeStart; i <= rangeEnd; i++) {
    pages.push(i);
  }

  if (rangeEnd < totalPages - 1) {
    pages.push(-1); // ellipsis
  }

  // Always show last page
  if (pages[pages.length - 1] !== totalPages) {
    pages.push(totalPages);
  }

  return pages;
}

const Pagination: React.FC<IPaginationProps> = (props) => {
  const { currentPage, totalCount, pageSize, showPaging, pageRange, onPageChange } = props;
  const totalPages: number = Math.max(1, Math.ceil(totalCount / pageSize));

  if (!showPaging || totalPages <= 1) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null;
  }

  const pageNumbers: number[] = buildPageNumbers(currentPage, totalPages, pageRange);

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
