import * as React from 'react';
import { useStore } from 'zustand';
import { OverflowSet } from '@fluentui/react/lib/OverflowSet';
import { IconButton } from '@fluentui/react/lib/Button';
import { type IOverflowSetItemProps } from '@fluentui/react/lib/OverflowSet';
import { type IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';

import styles from './SpSearchVerticals.module.scss';
import type { ISpSearchVerticalsProps } from './ISpSearchVerticalsProps';
import type { IVerticalDefinition } from '@interfaces/index';
import VerticalTab from './VerticalTab';

/**
 * Map of tabStyle prop to the corresponding CSS module class name.
 */
const STYLE_CLASS_MAP: Record<string, string> = {
  tabs: styles.styleTabs,
  pills: styles.stylePills,
  underline: styles.styleUnderline
};

interface IOverflowVerticalItem extends IOverflowSetItemProps {
  vertical: IVerticalDefinition;
  isActive: boolean;
  count: number | undefined;
  isDimmed: boolean;
}

/**
 * SpSearchVerticals -- a functional React component that renders search vertical tabs
 * with badge counts, overflow handling, and three visual styles (tabs, pills, underline).
 */
const SpSearchVerticalsInner: React.FC<ISpSearchVerticalsProps> = (props: ISpSearchVerticalsProps): React.ReactElement => {
  const { store, showCounts, hideEmptyVerticals, tabStyle } = props;

  // Subscribe to specific store slices
  const currentVerticalKey: string = useStore(store, function (s) { return s.currentVerticalKey; });
  const verticals: IVerticalDefinition[] = useStore(store, function (s) { return s.verticals; });
  const verticalCounts: Record<string, number> = useStore(store, function (s) { return s.verticalCounts; });
  const setVertical: (key: string) => void = useStore(store, function (s) { return s.setVertical; });

  // Container ref for overflow measurement
  // Using undefined cast to satisfy @rushstack/no-new-null while matching RefObject<HTMLDivElement>
  const containerRef = React.useRef<HTMLDivElement>(undefined as unknown as HTMLDivElement);
  const [maxVisibleCount, setMaxVisibleCount] = React.useState<number>(verticals.length);

  // Measure available width and compute how many tabs fit
  React.useEffect(function (): (() => void) | undefined {
    if (!containerRef.current || verticals.length === 0) {
      return undefined;
    }

    function measure(): void {
      if (!containerRef.current) {
        return;
      }
      const containerWidth: number = containerRef.current.offsetWidth;
      // Reserve 60px for the "More" overflow button when needed
      const moreButtonWidth: number = 60;
      const children: HTMLCollection = containerRef.current.children;
      let usedWidth: number = 0;
      let fitCount: number = 0;

      for (let i: number = 0; i < children.length; i++) {
        const child: HTMLElement = children[i] as HTMLElement;
        // Skip the overflow button if present
        if (child.getAttribute('data-overflow-button') === 'true') {
          continue;
        }
        const childWidth: number = child.offsetWidth;
        if (usedWidth + childWidth + moreButtonWidth <= containerWidth) {
          usedWidth += childWidth;
          fitCount++;
        } else {
          break;
        }
      }

      // If all tabs fit without the overflow button, show them all
      if (fitCount >= verticals.length) {
        setMaxVisibleCount(verticals.length);
      } else {
        setMaxVisibleCount(Math.max(1, fitCount));
      }
    }

    // Initial measurement after render
    measure();

    // Observe container resize
    let observer: ResizeObserver | undefined;
    if (typeof ResizeObserver !== 'undefined') {
      observer = new ResizeObserver(function () {
        measure();
      });
      observer.observe(containerRef.current);
    }

    return function (): void {
      if (observer) {
        observer.disconnect();
      }
    };
  }, [verticals.length]);

  const handleTabClick = React.useCallback(function (key: string): void {
    setVertical(key);
  }, [setVertical]);

  // Determine which verticals to show or hide
  const visibleVerticals: IVerticalDefinition[] = [];
  for (let i: number = 0; i < verticals.length; i++) {
    const v: IVerticalDefinition = verticals[i];
    const count: number | undefined = verticalCounts[v.key];
    const isEmpty: boolean = count !== undefined && count === 0;
    if (hideEmptyVerticals && isEmpty && v.key !== currentVerticalKey) {
      continue;
    }
    visibleVerticals.push(v);
  }

  // Split into primary (visible) and overflow
  const primaryItems: IVerticalDefinition[] = visibleVerticals.slice(0, maxVisibleCount);
  const overflowItems: IVerticalDefinition[] = visibleVerticals.slice(maxVisibleCount);

  // Style class for the chosen tab style
  const styleClass: string = STYLE_CLASS_MAP[tabStyle] || STYLE_CLASS_MAP.tabs;

  // Build overflow set items
  const primaryOverflowItems: IOverflowVerticalItem[] = primaryItems.map(function (v: IVerticalDefinition): IOverflowVerticalItem {
    const count: number | undefined = verticalCounts[v.key];
    const isEmpty: boolean = count !== undefined && count === 0;
    return {
      key: v.key,
      vertical: v,
      isActive: v.key === currentVerticalKey,
      count: count,
      isDimmed: isEmpty
    };
  });

  const overflowSetItems: IOverflowVerticalItem[] = overflowItems.map(function (v: IVerticalDefinition): IOverflowVerticalItem {
    const count: number | undefined = verticalCounts[v.key];
    const isEmpty: boolean = count !== undefined && count === 0;
    return {
      key: v.key,
      vertical: v,
      isActive: v.key === currentVerticalKey,
      count: count,
      isDimmed: isEmpty
    };
  });

  const onRenderItem = React.useCallback(function (item: IOverflowSetItemProps): React.ReactElement {
    const typedItem: IOverflowVerticalItem = item as IOverflowVerticalItem;
    return (
      <VerticalTab
        key={typedItem.key}
        verticalKey={typedItem.vertical.key}
        label={typedItem.vertical.label}
        iconName={typedItem.vertical.iconName}
        count={typedItem.count}
        isActive={typedItem.isActive}
        showCounts={showCounts}
        isDimmed={typedItem.isDimmed}
        onClick={handleTabClick}
      />
    );
  }, [showCounts, handleTabClick]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const onRenderOverflowButton = React.useCallback(function (items: any[] | undefined): JSX.Element {
    if (!items || items.length === 0) {
      return React.createElement(React.Fragment);
    }

    const menuItems: IContextualMenuItem[] = items.map(function (item: IOverflowSetItemProps): IContextualMenuItem {
      const typedItem: IOverflowVerticalItem = item as IOverflowVerticalItem;
      const countText: string = showCounts && typedItem.count !== undefined ? ' (' + String(typedItem.count) + ')' : '';
      return {
        key: typedItem.key,
        text: typedItem.vertical.label + countText,
        iconProps: typedItem.vertical.iconName ? { iconName: typedItem.vertical.iconName } : undefined,
        checked: typedItem.isActive,
        disabled: typedItem.isDimmed,
        onClick: function (): void {
          handleTabClick(typedItem.vertical.key);
        }
      };
    });

    return (
      <IconButton
        data-overflow-button="true"
        className={styles.moreButton}
        menuIconProps={{ iconName: 'More' }}
        title="More"
        ariaLabel="More verticals"
        menuProps={{
          items: menuItems
        }}
      />
    );
  }, [showCounts, handleTabClick]);

  if (verticals.length === 0) {
    return <div className={styles.spSearchVerticals} />;
  }

  const hasOverflow: boolean = overflowSetItems.length > 0;

  return (
    <div className={styles.spSearchVerticals}>
      <div
        className={styles.tabContainer + ' ' + styleClass}
        role="tablist"
        aria-label="Search verticals"
        ref={containerRef}
      >
        {!hasOverflow && primaryItems.map(function (v: IVerticalDefinition): React.ReactElement {
          const count: number | undefined = verticalCounts[v.key];
          const isEmpty: boolean = count !== undefined && count === 0;
          return (
            <VerticalTab
              key={v.key}
              verticalKey={v.key}
              label={v.label}
              iconName={v.iconName}
              count={count}
              isActive={v.key === currentVerticalKey}
              showCounts={showCounts}
              isDimmed={isEmpty}
              onClick={handleTabClick}
            />
          );
        })}
        {hasOverflow && (
          <OverflowSet
            role="tablist"
            items={primaryOverflowItems}
            overflowItems={overflowSetItems}
            onRenderItem={onRenderItem}
            onRenderOverflowButton={onRenderOverflowButton}
          />
        )}
      </div>
    </div>
  );
};

/**
 * Wrapped in ErrorBoundary for production safety.
 */
const SpSearchVerticals: React.FC<ISpSearchVerticalsProps> = (props: ISpSearchVerticalsProps): React.ReactElement => {
  return (
    <ErrorBoundary
      enableRetry={true}
      maxRetries={3}
      showDetailsButton={true}
    >
      <SpSearchVerticalsInner {...props} />
    </ErrorBoundary>
  );
};

export default SpSearchVerticals;
