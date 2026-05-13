import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import {
  setWebPartContextId,
  unregisterWebPartContextId,
  subscribeContextIdChanges,
  getPeerContextIds,
} from '@store/utils/contextIdRegistry';

/**
 * T3.D2 — edit-mode banner that warns admins when this web part's
 * `searchContextId` doesn't match peer web parts on the same page.
 *
 * The banner:
 *   - registers this web part's contextId in the window-backed registry
 *     on mount + on every contextId change
 *   - unregisters on unmount
 *   - subscribes to peer changes and re-renders when any peer comes,
 *     goes, or changes their id
 *   - renders nothing in view mode (Required scenario S2: edit-only)
 *   - renders nothing in edit mode when every peer matches this id
 *   - renders a Fluent `MessageBar` (severityError) in edit mode when at
 *     least one peer has a different id
 *
 * The wrapper variant `<SearchContextIdBannerWrapper>` puts the banner
 * above the supplied children, so each web part's `render()` can wrap
 * its existing tree in one line.
 */

export interface ISearchContextIdMismatchBannerProps {
  /** Stable identifier for this web part instance (typically `this.instanceId`). */
  webPartId: string;
  /** This web part's current `searchContextId`. */
  contextId: string;
  /** Friendly name for this web part — appears in the banner copy. */
  webPartLabel: string;
  /** When false, the banner is unmounted (view mode). */
  isEditMode: boolean;
}

export const SearchContextIdMismatchBanner: React.FC<ISearchContextIdMismatchBannerProps> = (props) => {
  const { webPartId, contextId, webPartLabel, isEditMode } = props;
  const [peers, setPeers] = React.useState<string[]>(() => getPeerContextIds(webPartId, contextId));

  React.useEffect(function registerAndSubscribe(): () => void {
    setWebPartContextId(webPartId, contextId);
    setPeers(getPeerContextIds(webPartId, contextId));
    const unsubscribe = subscribeContextIdChanges(function (): void {
      setPeers(getPeerContextIds(webPartId, contextId));
    });
    return function cleanup(): void {
      unsubscribe();
      unregisterWebPartContextId(webPartId);
    };
  }, [webPartId, contextId]);

  if (!isEditMode || peers.length === 0) {
    return null;
  }

  const peerLabel = peers.length === 1
    ? `another web part using "${peers[0]}"`
    : `other web parts using ${peers.map((p) => '"' + p + '"').join(', ')}`;

  return (
    <MessageBar
      messageBarType={MessageBarType.severeWarning}
      isMultiline={true}
      styles={{ root: { marginBottom: 8 } }}
    >
      <strong>{webPartLabel}</strong> is configured with <code>{contextId}</code> but the page also contains
      {' '}{peerLabel}. Web parts only exchange queries / filters / verticals when their Search context ID
      matches. Either give every web part the same ID to wire them together, or leave them distinct on
      purpose (multi-context page). Edit each web part&apos;s pane → Search context group → Search context ID.
    </MessageBar>
  );
};

export interface ISearchContextIdBannerWrapperProps extends ISearchContextIdMismatchBannerProps {
  children?: React.ReactNode;
}

/** Convenience wrapper: renders the banner above the children. */
export const SearchContextIdBannerWrapper: React.FC<ISearchContextIdBannerWrapperProps> = (props) => {
  const { children, ...bannerProps } = props;
  return (
    <>
      <SearchContextIdMismatchBanner {...bannerProps} />
      {children}
    </>
  );
};
