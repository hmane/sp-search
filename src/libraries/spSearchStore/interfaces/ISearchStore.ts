import {
  IQuerySlice,
  IFilterSlice,
  IResultSlice,
  IVerticalSlice,
  IUISlice,
  IUserSlice
} from './IStoreSlices';
import { IRegistryContainer } from './IRegistry';

/**
 * Root Zustand store type — flat intersection of all slices
 * plus registries and lifecycle methods.
 *
 * Uses the standard Zustand "slice pattern":
 * all slice properties live at the root level.
 */
export type ISearchStore = IQuerySlice &
  IFilterSlice &
  IResultSlice &
  IVerticalSlice &
  IUISlice &
  IUserSlice & {
    registries: IRegistryContainer;
    /** Reset all slices to default state */
    reset: () => void;
    /** Clean up subscriptions, abort controllers, URL listeners */
    dispose: () => void;
  };
