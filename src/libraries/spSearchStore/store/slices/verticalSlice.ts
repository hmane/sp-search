import { StateCreator } from 'zustand';
import { ISearchStore, IVerticalSlice } from '@interfaces/index';

export const createVerticalSlice: StateCreator<ISearchStore, [], [], IVerticalSlice> = (set) => ({
  currentVerticalKey: 'all',
  verticals: [],
  verticalCounts: {},

  setVertical: (key: string): void => {
    set({ currentVerticalKey: key });
  },

  setVerticalCounts: (counts: Record<string, number>): void => {
    set({ verticalCounts: counts });
  },
});
