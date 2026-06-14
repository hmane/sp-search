import type { ISearchHistoryEntry } from '@interfaces/index';

export interface ISearchHistoryDateGroup {
  key: string;
  label: string;
  count: number;
  entries: ISearchHistoryEntry[];
}

function padDatePart(value: number): string {
  return value < 10 ? '0' + String(value) : String(value);
}

function getLocalDateKey(date: Date): string {
  return [
    String(date.getFullYear()),
    padDatePart(date.getMonth() + 1),
    padDatePart(date.getDate()),
  ].join('-');
}

function getGroupLabel(date: Date, now: Date): string {
  const todayKey = getLocalDateKey(now);
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const dateKey = getLocalDateKey(date);

  if (dateKey === todayKey) {
    return 'Today';
  }
  if (dateKey === getLocalDateKey(yesterday)) {
    return 'Yesterday';
  }

  return date.toLocaleDateString(undefined, {
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  });
}

export function formatHistoryTime(date: Date): string {
  return date.toLocaleTimeString(undefined, {
    hour: 'numeric',
    minute: '2-digit',
  });
}

export function groupSearchHistoryByDate(
  history: ISearchHistoryEntry[],
  now: Date = new Date()
): ISearchHistoryDateGroup[] {
  const groups: ISearchHistoryDateGroup[] = [];
  const groupsByKey: Record<string, ISearchHistoryDateGroup> = {};

  for (let i = 0; i < history.length; i++) {
    const entry = history[i];
    const key = getLocalDateKey(entry.searchTimestamp);
    let group = groupsByKey[key];

    if (!group) {
      group = {
        key,
        label: getGroupLabel(entry.searchTimestamp, now),
        count: 0,
        entries: [],
      };
      groupsByKey[key] = group;
      groups.push(group);
    }

    group.entries.push(entry);
    group.count++;
  }

  return groups;
}
