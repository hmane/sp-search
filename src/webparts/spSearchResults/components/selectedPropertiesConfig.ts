export interface ISelectedPropertyItemConfig {
  uniqueId: string;
  property: string;
  alias: string;
}

function isTitlePropertyName(property: string): boolean {
  const normalized = (property || '').trim().toLowerCase();
  return normalized === 'title' || normalized === 'filename';
}

export function getDefaultSelectedPropertyItems(): ISelectedPropertyItemConfig[] {
  return [
    { uniqueId: 'sp-author', property: 'Author', alias: 'Author' },
    { uniqueId: 'sp-lastmodified', property: 'LastModifiedTime', alias: 'Modified' },
    { uniqueId: 'sp-filetype', property: 'FileType', alias: 'Type' },
    { uniqueId: 'sp-size', property: 'Size', alias: 'Size' },
    { uniqueId: 'sp-path', property: 'Path', alias: 'URL' },
    { uniqueId: 'sp-sitename', property: 'SiteName', alias: 'Site' }
  ];
}

export function normalizeSelectedPropertyItems(
  raw: ISelectedPropertyItemConfig[],
  shouldSeedDefaults: boolean
): ISelectedPropertyItemConfig[] {
  const result: ISelectedPropertyItemConfig[] = [];
  const seen = new Set<string>();
  let titleItem: ISelectedPropertyItemConfig | undefined;

  for (let i: number = 0; i < raw.length; i++) {
    const property = String(raw[i].property || '').trim();
    if (!property) {
      continue;
    }
    const lookup = property.toLowerCase();
    if (seen.has(lookup)) {
      continue;
    }
    seen.add(lookup);

    const normalizedItem: ISelectedPropertyItemConfig = {
      uniqueId: raw[i].uniqueId || ('sp-' + String(i)),
      property,
      alias: String(raw[i].alias || '').trim()
    };

    if (isTitlePropertyName(property)) {
      titleItem = {
        uniqueId: normalizedItem.uniqueId,
        property: 'Title',
        alias: normalizedItem.alias || 'Name'
      };
    } else {
      result.push(normalizedItem);
    }
  }

  if (!titleItem) {
    titleItem = { uniqueId: 'sp-title', property: 'Title', alias: 'Name' };
  }

  if (shouldSeedDefaults && result.length === 0) {
    const defaults = getDefaultSelectedPropertyItems();
    for (let i: number = 0; i < defaults.length; i++) {
      result.push(defaults[i]);
      seen.add(defaults[i].property.toLowerCase());
    }
  }

  return [titleItem, ...result];
}
