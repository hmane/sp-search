interface IFs {
  readFileSync(file: string, encoding: 'utf8'): string;
}

interface IPath {
  join(...parts: string[]): string;
}

declare const require: (name: string) => unknown;
declare const process: { cwd(): string };

const fs = require('fs') as IFs;
const path = require('path') as IPath;

function extractFieldNames(xml: string): string[] {
  const names: string[] = [];
  const fieldRegex = /<Field\b[^>]*\bName="([^"]+)"/g;
  let match = fieldRegex.exec(xml);

  while (match) {
    names.push(match[1]);
    match = fieldRegex.exec(xml);
  }

  return names;
}

function readProvisioningXml(relativePath: string): string {
  return fs.readFileSync(path.join(process.cwd(), 'provisioning', relativePath), 'utf8');
}

function readRepoFile(relativePath: string): string {
  return fs.readFileSync(path.join(process.cwd(), relativePath), 'utf8');
}

function extractFieldIds(xml: string): string[] {
  const ids: string[] = [];
  const fieldRegex = /<Field\b[^>]*\bID="\{([^"}]+)\}"/g;
  let match = fieldRegex.exec(xml);

  while (match) {
    ids.push(match[1]);
    match = fieldRegex.exec(xml);
  }

  return ids;
}

const savedQueryFields = [
  'QueryText',
  'SearchState',
  'SearchUrl',
  'EntryType',
  'Category',
  'SharedWith',
  'ResultCount',
  'LastUsed',
  'ExpiresAt'
];

const historyFields = [
  'QueryText',
  'QueryHash',
  'Vertical',
  'SearchPageUrl',
  'SearchState',
  'UseCount',
  'ResultCount',
  'IsZeroResult',
  'ClickedItems',
  'SearchTimestamp'
];

const collectionFields = [
  'ItemUrl',
  'ItemTitle',
  'ItemMetadata',
  'CollectionName',
  'Tags',
  'SharedWith',
  'SortOrder'
];

const telemetryConfigFields = [
  'IsEnabled',
  'DestinationEndpoint',
  'BatchIntervalSeconds',
  'BatchSizeMax',
  'PrivacyAcknowledgedBy',
  'PrivacyAcknowledgedAt'
];

const telemetryOptInFields = [
  'ConsentTimestamp',
  'ConsentVersion',
  'AnonHash'
];

const requiredIndexesByList: Record<string, string[]> = {
  SearchSavedQueries: ['Author', 'Title', 'EntryType', 'Category', 'LastUsed', 'ExpiresAt'],
  SearchHistory: ['Author', 'SearchTimestamp', 'QueryHash', 'Vertical'],
  SearchCollections: ['Author', 'Title', 'CollectionName']
};

describe('core hidden-list provisioning schemas', () => {
  it('SearchSavedQueries contains every field the runtime reads or writes', () => {
    const xml = readProvisioningXml(path.join('Lists', 'SearchSavedQueries.xml'));
    const fieldNames = extractFieldNames(xml);

    for (const fieldName of savedQueryFields) {
      expect(fieldNames).toContain(fieldName);
    }
  });

  it('SearchHistory contains every field the runtime reads or writes', () => {
    const xml = readProvisioningXml(path.join('Lists', 'SearchHistory.xml'));
    const fieldNames = extractFieldNames(xml);

    for (const fieldName of historyFields) {
      expect(fieldNames).toContain(fieldName);
    }
  });

  it('SearchCollections contains every field the runtime reads or writes', () => {
    const xml = readProvisioningXml(path.join('Lists', 'SearchCollections.xml'));
    const fieldNames = extractFieldNames(xml);

    for (const fieldName of collectionFields) {
      expect(fieldNames).toContain(fieldName);
    }
  });

  it('SearchTelemetryConfig contains the optional telemetry config fields', () => {
    const xml = readProvisioningXml(path.join('Lists', 'SearchTelemetryConfig.xml'));
    const fieldNames = extractFieldNames(xml);

    for (const fieldName of telemetryConfigFields) {
      expect(fieldNames).toContain(fieldName);
    }

    expect(xml).toContain('<pnp:DataRows KeyColumn="Title" UpdateBehavior="Overwrite">');
    expect(xml).toContain('<pnp:DataValue FieldName="IsEnabled">false</pnp:DataValue>');
    expect(xml).toContain('<pnp:DataValue FieldName="BatchIntervalSeconds">300</pnp:DataValue>');
    expect(xml).toContain('<pnp:DataValue FieldName="BatchSizeMax">50</pnp:DataValue>');
  });

  it('SearchTelemetryOptIn contains the optional telemetry consent fields', () => {
    const xml = readProvisioningXml(path.join('Lists', 'SearchTelemetryOptIn.xml'));
    const fieldNames = extractFieldNames(xml);

    for (const fieldName of telemetryOptInFields) {
      expect(fieldNames).toContain(fieldName);
    }
  });

  it('site template includes core and optional hidden-list schemas', () => {
    const xml = readProvisioningXml('SiteTemplate.xml');

    expect(xml).toContain('Lists/SearchSavedQueries.xml');
    expect(xml).toContain('Lists/SearchHistory.xml');
    expect(xml).toContain('Lists/SearchCollections.xml');
    expect(xml).toContain('Lists/SearchTelemetryConfig.xml');
    expect(xml).toContain('Lists/SearchTelemetryOptIn.xml');
  });

  it('uses valid unique field IDs across provisioned custom fields', () => {
    const xmlFiles = [
      readProvisioningXml(path.join('Lists', 'SearchSavedQueries.xml')),
      readProvisioningXml(path.join('Lists', 'SearchHistory.xml')),
      readProvisioningXml(path.join('Lists', 'SearchCollections.xml')),
      readProvisioningXml(path.join('Lists', 'SearchTelemetryConfig.xml')),
      readProvisioningXml(path.join('Lists', 'SearchTelemetryOptIn.xml'))
    ];
    const ids = xmlFiles.reduce<string[]>((acc, xml) => acc.concat(extractFieldIds(xml)), []);
    const uniqueIds = new Set(ids.map((id) => id.toLowerCase()));
    const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

    expect(uniqueIds.size).toBe(ids.length);
    for (const id of ids) {
      expect(id).toMatch(guidPattern);
    }
  });
});

describe('PowerShell provisioning paths', () => {
  const scripts = [
    'scripts/Provision-SPSearchLists.ps1',
    'scripts/Setup-SPSearchSite.ps1',
    'scripts/Provision-AccountDocumentsEnvironment.ps1'
  ];

  it('all list provisioning scripts include every support list and field', () => {
    for (const scriptPath of scripts) {
      const script = readRepoFile(scriptPath);

      expect(script).toContain('SearchSavedQueries');
      expect(script).toContain('SearchHistory');
      expect(script).toContain('SearchCollections');
      expect(script).toContain('SearchTelemetryConfig');
      expect(script).toContain('SearchTelemetryOptIn');

      for (const fieldName of savedQueryFields) {
        expect(script).toContain(fieldName);
      }
      for (const fieldName of historyFields) {
        expect(script).toContain(fieldName);
      }
      for (const fieldName of collectionFields) {
        expect(script).toContain(fieldName);
      }
      for (const fieldName of telemetryConfigFields) {
        expect(script).toContain(fieldName);
      }
      for (const fieldName of telemetryOptInFields) {
        expect(script).toContain(fieldName);
      }
    }
  });

  it('reset script removes all support lists created by the account-documents provisioner', () => {
    const script = readRepoFile('scripts/Reset-AccountDocumentsEnvironment.ps1');

    expect(script).toContain('SearchSavedQueries');
    expect(script).toContain('SearchHistory');
    expect(script).toContain('SearchCollections');
    expect(script).toContain('SearchTelemetryConfig');
    expect(script).toContain('SearchTelemetryOptIn');
  });

  it('runtime-critical indexes are declared in provisioning scripts and deploy repair', () => {
    const scriptsToCheck = scripts.concat(['scripts/Deploy-SPSearchSolution.ps1']);

    for (const scriptPath of scriptsToCheck) {
      const script = readRepoFile(scriptPath);
      for (const listName of Object.keys(requiredIndexesByList)) {
        for (const fieldName of requiredIndexesByList[listName]) {
          expect(script).toContain(listName);
          expect(script).toContain(fieldName);
        }
      }
    }
  });

  it('configuration export/import page lookup accepts URLs, folders, and recursive Site Pages matches', () => {
    const scriptsToCheck = [
      'scripts/Export-SPSearchPageConfig.ps1',
      'scripts/Import-SPSearchPageConfig.ps1'
    ];

    for (const scriptPath of scriptsToCheck) {
      const script = readRepoFile(scriptPath);

      expect(script).toContain('function Get-PageServerRelativeCandidates');
      expect(script).toContain('-split "[?#]"');
      expect(script).toContain("<View Scope='RecursiveAll'>");
      expect(script).toContain("<FieldRef Name='FileRef' />");
      expect(script).toContain("<FieldRef Name='FileLeafRef' />");
      expect(script).toContain("<FieldRef Name='Title' />");
      expect(script).toContain('full page URL without needing to remove query-string parameters');
    }
  });

  it('page resolver fails closed on off-target URLs and ambiguous matches', () => {
    const scriptsToCheck = [
      'scripts/Export-SPSearchPageConfig.ps1',
      'scripts/Import-SPSearchPageConfig.ps1'
    ];

    for (const scriptPath of scriptsToCheck) {
      const script = readRepoFile(scriptPath);

      // #3 — only genuine http(s) URLs are treated as absolute.
      expect(script).toContain('$absoluteUri.Scheme -eq "http"');
      // #1 — a URL on a different host or site collection throws instead of
      // silently falling through to a same-named page on the connected site.
      expect(script).toContain('points to host');
      expect(script).toContain('outside the connected site');
      // #1 — the bare-leaf SitePages fallback is gated on the absence of an
      // explicit caller-supplied path.
      expect(script).toContain('if (-not $hasExplicitPath)');
      // #2 — leaf-name / Title fallbacks fail closed when more than one page
      // matches rather than silently taking the first row.
      expect(script).toContain('to disambiguate');
      expect(script).toContain('to select one');
    }
  });
});

describe('client-side page provisioning schema', () => {
  it('uses the current Admin Manager dashboard properties', () => {
    const xml = readProvisioningXml('ClientSidePages.xml');

    expect(xml).toContain('&quot;defaultTab&quot;:&quot;dashboard&quot;');
    expect(xml).toContain('&quot;enableDashboard&quot;:true');
    expect(xml).not.toContain('&quot;defaultTab&quot;:&quot;coverage&quot;');
  });
});
