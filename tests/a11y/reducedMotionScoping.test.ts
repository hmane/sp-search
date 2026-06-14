interface IDirent {
  name: string;
  isDirectory(): boolean;
  isFile(): boolean;
}

interface IFs {
  readdirSync(dir: string, options: { withFileTypes: true }): IDirent[];
  readFileSync(file: string, encoding: 'utf8'): string;
}

interface IPath {
  join(...parts: string[]): string;
  resolve(...parts: string[]): string;
  relative(from: string, to: string): string;
}

declare const require: (name: string) => unknown;
declare const process: { cwd(): string };

const fs = require('fs') as IFs;
const path = require('path') as IPath;

function collectScssFiles(dir: string, acc: string[] = []): string[] {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      collectScssFiles(fullPath, acc);
    } else if (entry.isFile() && fullPath.endsWith('.module.scss')) {
      acc.push(fullPath);
    }
  }
  return acc;
}

describe('CSS reduced-motion scoping', () => {
  it('does not use unscoped universal selectors in module reduced-motion blocks', () => {
    const root = path.resolve(process.cwd(), 'src');
    const offenders: string[] = [];
    const files = collectScssFiles(root);
    const blockPattern = /@media\s*\(prefers-reduced-motion:\s*reduce\)\s*\{([\s\S]*?)\n\}/g;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const content = fs.readFileSync(file, 'utf8');
      let match: RegExpExecArray | null;
      while ((match = blockPattern.exec(content)) !== null) {
        if (/^\s*\*,|^\s*\*::before,|^\s*\*::after/m.test(match[1])) {
          offenders.push(path.relative(root, file));
        }
      }
    }

    expect(offenders).toEqual([]);
  });

  it('keeps generic vertical tab state classes scoped to the verticals web part root', () => {
    const root = path.resolve(process.cwd(), 'src');
    const file = path.join(root, 'webparts/spSearchVerticals/components/SpSearchVerticals.module.scss');
    const content = fs.readFileSync(file, 'utf8');
    const genericClassPattern = /^\.((?:active|hidden|tabContainer|verticalTab|tabIcon|tabLabel|countBadge|dimmed|styleTabs|stylePills|styleUnderline|moreButton|overflowWrapper))(?:\s|\.|:|\{)/gm;
    const offenders: string[] = [];
    let match: RegExpExecArray | null;

    while ((match = genericClassPattern.exec(content)) !== null) {
      offenders.push(match[1]);
    }

    expect(offenders).toEqual([]);
  });
});
