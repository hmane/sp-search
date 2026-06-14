interface IWebpackRule {
  test?: RegExp;
  use?: Array<string | ILoaderConfig>;
  include?: string[];
  exclude?: unknown;
}

interface ILoaderConfig {
  loader?: string;
  options?: {
    modules?: {
      auto?: boolean;
      localIdentName?: string;
      namedExport?: boolean;
      exportLocalsConvention?: string;
    };
  };
}

interface IWebpackConfig {
  mode: string;
  resolve: { alias: Record<string, string>; modules: string[] };
  module: { rules: IWebpackRule[] };
  plugins: unknown[];
}

declare const require: (path: string) => (config: IWebpackConfig) => IWebpackConfig;
const customizeWebpack = require('../../config/spfx-customize-webpack');

function createWebpackConfig(): IWebpackConfig {
  return {
    mode: 'development',
    resolve: { alias: {}, modules: [] },
    module: {
      rules: [
        {
          test: /\.module\.(?:css|scss|scss\.css)$/i,
          use: [],
        },
        {
          test: /\.css$/i,
          use: [],
        },
      ],
    },
    plugins: [],
  };
}

describe('spfx-customize-webpack PnP CSS modules', () => {
  it('forces PnP .module.scss.css files to export CSS module locals', () => {
    const config = customizeWebpack(createWebpackConfig()) as IWebpackConfig;
    const pnpRule = config.module.rules.find(function (rule: IWebpackRule): boolean {
      return Array.isArray(rule.include)
        && String(rule.include[0]).indexOf('node_modules/@pnp') >= 0
        && rule.test instanceof RegExp
        && rule.test.test('PeoplePickerComponent.module.scss.css');
    });

    expect(pnpRule).toBeDefined();

    const cssLoader = (pnpRule as IWebpackRule).use?.find(function (loaderConfig: string | ILoaderConfig): boolean {
      return typeof loaderConfig !== 'string'
        && typeof loaderConfig.loader === 'string'
        && loaderConfig.loader.indexOf('css-loader') >= 0;
    }) as ILoaderConfig | undefined;

    expect(cssLoader).toBeDefined();

    expect(cssLoader?.options?.modules).toMatchObject({
      localIdentName: '[local]',
      namedExport: false,
      exportLocalsConvention: 'as-is',
    });
    expect(cssLoader?.options?.modules?.auto).toBeUndefined();
  });
});
