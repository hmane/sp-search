import { TokenService, ITokenContext } from '../../src/libraries/spSearchStore/services/TokenService';
import { createMockTokenContext } from '../utils/testHelpers';

describe('TokenService', () => {
  let context: ITokenContext;

  beforeEach(() => {
    context = createMockTokenContext();
  });

  describe('resolveTokens', () => {
    describe('basic tokens', () => {
      it('should replace {searchTerms} with queryText', () => {
        const result = TokenService.resolveTokens('{searchTerms}', context);
        expect(result).toBe('annual report');
      });

      it('should replace {Site.ID} with siteId', () => {
        const result = TokenService.resolveTokens('{Site.ID}', context);
        expect(result).toBe('b5c3e9a1-2d4f-4e6a-8b7c-1d2e3f4a5b6c');
      });

      it('should replace {Site.URL} with siteUrl', () => {
        const result = TokenService.resolveTokens('{Site.URL}', context);
        expect(result).toBe('https://contoso.sharepoint.com/sites/intranet');
      });

      it('should replace {Web.ID} with webId', () => {
        const result = TokenService.resolveTokens('{Web.ID}', context);
        expect(result).toBe('a1b2c3d4-e5f6-7890-abcd-ef1234567890');
      });

      it('should replace {Web.URL} with webUrl', () => {
        const result = TokenService.resolveTokens('{Web.URL}', context);
        expect(result).toBe('https://contoso.sharepoint.com/sites/intranet/subweb');
      });

      it('should replace {Hub} with hubSiteUrl', () => {
        const result = TokenService.resolveTokens('{Hub}', context);
        expect(result).toBe('https://contoso.sharepoint.com/sites/hub');
      });

      it('should replace {User.Name} with userDisplayName', () => {
        const result = TokenService.resolveTokens('{User.Name}', context);
        expect(result).toBe('Jane Smith');
      });

      it('should replace {User.Email} with userEmail', () => {
        const result = TokenService.resolveTokens('{User.Email}', context);
        expect(result).toBe('jane.smith@contoso.com');
      });

      it('should replace {PageContext.listId} with listId', () => {
        const ctxWithList = createMockTokenContext({ listId: 'a1b2c3d4-list-guid' });
        const result = TokenService.resolveTokens('{PageContext.listId}', ctxWithList);
        expect(result).toBe('a1b2c3d4-list-guid');
      });

      it('should replace {PageContext.listId} with empty string when not on a list page', () => {
        const result = TokenService.resolveTokens('{PageContext.listId}', context);
        expect(result).toBe('');
      });
    });

    describe('multiple tokens in one template', () => {
      it('should replace all tokens in a complex template', () => {
        const template = '{searchTerms} Path:{Site.URL} Author:{User.Name}';
        const result = TokenService.resolveTokens(template, context);
        expect(result).toBe(
          'annual report Path:https://contoso.sharepoint.com/sites/intranet Author:Jane Smith'
        );
      });

      it('should replace the same token appearing multiple times', () => {
        const template = '{searchTerms} OR title:{searchTerms}';
        const result = TokenService.resolveTokens(template, context);
        expect(result).toBe('annual report OR title:annual report');
      });
    });

    describe('date tokens', () => {
      it('should replace {Today} with current date in YYYY-MM-DD format', () => {
        const result = TokenService.resolveTokens('{Today}', context);
        const today = new Date();
        const expected = formatDate(today);
        expect(result).toBe(expected);
      });

      it('should replace {Today+5} with date 5 days in the future', () => {
        const result = TokenService.resolveTokens('{Today+5}', context);
        const future = new Date();
        future.setDate(future.getDate() + 5);
        expect(result).toBe(formatDate(future));
      });

      it('should replace {Today-30} with date 30 days in the past', () => {
        const result = TokenService.resolveTokens('{Today-30}', context);
        const past = new Date();
        past.setDate(past.getDate() - 30);
        expect(result).toBe(formatDate(past));
      });

      it('should replace {Today+0} with current date', () => {
        const result = TokenService.resolveTokens('{Today+0}', context);
        const today = new Date();
        expect(result).toBe(formatDate(today));
      });

      it('should handle {Today-365} for one year ago', () => {
        const result = TokenService.resolveTokens('{Today-365}', context);
        const past = new Date();
        past.setDate(past.getDate() - 365);
        expect(result).toBe(formatDate(past));
      });

      it('should handle date tokens in a template', () => {
        const template = 'LastModifiedTime>={Today-7} AND LastModifiedTime<={Today}';
        const result = TokenService.resolveTokens(template, context);

        const today = new Date();
        const weekAgo = new Date();
        weekAgo.setDate(weekAgo.getDate() - 7);

        expect(result).toBe(
          `LastModifiedTime>=${formatDate(weekAgo)} AND LastModifiedTime<=${formatDate(today)}`
        );
      });
    });

    describe('unknown tokens', () => {
      it('should leave unknown tokens unreplaced', () => {
        const result = TokenService.resolveTokens('{UnknownToken}', context);
        expect(result).toBe('{UnknownToken}');
      });

      it('should leave unknown tokens intact among replaced tokens', () => {
        const template = '{searchTerms} {UnknownToken} Path:{Site.URL}';
        const result = TokenService.resolveTokens(template, context);
        expect(result).toBe(
          'annual report {UnknownToken} Path:https://contoso.sharepoint.com/sites/intranet'
        );
      });

      it('should not replace tokens with incorrect casing', () => {
        // Token names are case-sensitive
        const result = TokenService.resolveTokens('{searchterms}', context);
        expect(result).toBe('{searchterms}');
      });
    });

    describe('edge cases', () => {
      it('should return empty string for empty template', () => {
        const result = TokenService.resolveTokens('', context);
        expect(result).toBe('');
      });

      it('should return empty string for null-ish template', () => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const result = TokenService.resolveTokens(undefined as any, context);
        expect(result).toBe('');
      });

      it('should handle template with no tokens', () => {
        const result = TokenService.resolveTokens('plain KQL query', context);
        expect(result).toBe('plain KQL query');
      });

      it('should handle template with only curly braces but no content', () => {
        const result = TokenService.resolveTokens('{}', context);
        // {} matches the regex but the token key is empty — should be unknown
        expect(result).toBe('{}');
      });

      it('should handle context with empty queryText', () => {
        const emptyCtx = createMockTokenContext({ queryText: '' });
        const result = TokenService.resolveTokens('{searchTerms}', emptyCtx);
        expect(result).toBe('');
      });

      it('should handle context with special characters in queryText', () => {
        const specialCtx = createMockTokenContext({ queryText: 'C# "hello world" path:*.ts' });
        const result = TokenService.resolveTokens('{searchTerms}', specialCtx);
        expect(result).toBe('C# "hello world" path:*.ts');
      });
    });
  });
});

/** Helper to format date as YYYY-MM-DD matching TokenService._formatDate */
function formatDate(date: Date): string {
  const year = date.getFullYear();
  const month = ('0' + String(date.getMonth() + 1)).slice(-2);
  const day = ('0' + String(date.getDate())).slice(-2);
  return `${year}-${month}-${day}`;
}
