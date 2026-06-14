import { getPropertyPaneHelpTopic } from '../../src/propertyPaneControls/propertyPaneGroupHelp';

describe('propertyPaneGroupHelp topics', () => {
  it('returns detailed local content for existing property pane anchors', () => {
    const topic = getPropertyPaneHelpTopic('filters-behavior');

    expect(topic.title).toBe('Apply Mode And Clear All Behaviour');
    expect(topic.bullets.join(' ')).toContain('Instant mode');
    expect(topic.examples && topic.examples.length).toBeGreaterThan(0);
  });

  it('falls back to generic local content for unknown anchors', () => {
    const topic = getPropertyPaneHelpTopic('unknown-anchor');

    expect(topic.title).toBe('SP Search Help');
    expect(topic.bullets.join(' ')).toContain('Search context ID');
  });
});
