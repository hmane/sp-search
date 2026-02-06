import * as React from 'react';
import styles from './SpSearchManager.module.scss';

export interface ITagBadgesProps {
  tags: string[];
  maxVisible?: number;
}

/**
 * TagBadges — read-only display of tag pill badges.
 * Shows up to `maxVisible` tags with an overflow "+N" indicator.
 */
const TagBadges: React.FC<ITagBadgesProps> = (props) => {
  const { tags, maxVisible } = props;

  if (!tags || tags.length === 0) {
    return null;
  }

  const limit = maxVisible || 5;
  const visibleTags = tags.length > limit ? tags.slice(0, limit) : tags;
  const overflow = tags.length > limit ? tags.length - limit : 0;

  return (
    <div className={styles.tagContainer}>
      {visibleTags.map(function (tag): React.ReactElement {
        return (
          <span key={tag} className={styles.tagBadge} title={tag}>
            {tag}
          </span>
        );
      })}
      {overflow > 0 && (
        <span className={styles.tagBadge} title={tags.slice(limit).join(', ')}>
          +{String(overflow)}
        </span>
      )}
    </div>
  );
};

export default TagBadges;
