import * as React from 'react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import styles from './SpSearchFilters.module.scss';
import type { IActiveFilter, IFilterConfig, IRefinerValue } from '@interfaces/index';

export interface IPeoplePickerFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
}

function extractClaimFromItem(item: any): string | undefined {
  const loginName = item && (item.loginName || item.id);
  if (typeof loginName === 'string' && loginName.length > 0) {
    return loginName;
  }
  const email = item && (item.secondaryText || item.text || item.email);
  if (typeof email === 'string' && email.indexOf('@') >= 0) {
    return 'i:0#.f|membership|' + email;
  }
  return undefined;
}

function extractEmailFromClaim(claim: string): string {
  const parts = claim.split('|');
  const last = parts[parts.length - 1];
  return last || claim;
}

const PeoplePickerFilter: React.FC<IPeoplePickerFilterProps> = (props: IPeoplePickerFilterProps): React.ReactElement => {
  const { filterName, config, activeFilters, onToggleRefiner } = props;

  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';
  const selectionLimit = config && config.maxValues > 0 ? config.maxValues : 10;

  const selectedClaims = React.useMemo((): string[] => {
    const selected: string[] = [];
    for (let i = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        selected.push(activeFilters[i].value);
      }
    }
    return selected;
  }, [activeFilters, filterName]);

  const defaultSelectedUsers = React.useMemo((): string[] => {
    return selectedClaims.map(extractEmailFromClaim);
  }, [selectedClaims]);

  function handleChange(items: any[]): void {
    const nextClaims: string[] = [];
    for (let i = 0; i < items.length; i++) {
      const claim = extractClaimFromItem(items[i]);
      if (claim) {
        nextClaims.push(claim);
      }
    }

    for (let i = 0; i < nextClaims.length; i++) {
      if (selectedClaims.indexOf(nextClaims[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: nextClaims[i],
          operator,
        });
      }
    }

    for (let i = 0; i < selectedClaims.length; i++) {
      if (nextClaims.indexOf(selectedClaims[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: selectedClaims[i],
          operator,
        });
      }
    }
  }

  const pickerKey = selectedClaims.join('|');

  return (
    <div className={styles.peopleFilterContainer}>
      <PeoplePicker
        key={pickerKey}
        context={SPContext.peoplepickerContext}
        personSelectionLimit={selectionLimit}
        showtooltip={true}
        required={false}
        disabled={false}
        ensureUser={true}
        showHiddenInUI={false}
        principalTypes={[PrincipalType.User]}
        defaultSelectedUsers={defaultSelectedUsers}
        onChange={handleChange}
        resolveDelay={200}
        suggestionsLimit={10}
        placeholder="Select people"
        webAbsoluteUrl={SPContext.webAbsoluteUrl}
      />
    </div>
  );
};

export default PeoplePickerFilter;
