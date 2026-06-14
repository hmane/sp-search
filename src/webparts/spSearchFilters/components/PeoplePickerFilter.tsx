import * as React from 'react';
import { NormalPeoplePicker, ValidationState } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import styles from './SpSearchFilters.module.scss';
import type {
  IActiveFilter,
  IFilterConfig,
  IRefinerValue,
  IReplaceRefinerValuesPayload,
} from '@interfaces/index';

export interface IPeoplePickerFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
  /**
   * Batched callback (Task 1 foundation). Fires once per editor change with
   * the FULL intended selection — see TagBoxFilter / DropdownFilter /
   * TaxonomyTreeFilter for sibling implementations.
   */
  onReplaceRefinerValues?: (payload: IReplaceRefinerValuesPayload) => void;
}

interface IClaimSuggestion {
  loginName: string;
  displayName: string;
  email?: string;
}

/**
 * Pure helper — builds the batched `onReplaceRefinerValues` payload from a
 * Fluent NormalPeoplePicker selection. Exported for unit testing so we don't
 * have to drive NormalPeoplePicker via jsdom (its async resolve + onChange
 * are awkward to reach through @testing-library).
 *
 * `value` is the claim (persona.secondaryText); `displayValue` is the
 * resolved display name (persona.text). Order is preserved.
 */
export function buildPeoplePickerBatchPayload(input: {
  filterName: string;
  personas: IPersonaProps[];
  operator: 'AND' | 'OR';
}): IReplaceRefinerValuesPayload {
  const { filterName, personas, operator } = input;
  const values: IActiveFilter[] = personas.map(function (p: IPersonaProps): IActiveFilter {
    return {
      filterName: filterName,
      value: p.secondaryText || '',
      displayValue: p.text,
      operator: operator,
    };
  });
  return { filterName: filterName, values: values };
}

interface IClientPeoplePickerEntity {
  Key: string;
  DisplayText: string;
  EntityData?: { Email?: string };
}

async function resolvePeople(filter: string): Promise<IClaimSuggestion[]> {
  if (!filter || filter.length < 2) {
    return [];
  }

  const body = {
    queryParams: {
      QueryString: filter,
      MaximumEntitySuggestions: 25,
      PrincipalSource: 15,
      // Bitwise OR: User (1) + SecurityGroup (4). Keeps the picker
      // useful for audience-style refiners while excluding DL noise.
      PrincipalType: 1 | 4,
      AllowEmailAddresses: true,
    },
  };

  const url = SPContext.webAbsoluteUrl
    + '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser';

  const response = await SPContext.http.post<{ value: string }>(url, body);
  if (!response.ok || !response.data || !response.data.value) {
    return [];
  }

  let parsed: IClientPeoplePickerEntity[];
  try {
    parsed = JSON.parse(response.data.value) as IClientPeoplePickerEntity[];
  } catch {
    return [];
  }

  return parsed.map(function (p: IClientPeoplePickerEntity): IClaimSuggestion {
    return {
      loginName: p.Key,
      displayName: p.DisplayText,
      email: p.EntityData && p.EntityData.Email ? p.EntityData.Email : undefined,
    };
  });
}

const PeoplePickerFilter: React.FC<IPeoplePickerFilterProps> = (
  props: IPeoplePickerFilterProps
): React.ReactElement => {
  const {
    filterName,
    config,
    activeFilters,
    onToggleRefiner,
    onReplaceRefinerValues,
  } = props;

  const operator: 'AND' | 'OR' = config && config.operator === 'AND' ? 'AND' : 'OR';

  const initialPersonas: IPersonaProps[] = React.useMemo(function (): IPersonaProps[] {
    const result: IPersonaProps[] = [];
    for (let i = 0; i < activeFilters.length; i++) {
      const f = activeFilters[i];
      if (f.filterName === filterName) {
        result.push({
          text: f.displayValue || f.value,
          secondaryText: f.value,
        });
      }
    }
    return result;
  }, [activeFilters, filterName]);

  const [selectedPersonas, setSelectedPersonas] = React.useState<IPersonaProps[]>(initialPersonas);

  React.useEffect(function (): void {
    setSelectedPersonas(initialPersonas);
  }, [initialPersonas]);

  function handleResolveSuggestions(
    filter: string,
    currentPersonas?: IPersonaProps[]
  ): Promise<IPersonaProps[]> {
    const current: IPersonaProps[] = currentPersonas || [];
    return resolvePeople(filter).then(function (claims: IClaimSuggestion[]): IPersonaProps[] {
      const currentLoginNames: string[] = current.map(function (p: IPersonaProps): string {
        return p.secondaryText || '';
      });
      const filtered: IClaimSuggestion[] = [];
      for (let i = 0; i < claims.length; i++) {
        if (currentLoginNames.indexOf(claims[i].loginName) < 0) {
          filtered.push(claims[i]);
        }
      }
      return filtered.map(function (c: IClaimSuggestion): IPersonaProps {
        return {
          text: c.displayName,
          secondaryText: c.loginName,
          tertiaryText: c.email,
        };
      });
    });
  }

  function handleItemsChange(items?: IPersonaProps[]): void {
    const next: IPersonaProps[] = items || [];
    setSelectedPersonas(next);

    if (onReplaceRefinerValues) {
      onReplaceRefinerValues(buildPeoplePickerBatchPayload({
        filterName: filterName,
        personas: next,
        operator: operator,
      }));
      return;
    }

    // Back-compat fallback — compute delta against activeFilters and loop
    // onToggleRefiner. Parents wiring onReplaceRefinerValues skip this path.
    const previousClaims: string[] = [];
    for (let i = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        previousClaims.push(activeFilters[i].value);
      }
    }
    const nextLogins: string[] = next.map(function (p: IPersonaProps): string {
      return p.secondaryText || '';
    });
    for (let i = 0; i < nextLogins.length; i++) {
      if (previousClaims.indexOf(nextLogins[i]) < 0) {
        onToggleRefiner({
          filterName: filterName,
          value: nextLogins[i],
          displayValue: next[i].text,
          operator: operator,
        });
      }
    }
    for (let i = 0; i < previousClaims.length; i++) {
      if (nextLogins.indexOf(previousClaims[i]) < 0) {
        onToggleRefiner({
          filterName: filterName,
          value: previousClaims[i],
          operator: operator,
        });
      }
    }
  }

  return (
    <div className={styles.peopleFilterContainer}>
      <NormalPeoplePicker
        onResolveSuggestions={handleResolveSuggestions}
        selectedItems={selectedPersonas}
        onChange={handleItemsChange}
        onValidateInput={function (): ValidationState { return ValidationState.invalid; }}
        pickerSuggestionsProps={{
          suggestionsHeaderText: 'Suggested people',
          noResultsFoundText: 'No matches',
        }}
        resolveDelay={250}
      />
    </div>
  );
};

export default PeoplePickerFilter;
