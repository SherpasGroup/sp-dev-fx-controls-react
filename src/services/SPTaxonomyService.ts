import { BaseComponentContext } from '@microsoft/sp-component-base';
// import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

import { sp } from '@pnp/sp';
import { ITermInfo, ITermSetInfo } from '@pnp/sp/taxonomy';

export interface ODataCollection<T> {
  '@odata.context': string;
  '@odata.nextLink'?:string;
  value: T[];
}

const EmptyODataCollection = { '@odata.context': '', value: [] };

export class SPTaxonomyService {

  /**
   * Service constructor
   */
  constructor(private context: BaseComponentContext) {

  }

  public async getTerms(siteUrl: string, termGroupId: string, termSetId: string, hideDeprecatedTags?: boolean, pageSize: number = 50): Promise<ODataCollection<ITermInfo>> {
      if (!termGroupId || !termSetId) return EmptyODataCollection;
      const filterParam: string = hideDeprecatedTags ? '&$filter=isDeprecated eq false' : '';
      const endpoint: string = `${siteUrl}/_api/v2.1/termStore/termGroups/${termGroupId}/termSets/${termSetId}/getLegacyChildren?$top=50${filterParam}`;
      try {
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) {
          return EmptyODataCollection;
        }
        const json: ODataCollection<ITermInfo> = await response.json();
        return json;
      } catch (error) {
        return EmptyODataCollection;
      }
  }

  public async getTermsByIds(termSetId: string, termIds: string[]): Promise<ITermInfo[]> {
    if (!termSetId) return null;
    // const termSetId: string = await this.GetTermSetId(termset);
    const batch = sp.createBatch();
    const results = [];
    termIds.forEach(termId => {
      sp.termStore.sets.getById(termSetId).terms.getById(termId).inBatch(batch)()
        .then(term => results.push(term))
        .catch(reason => console.warn(`Error retreiving Term ID ${termId}: ${reason}`));
    });
    await batch.execute();
    return results;
  }

  // public async getTermSetIdFromTermSetName(termSetName: string): Promise<string | undefined> {
  //   // if (this.isGuid(termsetNameOrId)) return termsetNameOrId;

  //   // const termSetName: string = termsetNameOrId;
  //   const termGroups = await sp.termStore.groups.get();
  //   let candidates: ITermSetInfo[] = [];
  //   const filterNames = `localizedNames/any(n: n/name eq '${termSetName}')`;
  //   const promises = termGroups.map(async (grp) => {
  //     const matchingSets = await sp.termStore.groups.getById(grp.id).sets.filter(filterNames).get();
  //     if (matchingSets?.length) {
  //       candidates = candidates.concat(matchingSets);
  //     }
  //   });
  //   await Promise.all(promises);
  //   return candidates?.[0]?.id;
  // }

  public async getTermSetInfo(termSetId: string): Promise<ITermSetInfo | undefined> {
    if (!termSetId) return undefined;
    const tsInfo = await sp.termStore.sets.getById(termSetId)();
    return tsInfo;
  }
}
