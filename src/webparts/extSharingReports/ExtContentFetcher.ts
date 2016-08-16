import {
  ISecurableObjectStore,
  IExtContentFetcherProps,
  IGetExtContentFuncResponse,
  SPScope,
  Mode,
  ISecurableObject
} from './ExtSharingReportsInterfaces';

export default class ExtContentFetcher implements ISecurableObjectStore {

  public props: IExtContentFetcherProps;

  public constructor (props: IExtContentFetcherProps) {
    this.props = props;
  }

  public getAllExtDocuments(): Promise<IGetExtContentFuncResponse> {

    // TODO : do some clever caching

    const rowLimit: number = 500; // TODO : we need to get many pages with this etc..
    const baseUri: string = this.props.host.pageContext.webAbsoluteUrl + '/_api/search/query';

    let scopeFql: string = '';
    if (this.props.scope === SPScope.SiteCollection) {
      scopeFql = ' Path:' + this.props.host.pageContext.webAbsoluteUrl; // TODO: get site collection url
    }
    else if (this.props.scope === SPScope.Site) {
      scopeFql = ' Path:' + this.props.host.pageContext.webAbsoluteUrl;
    }
    else if (this.props.scope === SPScope.Tenant) {
      // do nothing
    }
    else {
      throw new Error("Unsupported scope: " + this.props.scope);
    }

    let modeFql: string = '';
    if (this.props.mode === Mode.AllExtSharedDocuments || this.props.mode === Mode.MyExtSharedDocuments) {
      modeFql = ' IsDocument=1'; // TODO: Do something better than this
    }
    else if (this.props.mode === Mode.AllExtSharedContainers || this.props.mode === Mode.MyExtSharedContainers) {
      modeFql = ' IsDocument=0'; // TODO: Do something better than this
    }
    else {
      throw new Error("Unsupported mode: " + this.props.mode);
    }

    const queryText: string = "querytext='" + scopeFql + modeFql + "'";
    const selectProps: string = "selectproperties='Title,ServerRedirectedURL,FileExtension'";
    const finalUri: string = baseUri + '?' + queryText + '&' + selectProps + '&rowlimit=' + rowLimit;

    return this.props.host.httpClient.get(finalUri)
      .then((r1: Response) => {
        return r1.json().then((r) => {
          return this._transformSearchResults(r, this.props.noResultsString);
        });
      });
  }

  private _transformSearchResults(response: any, noResultsString: string): IGetExtContentFuncResponse {
    // Simplify the data strucutre
    let shouldShowMessage: boolean = false;
    let message: string = "";

    const searchRowsSimplified: ISecurableObject[] = [];

    if (response.PrimaryQueryResult) {
      try {
        const searchRows: any[] = response.PrimaryQueryResult.RelevantResults.Table.Rows;
        searchRows.forEach((d: any) => {
          const doc: any = {};
          d.Cells.forEach((c: any) => {
            doc[c.Key] = c.Value;
          });
          // TODO : convert to ISecurableObject
          searchRowsSimplified.push(doc);
        });
      } catch (e) {
        // TODO: log something?
        shouldShowMessage = true;
        message = "Sorry there was an error";
      }
    }

    if (searchRowsSimplified.length < 1) {
      shouldShowMessage = true;
      message = noResultsString;
    }

    return {
      extContent: searchRowsSimplified,
      shouldShowMessage: shouldShowMessage,
      message: message
    };
  };
}
