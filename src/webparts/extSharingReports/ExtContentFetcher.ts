import {
  ISecurableObjectStore,
  IExtContentFetcherProps,
  IGetExtContentFuncResponse,
  SPScope,
  Mode,
  ISecurableObject,
  ControlMode
} from "./ExtSharingReportsInterfaces";

import {
  Logger
} from "./Logger";

export default class ExtContentFetcher implements ISecurableObjectStore {

  public props: IExtContentFetcherProps;
  public timeStamp: number;
  private log: Logger;

  public constructor (props: IExtContentFetcherProps) {
    this.props = props;
    this.log = new Logger("ExtContentFetcher");
    this.timeStamp = -1;
  }

  public getAllExtDocuments(): Promise<IGetExtContentFuncResponse> {
    this.log.logInfo("getAllExtDocuments()");

    // TODO : do some clever caching
    const rowLimit: number = 500; // TODO : we need to get many pages with this etc..
    const baseUri: string = this.props.host.pageContext.webAbsoluteUrl + "/_api/search/query";

    const extContentFql: string = "" + this.props.managedProperyName + ":#ext#";

    let scopeFql: string = "";
    if (this.props.scope === SPScope.SiteCollection) {
      scopeFql = " Path:" + this.props.host.pageContext.webAbsoluteUrl; // TODO: get site collection url
    }
    else if (this.props.scope === SPScope.Site) {
      scopeFql = " Path:" + this.props.host.pageContext.webAbsoluteUrl;
    }
    else if (this.props.scope === SPScope.Tenant) {
      // do nothing
    }
    else {
      this.log.logError("Unsupported scope: " + this.props.scope);
      return null;
    }

    // "MY" should represent things I have created or edited or shared.
    let modeFql: string = "";
    if (this.props.mode === Mode.AllExtSharedDocuments || this.props.mode === Mode.MyExtSharedDocuments) {
      modeFql = " IsDocument=1"; // TODO: Do something better than this
    }
    else if (this.props.mode === Mode.AllExtSharedContainers || this.props.mode === Mode.MyExtSharedContainers) {
      modeFql = " IsDocument=0"; // TODO: Do something better than this
    }
    else {
      this.log.logError("Unsupported mode: " + this.props.mode);
      return null;
    }

    const queryText: string = "querytext='" + extContentFql + scopeFql + modeFql + "'";
    const selectProps: string = "selectproperties='Title,ServerRedirectedURL,FileExtension'";
    const finalUri: string = baseUri + "?" + queryText + "&" + selectProps + "&rowlimit=" + rowLimit;

    this.log.logInfo("Submitting request to " + finalUri);
    return this.props.host.httpClient.get(finalUri)
      .then((r1: Response) => {
        return r1.json().then((r) => {
          this.log.logInfo("Recieved response from " + finalUri);
          const finalResults: IGetExtContentFuncResponse = this._transformSearchResults(r, this.props.noResultsString);
          this.timeStamp = finalResults.timeStamp;
          return finalResults;
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
      controlMode: shouldShowMessage ? ControlMode.Message : ControlMode.Content,
      message: message,
      timeStamp: (new Date()).getTime()
    };
  };
}
