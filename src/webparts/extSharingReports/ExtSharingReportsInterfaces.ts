import {
  IWebPartHost
} from '@ms/sp-client-platform';

export enum SPScope {
  Tenant = 1,
  SiteCollection = 2,
  Site = 3
}

export enum Mode {
  AllExtSharedDocuments = 1,
  MyExtSharedDocuments = 2,
  AllExtSharedContainers = 3, // may not be possible
  MyExtSharedContainers = 4, // may not be possible
}

export enum DisplayType {
  Table = 1,
  Tree = 2,
  BySite = 3,
  ByUser = 4,
  OverTime = 5
}

export enum SecurableObjectType {
  Document = 1,
  Library = 2,
  Web = 3,
  Site = 4
}

export interface IExtContentFetcherProps {
  host: IWebPartHost;
  scope: SPScope;
  mode: Mode;
  noResultsString: string;
}

export interface IGetExtContentFuncResponse {
  extContent: ISecurableObject[];
  shouldShowMessage: boolean;
  message: string;
}

export interface ISecurableObject {
  Title: string;
  URL: string;
  Type: SecurableObjectType;
  FileExtension: string;
  LastModifiedTime: string;
  SharedWith: string;
  SharedBy: string;
}

export interface IGetExtContentFunc {
    (props: IExtContentFetcherProps): Promise<IGetExtContentFuncResponse>;
}

export interface ISecurableObjectStore {
  getAllExtDocuments: IGetExtContentFunc;
  // getMyExtDocuments: IGetExtContentFunc;
  // getAllExtNonDocuments: IGetExtContentFunc;
  // getMyExtNonDocuments: IGetExtContentFunc;
}

export interface IExtSharingReportsWebPartProps {
  scope: SPScope;
  mode: Mode;
  displayType: DisplayType;
  noResultsString: string;
}

export interface IExtSharingReportsProps {
  store: ISecurableObjectStore;
}
