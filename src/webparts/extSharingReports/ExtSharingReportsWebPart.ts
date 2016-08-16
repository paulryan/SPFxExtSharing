import {
  DisplayMode
} from "@ms/sp-client-base";

import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  IWebPartData,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from "@ms/sp-client-platform";

import * as React from "react";
import * as ReactDom from "react-dom";

import strings from "./loc/Strings.resx";
import ExtSharingReports from "./ExtSharingReports";

import {
  SPScope,
  Mode,
  DisplayType,
  IExtSharingReportsProps,
  IExtSharingReportsWebPartProps
} from "./ExtSharingReportsInterfaces";

import ExtContentFetcher from "./ExtContentFetcher";

export default class ExtSharingReportsWebPart extends BaseClientSideWebPart<IExtSharingReportsWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(mode: DisplayMode, data?: IWebPartData): void {
    let element: React.ReactElement<IExtSharingReportsProps> = null;
    if (this.properties.displayType === DisplayType.Table) {
      element = React.createElement(ExtSharingReports, {
        store: new ExtContentFetcher({
          host: this.host,
          scope: this.properties.scope,
          mode: this.properties.mode,
          noResultsString: this.properties.noResultsString
        })
      });
    }
    else {
      throw new Error("This Display Type is not yet implemented: " + this.properties.displayType);
    }
    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Core",
              groupFields: [
                PropertyPaneDropdown("scope", {
                  label: "Where should we look for externally shared content?",
                  options: [
                    { key: SPScope.Tenant, text: "Across the entire tenancy" },
                    { key: SPScope.SiteCollection, text: "Only within this site collection" },
                    { key: SPScope.Site, text: "Only in this site (and in child sites)" }
                  ]
                }),
                PropertyPaneDropdown("mode", {
                  label: "What type content do you want to see?",
                  options: [
                    { key: Mode.AllExtSharedDocuments, text: "All externally shared documents" },
                    { key: Mode.MyExtSharedDocuments, text: "Documents which I have shared externally" },
                    { key: Mode.AllExtSharedContainers, text: "All externally shared sites, libraries, and folders" },
                    { key: Mode.MyExtSharedContainers, text: "Sites, libraries, and folders which I have shared externally" },
                  ]
                }),
                PropertyPaneDropdown("displayType", {
                  label: "How do you want the results rendered?",
                  options: [
                    { key: DisplayType.Table, text: "As a table" },
                    { key: DisplayType.Tree, text: "Hierarchically" },
                    { key: DisplayType.BySite, text: "Charted by site" },
                    { key: DisplayType.ByUser, text: "Charted by user" },
                    { key: DisplayType.OverTime, text: "Charted over time" },
                  ]
                }),
              ]
            },
            {
              groupName: "Other",
              groupFields: [
                PropertyPaneTextField("noResultsMessage", {
                  label: "No results message"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
