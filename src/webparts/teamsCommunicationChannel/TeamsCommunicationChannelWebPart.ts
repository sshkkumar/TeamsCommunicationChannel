import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";

import * as strings from "TeamsCommunicationChannelWebPartStrings";
import TeamsCommunicationChannel from "./components/TeamsCommunicationChannel";
import { ITeamsCommunicationChannelProps } from "./components/ITeamsCommunicationChannelProps";
import * as microsoftTeams from "@microsoft/teams-js";
import { sp, ItemAddResult } from "@pnp/sp";

import {
  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration,
  ISPHttpClientOptions
} from "@microsoft/sp-http";

export interface ITeamsCommunicationChannelWebPartProps {
  description: string;
}

export default class TeamsCommunicationChannelWebPart extends BaseClientSideWebPart<
  ITeamsCommunicationChannelWebPartProps
> {
  private teamsContext: microsoftTeams.Context;
  private digest: any;

  public getDigest() {
    let url = `https://m365x846523.sharepoint.com/sites/Dep1/_api/contextinfo`; //https://sshkkumar.sharepoint.com/_api/contextinfo

    let spOpts: ISPHttpClientOptions = {
      headers: new Headers(),
      method: "POST",
      mode: "cors"
      // headers: {
      //   // Accept: "application/json",
      //   // "Content-Type": "application/json" //,
      //   //"X-RequestDigest": this.props.digest
      // },
    };

    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
    // return $.ajax({
    //   url: `https://m365x846523.sharepoint.com/sites/Dep1/_api/contextinfo`, // ${teamsContext.teamSiteUrl}
    //   method: "POST",
    //   headers: {
    //     Accept: "application/json; odata=verbose",
    //     crossDomain: "true",
    //     credentials: "include"
    //   },
    //   xhrFields: { withCredentials: true }
    // });
  }

  protected onInit(): Promise<void> {
    return new Promise<void>(
      (resolve: () => void, reject: (error: any) => void): void => {
        //const digestCache: IDigestCache = this.context.serviceScope.consume(DigestCache.serviceKey);
        this.getDigest().then((digest: any): void => {
          // use the digest here
          this.digest = digest.FormDigestValue;

          sp.setup({
            sp: {
              baseUrl: this.context.pageContext.web.absoluteUrl
            }
          });

          if (this.context.microsoftTeams) {
            //retVal = new Promise((resolve, reject) => {
            this.context.microsoftTeams.getContext(context => {
              this.teamsContext = context;
              //resolve();
            });
            //});
          }

          resolve();
        });
      }
    );

    // let retVal: Promise<any> = Promise.resolve();

    // this.getDigest().then(data => {
    //   this.digest = data.FormDigestValue;
    // });
    // sp.setup({
    //   sp: {
    //     baseUrl: this.context.pageContext.web.absoluteUrl
    //   }
    // });

    // if (this.context.microsoftTeams) {
    //   retVal = new Promise((resolve, reject) => {
    //     this.context.microsoftTeams.getContext(context => {
    //       this.teamsContext = context;
    //       resolve();
    //     });
    //   });
    // }
    // return retVal;
  }

  public render(): void {
    const element: React.ReactElement<
      ITeamsCommunicationChannelProps
    > = React.createElement(TeamsCommunicationChannel, {
      description: this.properties.description,
      teamsContext: this.teamsContext,
      spContext: this.context,
      formDigest: this.digest
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
