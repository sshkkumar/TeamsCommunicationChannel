import { IWebPartContext } from "@microsoft/sp-webpart-base";

export interface ITeamsCommunicationChannelProps {
  description: string;
  teamsContext: any;
  spContext: IWebPartContext;
  formDigest: any;
}
