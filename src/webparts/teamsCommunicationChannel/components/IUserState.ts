import {
  Dropdown,
  IDropdownOption
} from "office-ui-fabric-react/lib/components/Dropdown";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import * as microsoftTeams from "@microsoft/teams-js";

export interface IUserState {
  title: string;
  description: string;
  customerName: string;
  documentType: string;
  dpselectedItem?: { key: string | number | undefined };
  dpselectedItems: IDropdownOption[];
  userManagerIDs: any[];
  required: string;
  onSubmission: boolean;
  filesToUpload: number[];
  items: string[];
  users: string[];
  context: IWebPartContext;
  teamsContext: microsoftTeams.Context;
  message: string;
}
