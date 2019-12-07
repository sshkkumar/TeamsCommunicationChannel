import * as React from "react";
import styles from "./TeamsCommunicationChannel.module.scss";
import { ITeamsCommunicationChannelProps } from "./ITeamsCommunicationChannelProps";
import { escape } from "@microsoft/sp-lodash-subset";

import User from "./User";
import { Tab, Tabs, TabList, TabPanel } from "react-tabs";
import "react-tabs/style/react-tabs.css";
import DocumentsPendingApproval from "./DocumentsPendingApproval";
import "bootstrap/dist/css/bootstrap.css";

export default class TeamsCommunicationChannel extends React.Component<
  ITeamsCommunicationChannelProps,
  {}
> {
  constructor(props: ITeamsCommunicationChannelProps) {
    super(props);
  }
  public render(): React.ReactElement<ITeamsCommunicationChannelProps> {
    return (
      <div>
        <Tabs>
          <TabList>
            <Tab>User</Tab>
            <Tab>Approver</Tab>
          </TabList>

          <TabPanel>
            <User
              teamsContext={this.props.teamsContext}
              spContext={this.props.spContext}
              formDigest={this.props.formDigest}
            />
          </TabPanel>
          <TabPanel>
            <DocumentsPendingApproval
              teamsContext={this.props.teamsContext}
              spContext={this.props.spContext}
              formDigest={this.props.formDigest}
            />
          </TabPanel>
        </Tabs>
      </div>
    );
  }
}
