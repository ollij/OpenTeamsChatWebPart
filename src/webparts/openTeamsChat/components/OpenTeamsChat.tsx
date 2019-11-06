import * as React from 'react';
import styles from './OpenTeamsChat.module.scss';
import { IOpenTeamsChatProps } from './IOpenTeamsChatProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ITeamsChatLinkProps, TeamsChatLink } from './TeamsChatLink'

export default class OpenTeamsChat extends React.Component<IOpenTeamsChatProps, {}> {
  public render(): React.ReactElement<IOpenTeamsChatProps> {
    console.log('OpenTeamsChat.render');
    return (
        <TeamsChatLink users={this.props.users} 
          topic={this.props.topic} message={this.props.message} caption={this.props.caption}/>
    );
  }
}
