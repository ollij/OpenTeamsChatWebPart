import * as React from 'react';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export interface ITeamsChatLinkProps {
    users: string;
    topic: string;
    message: string;
    caption: string;
  }
  
  export function TeamsChatLink(props: ITeamsChatLinkProps) {
    let link: string = 'https://teams.microsoft.com/l/chat/0/0?users='
                        +props.users;
    if (props.topic !== null && props.topic != undefined 
        && props.topic != '') {
      link = link + '&topicName='+props.topic;
    }
    if (props.message != null && props.message != undefined 
        && props.message != '') {
      link = link + '&message='+props.message;
    }
    return (
      <Link disabled={!props.users == null || props.users == undefined 
            || props.users === ''} 
              href={link} target='_blank'>
                <Icon iconName='TeamsLogo' /> {props.caption}
      </Link> 
    );
  }










