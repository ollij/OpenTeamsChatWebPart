import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'OpenTeamsChatWebPartStrings';
import OpenTeamsChat from './components/OpenTeamsChat';
import { IOpenTeamsChatProps } from './components/IOpenTeamsChatProps';
import { PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";


export interface IOpenTeamsChatWebPartProps {
  usePageInfo: boolean;
  caption: string;
  people: IPropertyFieldGroupOrPerson[];
  topic: string;
  message: string;
}

export default class OpenTeamsChatWebPart extends BaseClientSideWebPart<IOpenTeamsChatWebPartProps> {

  public render(): void {
    let users: string = '';
    if (this.properties.usePageInfo) {
      // TODO: use pagecontext information and query the page metadata (owners, url, title)
    } else {
      if (this.properties.people != undefined) {
        for(var i: number = 0; i < this.properties.people.length; i++) {
          users = users + this.properties.people[i].email;
          if (i+1 !== this.properties.people.length) {
            users = users + ',';
          }
        }
      }
    }
    
    const element: React.ReactElement<IOpenTeamsChatProps > = React.createElement(
      OpenTeamsChat,
      {
        caption: this.properties.caption,
        users: users,
        topic: this.properties.topic,
        message: this.properties.message
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onUsersChanged(propertyPath: string, oldValue: IPropertyFieldGroupOrPerson[], newValue: IPropertyFieldGroupOrPerson[]): void {
    
  };

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the web part'
          },
          groups: [
            {
              groupName: 'Web part properties', 
              groupFields: [
                PropertyPaneTextField('caption', {
                  label: 'Caption of the chat link'
                }),
                PropertyFieldPeoplePicker('people', {
                  label: 'Select users to chat with',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onUsersChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                }),
                PropertyPaneTextField('topic', {
                  label: 'Topic of conversation'
                }),
                PropertyPaneTextField('message', {
                  label: 'Initiating message of the conversation'
                }),
                PropertyPaneToggle('usePageInfo', {
                  label: 'Use page info to determine users'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
