import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FileTeamsSaverAddinWebPartStrings';
import FileTeamsSaverAddin from './components/FileTeamsSaverAddin';
import { IFileTeamsSaverAddinProps } from './components/FileTeamsSaverAddin';

export interface IFileTeamsSaverAddinWebPartProps {
  description: string;
}

export default class FileTeamsSaverAddinWebPart extends BaseClientSideWebPart <IFileTeamsSaverAddinWebPartProps> {

  public render(): void {
    if(this.context.sdks.office){
      console.log("Executing in Outlook!");    
      console.log(this.context.sdks.office.context.mailbox.item.itemId);
    }
    

    const element: React.ReactElement<IFileTeamsSaverAddinProps> = React.createElement(
      FileTeamsSaverAddin,
      {
        serviceScope: this.context.serviceScope,
        mailId: this.context.sdks.office.context.mailbox.item.itemId
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
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
