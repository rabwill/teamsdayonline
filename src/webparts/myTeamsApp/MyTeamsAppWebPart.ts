import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MyTeamsAppWebPartStrings';
import MyTeamsApp from './components/MyTeamsApp';
import { IMyTeamsAppProps } from './components/IMyTeamsAppProps';

export interface IMyTeamsAppWebPartProps {
  description: string;
}

export default class MyTeamsAppWebPart extends BaseClientSideWebPart<IMyTeamsAppWebPartProps> {

  public render(): void {
    if (this.context.sdks.microsoftTeams) {
      // We have teams context for the web part
      const element: React.ReactElement<IMyTeamsAppProps> = React.createElement(
        MyTeamsApp,
        {
          description:"Welcome to Teams",
          teamsContext:this.context.sdks.microsoftTeams
        }
      );
  
      ReactDom.render(element, this.domElement);
    }
    else
    {
      // We are rendered in normal SharePoint context
      const element: React.ReactElement<IMyTeamsAppProps> = React.createElement(
        MyTeamsApp,
        {
          description:"Welcome to SharePoint"
        }
      );
  
      ReactDom.render(element, this.domElement);
    }
 
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
