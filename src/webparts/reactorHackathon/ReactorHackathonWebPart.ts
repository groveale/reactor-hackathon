import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ReactorHackathonWebPartStrings';
import ReactorHackathon from './components/ReactorHackathon';
import { IReactorHackathonProps } from './components/IReactorHackathonProps';


import GraphPersona from './components/Persona/GraphPersona';
import { IGraphPersonaProps } from './components/Persona/IGraphPersonaProps';

import { MSGraphClient } from '@microsoft/sp-http';


export interface IReactorHackathonWebPartProps {
  description: string;
}

export default class ReactorHackathonWebPart extends BaseClientSideWebPart<IReactorHackathonWebPartProps> {

  public render(): void {
    // const element: React.ReactElement<IReactorHackathonProps > = React.createElement(
    //   ReactorHackathon,
    //   {
    //     description: this.properties.description
    //   }
    // );

    // ReactDom.render(element, this.domElement);


    this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient): void => {
        const element: React.ReactElement<IGraphPersonaProps> = React.createElement(
          GraphPersona,
          {
            spfxContext: this.context,
            graphClient: client
          }
        );

        ReactDom.render(element, this.domElement);
      });
    
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
