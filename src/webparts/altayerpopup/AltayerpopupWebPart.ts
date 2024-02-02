import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'AltayerpopupWebPartStrings';
import Altayerpopup from './components/Altayerpopup';
import { IAltayerpopupProps } from './components/IAltayerpopupProps';

export interface IAltayerpopupWebPartProps {
  description: string;
  listName:string;
  resPonselistName:string;
  consentTerms:string;
  neverShowText:string;
}

export default class AltayerpopupWebPart extends BaseClientSideWebPart<IAltayerpopupWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IAltayerpopupProps> = React.createElement(
      Altayerpopup,
      {
        description: this.properties.description,
        userDisplayName: this.context.pageContext.user.displayName,
        listName: this.properties.listName ?? 'Consent Message',
        responseListName :this.properties.resPonselistName ?? 'Consent Response',
        webpartContext: this.context,
        consentTerms:this.properties.consentTerms ?? 'I Acknowledge that I have Read And Understand The Terms.',
        neverShowText:this.properties.neverShowText ?? `don't show any more`
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onInit(): Promise<void> {
    //create list
    return Promise.resolve();
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
                PropertyPaneTextField('consentTerms', {
                  label: 'Enter Description 1'
                }) ,
                PropertyPaneTextField('neverShowText', {
                  label:  'Enter Description 2'
                })     
              ]
            }
          ]
        }
      ]
    };
  }
}
