import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'OboLoginWebPartStrings';
import OboLogin from './components/OboLogin';
import { IOboLoginProps } from './components/IOboLoginProps';

export interface IOboLoginWebPartProps {
  apiUrl: string;
}

export default class OboLoginWebPart extends BaseClientSideWebPart<IOboLoginWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _apiClient: AadHttpClient;

  public render(): void {
    const element: React.ReactElement<IOboLoginProps> = React.createElement(
      OboLogin,
      {
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        apiUrl: this.properties.apiUrl,
        apiClient: this._apiClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    if (!this.properties.apiUrl) {
      this.properties.apiUrl = 'http://localhost:3000'
    }

    this._apiClient = await this.context.aadHttpClientFactory.getClient('api://a7607af0-ef1f-42e2-89f8-6f455217262e');
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
                PropertyPaneTextField('apiUrl', {
                  label: 'Backend API URL'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
