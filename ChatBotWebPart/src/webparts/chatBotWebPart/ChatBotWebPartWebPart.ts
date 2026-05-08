import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import ChatBotWebPart from './components/ChatBotWebPart';
import { IChatBotWebPartProps } from './components/IChatBotWebPartProps';

export interface IChatBotWebPartWebPartProps {
  agentUrl: string;
  tokenEndpointUrl: string;
  title: string;
  height: number;
  width: number;
}

export default class ChatBotWebPartWebPart extends BaseClientSideWebPart<IChatBotWebPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IChatBotWebPartProps> = React.createElement(
      ChatBotWebPart,
      {
        context: this.context,
        agentUrl: this.properties.agentUrl,
        tokenEndpointUrl: this.properties.tokenEndpointUrl,
        title: this.properties.title || 'Chat',
        height: this.properties.height || 600,
        width: this.properties.width || 400
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
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
            description: 'Configure the Chat Bot web part'
          },
          groups: [
            {
              groupName: 'Bot Settings',
              groupFields: [
                PropertyPaneTextField('agentUrl', {
                  label: 'Agent URL',
                  description: 'Published Copilot Studio/Azure agent web URL (iframe mode)',
                  multiline: false
                }),
                PropertyPaneTextField('tokenEndpointUrl', {
                  label: 'Token Endpoint URL',
                  description: 'Optional fallback: Direct Line token endpoint URL',
                  multiline: false
                }),
                PropertyPaneTextField('title', {
                  label: 'Widget Title'
                }),
                PropertyPaneSlider('height', {
                  label: 'Height (px)',
                  min: 200,
                  max: 1200,
                  step: 50,
                  value: 600
                }),
                PropertyPaneSlider('width', {
                  label: 'Width (px)',
                  min: 200,
                  max: 1200,
                  step: 50,
                  value: 400
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
