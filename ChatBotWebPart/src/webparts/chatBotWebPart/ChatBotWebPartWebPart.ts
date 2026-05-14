/**
 * ChatBotWebPartWebPart.ts
 *
 * SPFx Web Part entry point for the ChatBot web part.
 * Responsible for mounting the ChatBotWebPart React component into the
 * SharePoint page, managing its lifecycle, and exposing a property pane
 * for editor configuration.
 *
 * Property pane controls:
 *  - agentUrl          : iframe URL for a published Copilot Studio / Azure agent.
 *  - tokenEndpointUrl  : Direct Line token endpoint (fallback when agentUrl is blank).
 *  - title             : Widget header label.
 *  - height / width    : Pixel dimensions of the embedded chat widget.
 */
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

/**
 * Strongly-typed property bag for this web part.
 * Values are persisted in the SharePoint page model and surfaced to the
 * property pane for editor configuration.
 */
export interface IChatBotWebPartWebPartProps {
  /** Published Copilot Studio / Azure Bot Service iframe URL. */
  agentUrl: string;

  /** Direct Line token endpoint URL (used when agentUrl is not set). */
  tokenEndpointUrl: string;

  /** Display label shown in the chat widget header bar. */
  title: string;

  /** Height of the chat widget in pixels (property pane range 200–1200). */
  height: number;

  /** Width of the chat widget in pixels (property pane range 200–1200). */
  width: number;
}

/**
 * ChatBotWebPartWebPart
 *
 * Extends BaseClientSideWebPart to integrate the ChatBot React component
 * into the SharePoint page lifecycle.  Handles rendering, disposal, and
 * the property pane panel shown to page editors.
 */

export default class ChatBotWebPartWebPart extends BaseClientSideWebPart<IChatBotWebPartWebPartProps> {

  /**
   * render()
   *
   * Called by the SPFx framework whenever the web part needs to be (re-)rendered.
   * Creates the root React element with resolved props and mounts it into domElement.
   *
   * Default values ensure a usable initial experience when the property pane has
   * not yet been configured (e.g., immediately after adding the web part to a page).
   */
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

  /**
   * onInit()
   *
   * SPFx lifecycle hook called once before the first render.
   * No asynchronous initialisation is required by this web part.
   */
  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  /**
   * onDispose()
   *
   * SPFx lifecycle hook called when the web part is removed from the page
   * (e.g., page edit mode exit, or web part deletion).
   * Unmounts the React component tree to release event listeners and prevent
   * memory leaks.
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * dataVersion
   *
   * Semantic version of the serialised property bag (IChatBotWebPartWebPartProps).
   * Increment this value when making breaking changes to the property schema so
   * the framework can detect and migrate stale stored data.
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * getPropertyPaneConfiguration()
   *
   * Defines the structure and controls rendered inside the property pane panel
   * when a site editor clicks "Edit web part".
   *
   * Bot Settings group:
   *   agentUrl          — TextField: iframe URL for Copilot Studio / Azure agent.
   *   tokenEndpointUrl  — TextField: fallback Direct Line token endpoint URL.
   *   title             — TextField: widget header label.
   *   height / width    — Sliders: pixel dimensions (range 200–1200, step 50).
   */
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
