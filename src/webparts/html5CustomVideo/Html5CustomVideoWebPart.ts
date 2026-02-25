import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'Html5CustomVideoWebPartStrings';
import Html5CustomVideo from './components/Html5CustomVideo';
import { IHtml5CustomVideoProps } from './components/IHtml5CustomVideoProps';

export interface IHtml5CustomVideoWebPartProps {
  description: string;
  videoUrl: string;
  videoTitle: string;
  autoplay: boolean;
  controls: boolean;
  preload: string;
  width: string;
  height: string;
}

export default class Html5CustomVideoWebPart extends BaseClientSideWebPart<IHtml5CustomVideoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IHtml5CustomVideoProps> = React.createElement(
      Html5CustomVideo,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        videoUrl: this.properties.videoUrl || '',
        videoTitle: this.properties.videoTitle || '',
        autoplay: this.properties.autoplay || false,
        controls: this.properties.controls !== false,
        preload: this.properties.preload || 'auto',
        width: this.properties.width || '100%',
        height: this.properties.height || '540px'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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
              groupName: strings.VideoSourceGroupName,
              groupFields: [
                PropertyPaneTextField('videoUrl', {
                  label: strings.VideoUrlLabel,
                  description: strings.VideoUrlDescription,
                  placeholder: 'https://example.com/video.mp4'
                }),
                PropertyPaneTextField('videoTitle', {
                  label: strings.VideoTitleLabel
                })
              ]
            },
            {
              groupName: strings.PlayerSettingsGroupName,
              groupFields: [
                PropertyPaneToggle('autoplay', {
                  label: strings.AutoplayLabel
                }),
                PropertyPaneToggle('controls', {
                  label: strings.ControlsLabel
                }),
                PropertyPaneDropdown('preload', {
                  label: strings.PreloadLabel,
                  options: [
                    { key: 'auto', text: 'Auto' },
                    { key: 'metadata', text: 'Metadata' },
                    { key: 'none', text: 'None' }
                  ]
                })
              ]
            },
            {
              groupName: strings.DimensionsGroupName,
              groupFields: [
                PropertyPaneTextField('width', {
                  label: strings.WidthLabel,
                  description: strings.WidthDescription,
                  placeholder: '100%'
                }),
                PropertyPaneTextField('height', {
                  label: strings.HeightLabel,
                  description: strings.HeightDescription,
                  placeholder: '540px'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
