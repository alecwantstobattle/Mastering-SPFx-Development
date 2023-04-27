import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ConsumerWebPartWebPartStrings';
import ConsumerWebPart from './components/ConsumerWebPart';
import { IConsumerWebPartProps } from './components/IConsumerWebPartProps';

import { DynamicProperty } from '@microsoft/sp-component-base';

import {
  DynamicDataSharedDepth,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
  IPropertyPaneConditionalGroup,
  IWebPartPropertiesMetadata,
} from '@microsoft/sp-webpart-base';

export interface IConsumerWebPartWebPartProps {
  description: string;
  DeptTitleId: DynamicProperty<string>;
}

export default class ConsumerWebPartWebPart extends BaseClientSideWebPart<IConsumerWebPartWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IConsumerWebPartProps> =
      React.createElement(ConsumerWebPart, {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        DeptTitleId: this.properties.DeptTitleId,
      });

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty(
      '--linkHovered',
      semanticColors.linkHovered
    );
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      DeptTitleId: { dynamicPropertyType: 'string' },
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'Select Department ID',
                  fields: [
                    PropertyPaneDynamicField('DeptTitleId', {
                      label: 'Department ID',
                    }),
                  ],
                  sharedConfiguration: {
                    depth: DynamicDataSharedDepth.Property,
                    source: {
                      sourcesLabel:
                        'Select the web part containing the list of Departments',
                    },
                  },
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
