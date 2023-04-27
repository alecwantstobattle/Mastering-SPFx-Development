import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ProviderWebPartWebPartStrings';
import ProviderWebPart from './components/ProviderWebPart';
import { IProviderWebPartProps } from './components/IProviderWebPartProps';

import {
  IDynamicDataPropertyDefinition,
  IDynamicDataCallables,
} from '@microsoft/sp-dynamic-data';
import { IDepartment } from './components/IDepartment';

export interface IProviderWebPartWebPartProps {
  description: string;
}

export default class ProviderWebPartWebPart
  extends BaseClientSideWebPart<IProviderWebPartWebPartProps>
  implements IDynamicDataCallables
{
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _selectedDepartment: IDepartment;

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    this.context.dynamicDataSourceManager.initializeSource(this);

    return Promise.resolve();
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'id',
        title: 'Selected Department ID',
      },
    ];
  }

  public getPropertyValue(propertyId: string): string | IDepartment {
    switch (propertyId) {
      case 'id':
        return this._selectedDepartment.Id.toString();
    }

    throw new Error('Invalid property ID');
  }

  private handleDepartmentChangeSelected = (department: IDepartment): void => {
    this._selectedDepartment = department;
    this.context.dynamicDataSourceManager.notifyPropertyChanged('id');
    console.log('End Of Handle Event : ' + department.Id + department.Title);
  };

  public render(): void {
    const element: React.ReactElement<IProviderWebPartProps> =
      React.createElement(ProviderWebPart, {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        onDepartmentSelected: this.handleDepartmentChangeSelected,
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
