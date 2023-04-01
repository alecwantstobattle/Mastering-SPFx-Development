import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MultiGroupsWebPartWebPart.module.scss';
import * as strings from 'MultiGroupsWebPartWebPartStrings';

export interface IMultiGroupsWebPartWebPartProps {
  description: string;

  productName: string;
  isCertified: boolean;
}

export default class MultiGroupsWebPartWebPart extends BaseClientSideWebPart<IMultiGroupsWebPartWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.multiGroupsWebPart} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ''
    }">
    <p>${escape(this.properties.description)}</p>
    <p>${escape(this.properties.productName)}</p>
    <p>${this.properties.isCertified}</p>

    </section>`;
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
              groupName: 'First Group',
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: 'Product Name 1',
                }),
              ],
            },
            {
              groupName: 'Second Group',
              groupFields: [
                PropertyPaneToggle('isCertified', {
                  label: 'Is Certified 1?',
                }),
              ],
            },
          ],
          displayGroupsAsAccordion: true,
        },
      ],
    };
  }
}
