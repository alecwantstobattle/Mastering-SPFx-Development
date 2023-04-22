import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphApiUserProfileWebPartWebPart.module.scss';
import * as strings from 'GraphApiUserProfileWebPartWebPartStrings';

import { MSGraphClient } from '@microsoft/sp-http';

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphApiUserProfileWebPartWebPartProps {
  description: string;
}

export default class GraphApiUserProfileWebPartWebPart extends BaseClientSideWebPart<IGraphApiUserProfileWebPartWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.context.msGraphClientFactory
      .getClient()
      .then((graphclient: MSGraphClient): void => {
        graphclient
          .api('/me')
          .get((error, user: MicrosoftGraph.User, rawResponse?: any) => {
            this.domElement.innerHTML = `
          <div>
            <p class="${styles.links}">Display Name: ${user.displayName}</p>
            <p class="${styles.links}">Given Name: ${user.givenName}</p>
            <p class="${styles.links}">Surname: ${user.surname}</p>
            <p class="${styles.links}">Email ID: ${user.mail}</p>
            <p class="${styles.links}">Mobile Phone: ${user.mobilePhone}</p>   
          </div>`;
          });
      });
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
