import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MicrosoftTeamsTabWebPart.module.scss';
import * as strings from 'MicrosoftTeamsTabWebPartStrings';
import * as microsoftteams from '@microsoft/teams-js';

export interface IMicrosoftTeamsTabWebPartProps {
  description: string;
}

export default class MicrosoftTeamsTabWebPart extends BaseClientSideWebPart<IMicrosoftTeamsTabWebPartProps> {
  private _teamsContext: microsoftTeams.Context;

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<any> {
    this._environmentMessage = this._getEnvironmentMessage();

    let retVal: Promise<void> = Promise.resolve();

    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext((context) => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {
    let greetingsTitle: string = '';
    let siteOrTabTitle: string = '';

    if (this._teamsContext) {
      greetingsTitle = 'Welcome to Microsoft Teams!';
      siteOrTabTitle = 'Team Name is : ' + this._teamsContext.teamName;
    } else {
      greetingsTitle = 'Welcome to SharePoint!';
      siteOrTabTitle = 'SharePoint site: ' + this.context.pageContext.web.title;
    }

    this.domElement.innerHTML = `<div class="${styles.microsoftTeamsTab}">
      <div class="${styles.container}">
        <div class="${styles.row}">
          <div class="${styles.column}">
          <span class="${styles.title}">Welcome to SharePoint!</span>
          <p class="${
            styles.subTitle
          }">Customize SharePoint experiences using Web Parts.</p>
          <p class="${styles.description}">${escape(
      this.properties.description
    )}</p>
          <p class="${styles.title}">${greetingsTitle}</p>
          <p class="${styles.description}">${siteOrTabTitle}</p>
          <a href="https://aka.ms/spfx" class="${styles.button}">
            <span class="${styles.label}">Learn more</span>
          </a>
          </div>
        </div>
      </div>
    </div>`;
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
