import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListCreationWebPartWebPart.module.scss';
import * as strings from 'ListCreationWebPartWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from '@microsoft/sp-http';

export interface IListCreationWebPartWebPartProps {
  description: string;
}

export default class ListCreationWebPartWebPart extends BaseClientSideWebPart<IListCreationWebPartWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.listCreationWebPart} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ''
    }">
      <h3>Creating a New List Dynamically</h3>
      <br/>
      <p>Please fill out the below details to create a new list programatically </p>
      <br/>   
      New List Name: 
      <br/>
      <input type='text' id='txtNewListName' />
      <br/>
      New List Description: 
      <br/>
      <input type='text' id='txtNewListDescription'/>
      <br/>
      <input type="button" id="btnCreateNewList" value="Create a New List"/>
      <br/>
    </section>`;

    this.bindEvents();
  }

  private bindEvents(): void {
    this.domElement
      .querySelector('#btnCreateNewList')
      .addEventListener('click', () => {
        this.createNewList();
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

  private createNewList(): void {
    var newListName = document.getElementById('txtNewListName')['value'];

    var newListDescription = document.getElementById('txtNewListDescription')[
      'value'
    ];

    const listUrl: string =
      this.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/GetByTitle('" +
      newListName +
      "')";

    this.context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          alert('A List already does exist with this name.');
          return;
        }
        if (response.status === 404) {
          const url: string =
            this.context.pageContext.web.absoluteUrl + '/_api/web/lists';

          const listDefinition: any = {
            Title: newListName,
            Description: newListDescription,
            AllowContentTypes: true,
            BaseTemplate: 105,
            ContentTypesEnabled: true,
          };
          const spHttpClientOptions: ISPHttpClientOptions = {
            body: JSON.stringify(listDefinition),
          };
          this.context.spHttpClient
            .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then((response: SPHttpClientResponse) => {
              if (response.status === 201) {
                alert('A new List has been created successfully.');
              } else {
                alert(
                  'Error Message  ' +
                    response.status +
                    ' - ' +
                    response.statusText
                );
              }
            });
        } else {
          alert(
            'Error Message. ' + response.status + ' ' + response.statusText
          );
        }
      });
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
