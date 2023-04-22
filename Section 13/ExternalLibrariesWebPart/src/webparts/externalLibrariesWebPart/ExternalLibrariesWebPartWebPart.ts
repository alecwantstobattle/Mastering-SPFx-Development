import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ExternalLibrariesWebPartWebPart.module.scss';
import * as strings from 'ExternalLibrariesWebPartWebPartStrings';

import * as $ from 'jquery';
import 'jqueryui';

import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IExternalLibrariesWebPartWebPartProps {
  description: string;
}

export default class ExternalLibrariesWebPartWebPart extends BaseClientSideWebPart<IExternalLibrariesWebPartWebPartProps> {
  constructor() {
    super();

    SPComponentLoader.loadCss(
      'https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/themes/smoothness/jquery-ui.css'
    );
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.externalLibrariesWebPart} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ''
    }">
      <div>
        <div class="accordion">
          <h3>Lesson 14 - ECMAScript Implementation</h3>
          <div>
              <ul>
                <li>Overview of ECMAScript</li>
                <li>using ECMAScript in Application Pages</li>
                <li>Using ECMAScript in Web Parts</li>
                <li>Implementing onSucess Function</li>
                <li>Implementing onFail Function</li>           
              </ul>
          </div>      
          <h3>Lesson 15 - Silverlight with SharePoint</h3>
          <div>
            <ul>
            <li>Overview of Silverlight Implemention</li>
            <li>Using Load Function to load resources</li>
            <li>Adding fields to a custom list using Silverlight Implementation</li>
            <li>Exception handling with Silverlight Implementation</li>
            <li>Cross Domain Policy</li>
            </ul>
          </div>
          <h3>Lesson 16 - Developing Custom Dialogs</h3>
          <div>
            <ul>
              <li>Create a Custom Dialog for Data Entry</li>
              <li>JavaScript and the Client Object Model</li>
              <li>Modal Dialogs</li>
              <li>Creating a Custom Dialog</li>
              <li>Controlling the Client Side Behavior and Visibility of the Dialog</li>
              <li>Adding Server Side Functionality to the Dialog</li>
              <li>Deploying and Testing the Dialog User Control</li>  
            </ul>
          </div>
        </div>
      </div>
    </section>`;

    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: true,
      collapsible: false,
      icons: {
        header: 'ui-icon-circle-arrow-e',
        activeHeader: 'ui-icon-circle-arrow-s',
      },
    };

    ($('.accordion', this.domElement) as any).accordion(accordionOptions);
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
