import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CultureInfoDetailsWebPart.module.scss';
import * as strings from 'CultureInfoDetailsWebPartStrings';

export interface ICultureInfoDetailsWebPartProps {
  description: string;
}

export default class CultureInfoDetailsWebPart extends BaseClientSideWebPart <ICultureInfoDetailsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.cultureInfoDetails }">
    <div class="${ styles.container }">
      <div class="${ styles.row }">
        <div class="${ styles.column }">
          <span class="${ styles.title }">Welcome to SharePoint!</span>
  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
    <p class="${ styles.description }">${escape(this.properties.description)}</p>


    <ul>
    <li><strong>current Culture Name</strong>: ${escape(this.context.pageContext.cultureInfo.currentCultureName)}</li>
    <li><strong>current UI Culture Name</strong>: ${escape(this.context.pageContext.cultureInfo.currentUICultureName)}</li>
    <li><strong>isRightToLeft?</strong>: ${this.context.pageContext.cultureInfo.isRightToLeft}</li>
  </ul>


      <a href="https://aka.ms/spfx" class="${ styles.button }">
        <span class="${ styles.label }">Learn more</span>
          </a>
          </div>
          </div>
          </div>
          </div>`;
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
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              })
            ]
          }
        ]
      }
    ]
  };
}
}
