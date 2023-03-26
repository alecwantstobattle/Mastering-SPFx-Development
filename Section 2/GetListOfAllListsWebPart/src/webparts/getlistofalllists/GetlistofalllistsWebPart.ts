import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetlistofalllistsWebPart.module.scss';
import * as strings from 'GetlistofalllistsWebPartStrings';
import { ISPList } from '../ISPList';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IGetlistofalllistsWebPartProps {
  description: string;
}

export default class GetlistofalllistsWebPart extends BaseClientSideWebPart<IGetlistofalllistsWebPartProps> {

  public render(): void {  
    let listItems: string = "";
    this._getSharePointLists().then(lists => {
      lists.forEach(list => {
        listItems += `
    <div>
    <img src='${list.ImageUrl}'/>&nbsp;${list.Title}<br/>
    ID: ${list.Id}<br/>
    </div><hr/>`;
      });
      this.domElement.innerHTML = `
      <div class="${ styles.getlistofalllists}">        
            <h1>All SharePoint Lists:</h1><div>${listItems}</div>
            </div>
          
        
      </div>`;
    });
  } 

    private _getSharePointLists(): Promise<ISPList[]> {
      const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
      return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then(response => {
          return response.json();
        })
        .then(json => {
          return json.value;
        }) as Promise<ISPList[]>;
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
