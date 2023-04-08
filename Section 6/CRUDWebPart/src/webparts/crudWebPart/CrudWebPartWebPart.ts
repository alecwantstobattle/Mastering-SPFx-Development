import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudWebPartWebPart.module.scss';
import * as strings from 'CrudWebPartWebPartStrings';

import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse,
} from '@microsoft/sp-http';
import { ISoftwareListItem } from './ISoftwareListItem';

export interface ICrudWebPartWebPartProps {
  description: string;
}

export default class CrudWebPartWebPart extends BaseClientSideWebPart<ICrudWebPartWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.crudWebPart} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ''
    }">
      <div>
        <table border='5' bgcolor='aqua'>
          <tr>
            <td>Please Enter Software ID </td>
            <td><input type='text' id='txtID' />
            <td><input type='submit' id='btnRead' value='Read Details' />
          </td>
          </tr>
          <tr>
            <td>Software Title</td>
            <td><input type='text' id='txtSoftwareTitle' />
          </tr>
          <tr>
            <td>Software Name</td>
            <td><input type='text' id='txtSoftwareName' />
          </tr>
          <tr>
            <td>Software Vendor</td>
            <td>
              <select id="ddlSoftwareVendor">
                <option value="Microsoft">Microsoft</option>
                <option value="Sun">Sun</option>
                <option value="Oracle">Oracle</option>
                <option value="Google">Google</option>
              </select>  
            </td>      
          </tr>
          <tr>
            <td>Software Version</td>
            <td><input type='text' id='txtSoftwareVersion' />
          </tr>
          <tr>
            <td>Software Description</td>
            <td><textarea rows='5' cols='40' id='txtSoftwareDescription'> </textarea> </td>
          </tr>
          <tr>
            <td colspan='2' align='center'>
              <input type='submit'  value='Insert Item' id='btnSubmit' />
              <input type='submit'  value='Update' id='btnUpdate' />
              <input type='submit'  value='Delete' id='btnDelete' />      
            </td>
          </tr>
        </table>
      </div>
      <div id="divStatus"/>
    </section>`;
    this._bindEvents();
  }

  private _bindEvents(): void {
    this.domElement
      .querySelector('#btnSubmit')
      .addEventListener('click', () => {
        this.addListItem();
      });

    this.domElement.querySelector('#btnRead').addEventListener('click', () => {
      this.readListItem();
    });

    this.domElement
      .querySelector('#btnUpdate')
      .addEventListener('click', () => {
        this.updateListItem();
      });

    this.domElement
      .querySelector('#btnDelete')
      .addEventListener('click', () => {
        this.deleteListItem();
      });
  }

  private addListItem(): void {
    var softwaretitle = document.getElementById('txtSoftwareTitle')['value'];
    var softwarename = document.getElementById('txtSoftwareName')['value'];
    var softwareversion =
      document.getElementById('txtSoftwareVersion')['value'];
    var softwarevendor = document.getElementById('ddlSoftwareVendor')['value'];
    var softwareDescription = document.getElementById('txtSoftwareDescription')[
      'value'
    ];

    const siteurl: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SoftwareCatalog')/items";

    const itemBody: any = {
      Title: softwaretitle,
      SoftwareVendor: softwarevendor,
      SoftwareDescription: softwareDescription,
      SoftwareName: softwarename,
      SoftwareVersion: softwareversion,
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(itemBody),
    };

    this.context.spHttpClient
      .post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          let statusmessage: Element =
            this.domElement.querySelector('#divStatus');
          statusmessage.innerHTML = 'List Item has been created successfully.';
          this.clear();
        } else {
          let statusmessage: Element =
            this.domElement.querySelector('#divStatus');
          statusmessage.innerHTML =
            'An error has occured i.e.  ' +
            response.status +
            ' - ' +
            response.statusText;
        }
      });
  }

  private clear(): void {
    document.getElementById('txtSoftwareTitle')['value'] = '';
    document.getElementById('ddlSoftwareVendor')['value'] = 'Microsoft';
    document.getElementById('txtSoftwareDescription')['value'] = '';
    document.getElementById('txtSoftwareVersion')['value'] = '';
    document.getElementById('txtSoftwareName')['value'] = '';
  }

  private readListItem(): void {
    let id: string = document.getElementById('txtID')['value'];
    this._getListItemByID(id)
      .then((listItem) => {
        document.getElementById('txtSoftwareTitle')['value'] = listItem.Title;
        document.getElementById('ddlSoftwareVendor')['value'] =
          listItem.SoftwareVendor;
        document.getElementById('txtSoftwareDescription')['value'] =
          listItem.SoftwareDescription;
        document.getElementById('txtSoftwareName')['value'] =
          listItem.SoftwareName;
        document.getElementById('txtSoftwareVersion')['value'] =
          listItem.SoftwareVersion;
      })
      .catch((error) => {
        let message: Element = this.domElement.querySelector('#divStatus');
        message.innerHTML = 'Read: Could not fetch details.. ' + error.message;
      });
  }

  private _getListItemByID(id: string): Promise<ISoftwareListItem> {
    const url: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SoftwareCatalog')/items?$filter=Id eq " +
      id;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((listItems: any) => {
        const untypedItem: any = listItems.value[0];
        const listItem: ISoftwareListItem = untypedItem as ISoftwareListItem;
        return listItem;
      }) as Promise<ISoftwareListItem>;
  }

  private updateListItem(): void {
    var title = document.getElementById('txtSoftwareTitle')['value'];
    var softwareVendor = document.getElementById('ddlSoftwareVendor')['value'];
    var softwareDescription = document.getElementById('txtSoftwareDescription')[
      'value'
    ];
    var softwareName = document.getElementById('txtSoftwareName')['value'];
    var softwareVersion =
      document.getElementById('txtSoftwareVersion')['value'];

    let id: string = document.getElementById('txtID')['value'];

    const url: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SoftwareCatalog')/items(" +
      id +
      ')';
    const itemBody: any = {
      Title: title,
      SoftwareVendor: softwareVendor,
      SoftwareDescription: softwareDescription,
      SoftwareName: softwareName,
      SoftwareVersion: softwareVersion,
    };
    const headers: any = {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*',
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: headers,
      body: JSON.stringify(itemBody),
    };

    this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = 'List Item has been updated successfully.';
        } else {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML =
            'List Item updation failed. ' +
            response.status +
            ' - ' +
            response.statusText;
        }
      });
  }

  private deleteListItem(): void {
    let id: string = document.getElementById('txtID')['value'];
    const url: string =
      this.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('SoftwareCatalog')/items(" +
      id +
      ')';
    const headers: any = { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' };

    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: headers,
    };

    this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML =
            'Delete: List Item has been deleted successfully.';
        } else {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML =
            'Failed to Delete...' +
            response.status +
            ' - ' +
            response.statusText;
        }
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
