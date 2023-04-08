import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpPnPcrudWebPartWebPart.module.scss';
import * as strings from 'SpPnPcrudWebPartWebPartStrings';

import * as pnp from 'sp-pnp-js';

export interface ISpPnPcrudWebPartWebPartProps {
  description: string;
}

export default class SpPnPcrudWebPartWebPart extends BaseClientSideWebPart<ISpPnPcrudWebPartWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.spPnPcrudWebPart} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ''
    }">
      <div>
        <div>
        <table border='5' bgcolor='aqua'>
          <tr>
            <td>Please Enter Software ID </td>
            <td><input type='text' id='txtID' />
            <td><input type='submit' id='btnRead' value='Read Details' /></td>
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
            <input type='submit'  value='Show All Records' id='btnReadAll' />
          </td>
        </table>
        </div>
      <div id="divStatus"/>
      <h2>Get All List Items</h2>
      <hr/>
      <div id="spListData" />
      </div>
    </section>`;
    this._bindEvents();
  }

  private _bindEvents(): void {
    this.domElement
      .querySelector('#btnSubmit')
      .addEventListener('click', () => {
        this.addListItem();
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

    pnp.sp.web.lists
      .getByTitle('SoftwareCatalog')
      .items.add({
        Title: softwaretitle,
        SoftwareVendor: softwarevendor,
        SoftwareName: softwarename,
        SoftwareVersion: softwareversion,
        SoftwareDescription: softwareDescription,
      })
      .then((r) => {
        alert('success');
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
