import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PropertyPaneControlsWebPart.module.scss';
import * as strings from 'PropertyPaneControlsWebPartStrings';

export interface IPropertyPaneControlsWebPartProps {
  description: string;

  productName: string;
  productDescription: string;
  productCost: number;
  quantity: number;
  billAmount: number;
  discount: number;
  netBillAmount: number;
}

export default class PropertyPaneControlsWebPart extends BaseClientSideWebPart<IPropertyPaneControlsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return new Promise<void>((resolve, _reject) => {
      this.properties.productName = 'Mouse';
      this.properties.productDescription = 'Mouse Description';
      this.properties.quantity = 500;
      this.properties.productCost = 300;

      resolve(undefined);
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.propertyPaneControls} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ''
    }">
    <table>   
    <tr>
    <td>Product Name</td>
    <td>${this.properties.productName}</td>
    </tr>
    <tr>
    <td>Description</td>
    <td>${this.properties.productDescription}</td>
    </tr>
    <tr>
    <td>Product Cost</td>
    <td>${this.properties.productCost}</td>
    </tr>
    <tr>
    <td>Product Quantity</td>
    <td>${this.properties.quantity}</td>
    </tr>
    <tr>
          <td>Bill Amount</td>
          <td>${(this.properties.billAmount =
            this.properties.productCost * this.properties.quantity)} </td>
    </tr>
          <tr>
          <td>Discount</td>
          <td>${(this.properties.discount =
            (this.properties.billAmount * 10) / 100)}</td>
          </tr>

          <tr>
          <td>Net Bill Amount</td>
          <td>${(this.properties.netBillAmount =
            this.properties.billAmount - this.properties.discount)}</td>
          </tr>          
     </tr>
    </table>
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
          groups: [
            {
              groupName: 'Product Details',
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: 'Product Name',
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: 'Please enter product name',
                  description: 'Name property field',
                }),

                PropertyPaneTextField('productDescription', {
                  label: 'Product Description',
                  multiline: true,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: 'Please enter Product Description',
                  description: 'Name property field',
                }),

                PropertyPaneTextField('productCost', {
                  label: 'Product Cost',
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: 'Please enter product Cost',
                  description: 'Number property field',
                }),

                PropertyPaneTextField('quantity', {
                  label: 'Product Quantity',
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: 'Please enter product Quantity',
                  description: 'Number property field',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
