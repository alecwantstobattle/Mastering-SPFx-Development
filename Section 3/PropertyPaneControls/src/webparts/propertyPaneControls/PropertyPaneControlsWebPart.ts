import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneLink,
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

  currentTime: Date;
  isCertified: boolean;
  rating: number;
  processorType: string;
  invoiceFileType: string;
  newProcessorType: string;
  discountCoupon: boolean;
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

  protected get disableReactivePropertyChanges(): boolean {
    return false;
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
        <tr>
          <td>Is Certified?</td>
          <td>${this.properties.isCertified}</td>
        </tr>
        </tr>
        <tr>
          <td>Rating</td>
          <td>${this.properties.rating}</td>
        </tr>
        <tr>
          <td>Processor Type</td>
          <td>${this.properties.processorType}</td>
        </tr>
        <tr>
            <td>Invoice File Type</td>
            <td>${this.properties.invoiceFileType}</td>
        </tr>
        <tr>
           <td>New Processor Type</td>
           <td>${this.properties.newProcessorType}</td>
        </tr>
        <tr>
          <td>Do u have a discount coupon?</td>
          <td>${this.properties.discountCoupon}</td>
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
                PropertyPaneToggle('isCertified', {
                  key: 'isCertified',
                  label: 'Is it Certified?',
                  onText: 'ISI Certified!',
                  offText: 'Not an ISI Certified Product',
                }),
                PropertyPaneSlider('rating', {
                  label: 'Select Your Rating',
                  min: 1,
                  max: 10,
                  step: 1,
                  showValue: true,
                  value: 1,
                }),
                PropertyPaneChoiceGroup('processorType', {
                  label: 'Choices',
                  options: [
                    { key: 'I5', text: 'Intel I5' },
                    { key: 'I7', text: 'Intel I7', checked: true },
                    { key: 'I9', text: 'Intel I9' },
                  ],
                }),
                PropertyPaneChoiceGroup('invoiceFileType', {
                  label: 'Select Invoice File type:',
                  options: [
                    {
                      key: 'MSWord',
                      text: 'MSWord',
                      imageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/docx_32x1.png',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/docx_32x1.png',
                    },
                    {
                      key: 'MSExcel',
                      text: 'MSExcel',
                      imageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/xlsx_32x1.png',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/xlsx_32x1.png',
                    },
                    {
                      key: 'MSPowerPoint',
                      text: 'MSPowerPoint',
                      imageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/pptx_32x1.png',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/pptx_32x1.png',
                    },
                    {
                      key: 'OneNote',
                      text: 'OneNote',
                      imageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/one_32x1.png',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc:
                        'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/one_32x1.png',
                    },
                  ],
                }),
                PropertyPaneDropdown('newProcessorType', {
                  label: 'New Processor Type',
                  options: [
                    { key: 'I5', text: 'Intel I5' },
                    { key: 'I7', text: 'Intel I7' },
                    { key: 'I9', text: 'Intel I9' },
                  ],
                  selectedKey: 'I7',
                }),
                PropertyPaneCheckbox('discountCoupon', {
                  text: 'Do You have a Discount Coupon?',
                  checked: false,
                  disabled: false,
                }),
                PropertyPaneLink('', {
                  href: 'https://www.amazon.in',
                  text: 'Buy Intel Processor from the best Seller',
                  target: '_blank',
                  popupWindowProps: {
                    height: 500,
                    width: 500,
                    positionWindowPosition: 2,
                    title: 'Amazon',
                  },
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
