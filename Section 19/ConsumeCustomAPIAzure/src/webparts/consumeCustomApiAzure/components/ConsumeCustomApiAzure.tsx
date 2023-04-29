import * as React from 'react';
import styles from './ConsumeCustomApiAzure.module.scss';
import { IConsumeCustomApiAzureProps } from './IConsumeCustomApiAzureProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { ICustomer } from './ICustomer';
import { ICustomerState } from './ICustomerState';

import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';

import {
  TextField,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  ITextFieldStyles,
  IDropdownStyles,
  DetailsRowCheck,
  Selection,
} from 'office-ui-fabric-react';

let _customerListColumns = [
  {
    key: 'id',
    name: 'id',
    fieldName: 'id',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'joined',
    name: 'joined',
    fieldName: 'joined',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'name',
    name: 'name',
    fieldName: 'name',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'city',
    name: 'city',
    fieldName: 'city',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'orderTotal',
    name: 'orderTotal',
    fieldName: 'orderTotal',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
];

const textFieldStyles: Partial<ITextFieldStyles> = {
  fieldGroup: { width: 300 },
};
const narrowTextFieldStyles: Partial<ITextFieldStyles> = {
  fieldGroup: { width: 100 },
};

const siteUrl: string = 'https://beeslyfunctionapp.azurewebsites.net';

export default class ConsumeCustomApiAzure extends React.Component<
  IConsumeCustomApiAzureProps,
  ICustomerState
> {
  private _selection: Selection;

  private _onItemsSelectionChanged = () => {
    this.setState({
      CustomerListItem: this._selection.getSelection()[0] as ICustomer,
    });
  };

  constructor(props: IConsumeCustomApiAzureProps, state: ICustomerState) {
    super(props);
    this.state = {
      status: 'Ready',
      CustomerListItems: [],
      CustomerListItem: {
        id: '',
        joined: '',
        name: '',
        city: '',
        orderTotal: 0,
      },
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged,
    });
  }

  public _getListItems(): Promise<ICustomer[]> {
    return new Promise<ICustomer[]>((resolve, reject) => {
      const azurefunctionendpoint: string = `${siteUrl}/api/customerdetails`;

      this.props.context.httpClient
        .get(azurefunctionendpoint, HttpClient.configurations.v1)
        .then((response: HttpClientResponse) => {
          return response.json();
        })
        .then((response: ICustomer[]) => {
          resolve(response);
        });
    });
  }

  public bindDetailsList(message: string): void {
    this._getListItems().then((listItems) => {
      console.log(listItems);
      this.setState({ CustomerListItems: listItems, status: message });
    });
  }

  public componentDidMount(): void {
    this.bindDetailsList('All Records have been loaded Successfully');
  }

  public render(): React.ReactElement<IConsumeCustomApiAzureProps> {
    return (
      <div className={styles.consumeCustomApiAzure}>
        <DetailsList
          items={this.state.CustomerListItems}
          columns={_customerListColumns}
          setKey="id"
          checkboxVisibility={CheckboxVisibility.onHover}
          selectionMode={SelectionMode.single}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          compact={false}
          selection={this._selection}
        />
      </div>
    );
  }
}
