import * as React from 'react';
import styles from './ProviderWebPart.module.scss';
import { IProviderWebPartProps } from './IProviderWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IProviderWebPartState } from './IDepartmentState';
import { IDepartment } from './IDepartment';

import {
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  DetailsRowCheck,
  Selection,
} from 'office-ui-fabric-react';

import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse,
} from '@microsoft/sp-http';

let _departmentListColumns = [
  {
    key: 'ID',
    name: 'ID',
    fieldName: 'ID',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'Title',
    name: 'Title',
    fieldName: 'Title',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
];

export default class ProviderWebPart extends React.Component<
  IProviderWebPartProps,
  IProviderWebPartState
> {
  private _selection: Selection;

  private _onItemsSelectionChanged = () => {
    this.props.onDepartmentSelected(
      this._selection.getSelection()[0] as IDepartment
    );

    this.setState({
      DepartmentListItem: this._selection.getSelection()[0] as IDepartment,
    });
  };
  constructor(props: IProviderWebPartProps, state: IProviderWebPartState) {
    super(props);

    this.state = {
      status: 'Ready',
      DepartmentListItems: [],
      DepartmentListItem: {
        Id: 0,
        Title: '',
      },
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged,
    });
  }

  private _getListItems(): Promise<IDepartment[]> {
    const url: string =
      this.props.siteUrl + "/_api/web/lists/getbytitle('Departments')/items";
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        return json.value;
      }) as Promise<IDepartment[]>;
  }

  public bindDetailsList(message: string): void {
    this._getListItems().then((listItems) => {
      this.setState({ DepartmentListItems: listItems, status: message });
    });
  }

  public componentDidMount(): void {
    this.bindDetailsList('All Records have been loaded Successfully');
  }

  public render(): React.ReactElement<IProviderWebPartProps> {
    return (
      <div>
        <DetailsList
          items={this.state.DepartmentListItems}
          columns={_departmentListColumns}
          setKey="Id"
          checkboxVisibility={CheckboxVisibility.always}
          selectionMode={SelectionMode.single}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          compact={true}
          selection={this._selection}
        />
      </div>
    );
  }
}
