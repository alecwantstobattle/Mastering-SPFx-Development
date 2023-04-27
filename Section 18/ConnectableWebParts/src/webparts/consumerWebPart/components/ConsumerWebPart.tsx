import * as React from 'react';
import styles from './ConsumerWebPart.module.scss';
import { IConsumerWebPartProps } from './IConsumerWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IEmployee } from './IEmployee';
import { IConsumerWebPartState } from './IEmployeeState';

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

import { IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';

let _employeeListColumns = [
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
  {
    key: 'DeptTitle',
    name: 'DeptTitle',
    fieldName: 'DeptTitleId',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'Designation',
    name: 'Designation',
    fieldName: 'Designation',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
];

export default class ConsumerWebPart extends React.Component<
  IConsumerWebPartProps,
  IConsumerWebPartState
> {
  constructor(props: IConsumerWebPartProps, state: IConsumerWebPartState) {
    super(props);

    this.state = {
      status: 'Ready',
      EmployeeListItems: [],
      EmployeeListItem: {
        Id: 0,
        Title: '',
        DeptTitle: '',
        Designation: '',
      },
      DeptTitleId: '',
    };
  }

  private _getListItems(): Promise<IEmployee[]> {
    const url: string =
      this.props.siteUrl +
      "/_api/web/lists/getbytitle('Employee')/items?$filter=DeptTitleId eq " +
      this.props.DeptTitleId.tryGetValue();
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        return json.value;
      }) as Promise<IEmployee[]>;
  }

  public bindDetailsList(message: string): void {
    this._getListItems().then((listItems) => {
      console.log(listItems);
      this.setState({
        EmployeeListItems: listItems,
        status: message,
        DeptTitleId: this.props.DeptTitleId.tryGetValue().toString(),
      });
    });
  }

  public componentDidMount(): void {
    // this.bindDetailsList("All Records have been loaded Successfully");
  }

  public render(): React.ReactElement<IConsumerWebPartProps> {
    if (this.state.DeptTitleId != this.props.DeptTitleId.tryGetValue()) {
      this.bindDetailsList('All Records have been loaded Successfully');
    }

    return (
      <div className={styles.consumerWebPart}>
        <div>
          <h1>
            Selected Department is : {this.props.DeptTitleId.tryGetValue()}
          </h1>
        </div>
        <DetailsList
          items={this.state.EmployeeListItems}
          columns={_employeeListColumns}
          setKey="Id"
          checkboxVisibility={CheckboxVisibility.always}
          selectionMode={SelectionMode.single}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          compact={true}
        />
      </div>
    );
  }
}
