import * as React from 'react';
import styles from './CrudReactWebPart.module.scss';
import { ICrudReactWebPartProps } from './ICrudReactWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICrudReactWebPartState } from './ICrudReactWebPartState';

import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse,
} from '@microsoft/sp-http';

import {
  TextField,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  Dropdown,
  IDropdown,
  IDropdownOption,
  ITextFieldStyles,
  IDropdownStyles,
  DetailsRowCheck,
  Selection,
} from 'office-ui-fabric-react';
import { ISoftwareListItem } from './ISoftwareListItem';

let _softwareListColumns = [
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
    key: 'SoftwareName',
    name: 'SoftwareName',
    fieldName: 'SoftwareName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'SoftwareVendor',
    name: 'SoftwareVendor',
    fieldName: 'SoftwareVendor',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'SoftwareVersion',
    name: 'SoftwareVersion',
    fieldName: 'SoftwareVersion',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'SoftwareDescription',
    name: 'SoftwareDescription',
    fieldName: 'SoftwareDescription',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true,
  },
];

export default class CrudReactWebPart extends React.Component<
  ICrudReactWebPartProps,
  ICrudReactWebPartState
> {
  constructor(props: ICrudReactWebPartProps, state: ICrudReactWebPartState) {
    super(props);

    this.state = {
      status: 'Ready',
      SoftwareListItems: [],
      SoftwareListItem: {
        Id: 0,
        Title: '',
        SoftwareName: '',
        SoftwareDescription: '',
        SoftwareVendor: 'Select an option',
        SoftwareVersion: '',
      },
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged,
    });
  }

  public render(): React.ReactElement<ICrudReactWebPartProps> {
    const dropdownRef = React.createRef<IDropdown>();

    return (
      <div className={styles.crudWithReact}>
        <TextField
          label="ID"
          required={false}
          value={this.state.SoftwareListItem.Id.toString()}
          styles={textFieldStyles}
          onChange={(e, v) => {
            this.state.SoftwareListItem.Id = +v;
          }}
        />
        <TextField
          label="Software Title"
          required={true}
          value={this.state.SoftwareListItem.Title}
          styles={textFieldStyles}
          onChange={(e, v) => {
            this.state.SoftwareListItem.Title = v;
          }}
        />
        <TextField
          label="Software Name"
          required={true}
          value={this.state.SoftwareListItem.SoftwareName}
          styles={textFieldStyles}
          onChange={(e, v) => {
            this.state.SoftwareListItem.SoftwareName = v;
          }}
        />
        <TextField
          label="Software Description"
          required={true}
          value={this.state.SoftwareListItem.SoftwareDescription}
          styles={textFieldStyles}
          onChange={(e, v) => {
            this.state.SoftwareListItem.SoftwareDescription = v;
          }}
        />
        <TextField
          label="Software Version"
          required={true}
          value={this.state.SoftwareListItem.SoftwareVersion}
          styles={textFieldStyles}
          onChange={(e, v) => {
            this.state.SoftwareListItem.SoftwareVersion = v;
          }}
        />
        <Dropdown
          componentRef={dropdownRef}
          placeholder="Select an option"
          label="Software Vendor"
          options={[
            { key: 'Microsoft', text: 'Microsoft' },
            { key: 'Sun', text: 'Sun' },
            { key: 'Oracle', text: 'Oracle' },
            { key: 'Google', text: 'Google' },
          ]}
          defaultSelectedKey={this.state.SoftwareListItem.SoftwareVendor}
          required
          styles={narrowDropdownStyles}
          onChange={(e, v) => {
            this.state.SoftwareListItem.SoftwareVendor = v.text;
          }}
        />

        <p className={styles.title}>
          <PrimaryButton text="Add" title="Add" onClick={this.btnAdd_click} />

          <PrimaryButton text="Update" onClick={this.btnUpdate_click} />

          <PrimaryButton text="Delete" onClick={this.btnDelete_click} />
        </p>

        <div id="divStatus">{this.state.status}</div>

        <div>
          <DetailsList
            items={this.state.SoftwareListItems}
            columns={_softwareListColumns}
            setKey="Id"
            checkboxVisibility={CheckboxVisibility.onHover}
            selectionMode={SelectionMode.single}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            compact={true}
            selection={this._selection}
          />
        </div>
      </div>
    );
  }
}
