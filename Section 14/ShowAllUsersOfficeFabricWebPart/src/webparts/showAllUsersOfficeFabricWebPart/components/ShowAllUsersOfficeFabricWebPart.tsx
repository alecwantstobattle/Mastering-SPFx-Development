import * as React from 'react';
import styles from './ShowAllUsersOfficeFabricWebPart.module.scss';
import { IShowAllUsersOfficeFabricWebPartProps } from './IShowAllUsersOfficeFabricWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IShowAllUsersOfficeFabricState } from './IShowAllUsersOfficeFabricState';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  TextField,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
} from 'office-ui-fabric-react';

import * as strings from 'ShowAllUsersOfficeFabricWebPartWebPartStrings';
import { IUser } from './IUser';

// Configure the columns for the DetailsList component
let _usersListColumns = [
  {
    key: 'displayName',
    name: 'Display name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true,
  },
  {
    key: 'givenName',
    name: 'Given Name',
    fieldName: 'givenName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'surName',
    name: 'SurName',
    fieldName: 'surname',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'mail',
    name: 'Mail',
    fieldName: 'mail',
    minWidth: 150,
    maxWidth: 150,
    isResizable: true,
  },
  {
    key: 'mobilePhone',
    name: 'mobile Phone',
    fieldName: 'mobilePhone',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: 'userPrincipalName',
    name: 'User Principal Name',
    fieldName: 'userPrincipalName',
    minWidth: 200,
    maxWidth: 200,
    isResizable: true,
  },
];

export default class ShowAllUsersOfficeFabricWebPart extends React.Component<
  IShowAllUsersOfficeFabricWebPartProps,
  IShowAllUsersOfficeFabricState
> {
  constructor(
    props: IShowAllUsersOfficeFabricWebPartProps,
    state: IShowAllUsersOfficeFabricState
  ) {
    super(props);

    this.state = {
      users: [],
      searchFor: '',
    };
  }

  public componentDidMount(): void {
    this.fetchUserDetails();
  }

  public _search(): void {
    this.fetchUserDetails();
  }

  private _onSearchForChanged(e: any, newValue: string): void {
    this.setState({
      searchFor: newValue,
    });
  }

  private _getSearchForErrorMessage(value: string): string {
    return value == null || value.length == 0 || value.indexOf(' ') < 0
      ? ''
      : `${strings.SearchForValidationErrorMessage}`;
  }

  public fetchUserDetails(): void {
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('users')
          .version('v1.0')
          .select('*')
          .filter(`startswith(givenname,'${escape(this.state.searchFor)}')`)
          .get((error: any, response, rawResponse?: any) => {
            if (error) {
              console.error('Message is : ' + error);
              return;
            }

            // Prepare the output array
            var allUsers: Array<IUser> = new Array<IUser>();

            // Map the JSON response to the output array
            response.value.map((item: IUser) => {
              allUsers.push({
                displayName: item.displayName,
                givenName: item.givenName,
                surname: item.surname,
                mail: item.mail,
                mobilePhone: item.mobilePhone,
                userPrincipalName: item.userPrincipalName,
              });
            });

            this.setState({ users: allUsers });
          });
      });
  }

  public render(): React.ReactElement<IShowAllUsersOfficeFabricWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div className={styles.showAllUsersOfficeFabricWebPart}>
        <TextField
          label={strings.SearchFor}
          required={true}
          value={this.state.searchFor}
          onChange={(e, v) => this._onSearchForChanged(e, v)}
          onGetErrorMessage={this._getSearchForErrorMessage}
        />
        <p>
          <PrimaryButton
            text="Search"
            title="Search"
            onClick={() => this._search()}
          />
        </p>
        {this.state.users != null && this.state.users.length > 0 ? (
          <p>
            <DetailsList
              items={this.state.users}
              columns={_usersListColumns}
              setKey="set"
              checkboxVisibility={CheckboxVisibility.onHover}
              selectionMode={SelectionMode.single}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              compact={true}
            />
          </p>
        ) : null}
      </div>
    );
  }
}
