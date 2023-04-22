import * as React from 'react';
import styles from './ShowAllUsersOfficeFabricWebPart.module.scss';
import { IShowAllUsersOfficeFabricWebPartProps } from './IShowAllUsersOfficeFabricWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IShowAllUsersOfficeFabricState } from './IShowAllUsersOfficeFabricState';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  TextField,
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
} from 'office-ui-fabric-react';

import * as strings from 'ShowAllUsersOfficeFabricWebPartWebPartStrings';

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

  public render(): React.ReactElement<IShowAllUsersOfficeFabricWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.showAllUsersOfficeFabricWebPart} ${
          hasTeamsContext ? styles.teams : ''
        }`}>
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require('../assets/welcome-dark.png')
                : require('../assets/welcome-light.png')
            }
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part property value: <strong>{escape(description)}</strong>
          </div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for
            Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way
            to extend Microsoft 365 with automatic Single Sign On, automatic
            hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li>
              <a href="https://aka.ms/spfx" target="_blank">
                SharePoint Framework Overview
              </a>
            </li>
            <li>
              <a href="https://aka.ms/spfx-yeoman-graph" target="_blank">
                Use Microsoft Graph in your solution
              </a>
            </li>
            <li>
              <a href="https://aka.ms/spfx-yeoman-teams" target="_blank">
                Build for Microsoft Teams using SharePoint Framework
              </a>
            </li>
            <li>
              <a href="https://aka.ms/spfx-yeoman-viva" target="_blank">
                Build for Microsoft Viva Connections using SharePoint Framework
              </a>
            </li>
            <li>
              <a href="https://aka.ms/spfx-yeoman-store" target="_blank">
                Publish SharePoint Framework applications to the marketplace
              </a>
            </li>
            <li>
              <a href="https://aka.ms/spfx-yeoman-api" target="_blank">
                SharePoint Framework API reference
              </a>
            </li>
            <li>
              <a href="https://aka.ms/m365pnp" target="_blank">
                Microsoft 365 Developer Community
              </a>
            </li>
          </ul>
        </div>
      </section>
    );
  }
}
