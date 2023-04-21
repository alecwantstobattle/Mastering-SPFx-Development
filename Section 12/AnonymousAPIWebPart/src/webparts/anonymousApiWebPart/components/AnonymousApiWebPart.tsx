import * as React from 'react';
import styles from './AnonymousApiWebPart.module.scss';
import { IAnonymousApiWebPartProps } from './IAnonymousApiWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { AnonymousApiWebPartState } from './AnonymousApiWebPartState';

export default class AnonymousApiWebPart extends React.Component<
  IAnonymousApiWebPartProps,
  AnonymousApiWebPartState
> {
  public constructor(
    props: IAnonymousApiWebPartProps,
    state: AnonymousApiWebPartState
  ) {
    super(props);

    this.state = {
      id: null,
      name: null,
      username: null,
      email: null,
      address: null,
      phone: null,
      website: null,
      company: null,
    };
  }

  private getUserDetails(): Promise<any> {
    let url = this.props.apiUrl + '/' + this.props.userId;

    return this.props.context.httpClient
      .get(url, HttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse;
      }) as Promise<any>;
  }

  public render(): React.ReactElement<IAnonymousApiWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.anonymousApiWebPart} ${
          hasTeamsContext ? styles.teams : ''
        }`}>
        <div>
          <strong>ID: </strong>
          {this.state.id}
        </div>
        <br />
        <div>
          <strong>User Name: </strong>
          {this.state.username}
        </div>
        <br />
        <div>
          <strong>Name: </strong>
          {this.state.name}
        </div>
        <br />
        <div>
          <strong>Address: </strong>
          {this.state.address}
        </div>
        <br />
        <div>
          <strong>Email: </strong>
          {this.state.email}
        </div>
        <br />
        <div>
          <strong>Phone: </strong>
          {this.state.phone}
        </div>
        <br />
        <div>
          <strong>WebSite: </strong>
          {this.state.website}
        </div>
        <br />
        <div>
          <strong>Company: </strong>
          {this.state.company}
        </div>
        <br />
      </section>
    );
  }
}
