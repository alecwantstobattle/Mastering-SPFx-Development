import * as React from 'react';
import styles from './AnonymousApiWebPart.module.scss';
import { IAnonymousApiWebPartProps } from './IAnonymousApiWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class AnonymousApiWebPart extends React.Component<
  IAnonymousApiWebPartProps,
  {}
> {
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
          {this.props.id}
        </div>
        <br />
        <div>
          <strong>User Name: </strong>
          {this.props.username}
        </div>
        <br />
        <div>
          <strong>Name: </strong>
          {this.props.name}
        </div>
        <br />
        <div>
          <strong>Address: </strong>
          {this.props.address}
        </div>
        <br />
        <div>
          <strong>Email: </strong>
          {this.props.email}
        </div>
        <br />
        <div>
          <strong>Phone: </strong>
          {this.props.phone}
        </div>
        <br />
        <div>
          <strong>WebSite: </strong>
          {this.props.website}
        </div>
        <br />
        <div>
          <strong>Company: </strong>
          {this.props.company}
        </div>
        <br />
      </section>
    );
  }
}
